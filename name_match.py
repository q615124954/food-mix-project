# -*- coding: utf-8 -*-
"""
品名归一化匹配（阶段 2：变体表持久化 + 单位辅助 + 多路召回打分）
思路：多路召回 + 打分 + 置信度分级，替代「精确匹配纠错字典」。
  1) 精确匹配（归一化后）
  2) 变体表（标准名 -> 变体列表，按类别隔离，可持续回流增长，覆盖形近/同音/漏字）
  3) 同义词（别名 -> 标准名）
  4) 拼音全拼 / 首字母（覆盖同音字）
  5) 子串包含（覆盖「生抽」⊂「鲜味生抽」这类别名）
  6) 编辑相似度（difflib，兜底漏字/多字）
  7) 单位辅助（候选单位与 OCR 单位一致时小幅加分）
返回 (best_name, confidence, source, top_candidates)。调用方按置信度分级：
  >= AUTO_CONF 自动纠；SUGGEST_CONF~AUTO_CONF 交 LLM 定夺；< SUGGEST_CONF 不纠（交人工）。
"""
import os
import re
import json
from difflib import SequenceMatcher
from pypinyin import lazy_pinyin

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
VARIANTS_FILE = os.path.join(BASE_DIR, "webapp", "name_variants.json")

# 高置信阈值 / 推荐阈值
AUTO_CONF = 0.9
SUGGEST_CONF = 0.6

# 默认变体表（首次运行时写入 JSON；seed 自 8 张菜类单实测 OCR 错例）
DEFAULT_VARIANTS = {
    "蔬菜": {
        "薄荷叶": ["莴青叶", "莲荷叶"],
        "香菜": ["香蕊", "香茗", "香芯"],
        "金针菇": ["全针蕊", "金针"],
        "娃娃菜": ["搭娃亲"],
        "猪腱子肉": ["惜腱子肉", "腱肉"],
        "亿亩田大米": ["亿百日", "亿亩田"],
        "细面": ["细面条"],
        "蒜头": ["蒜末"],
        "东古鸡精": ["鸡精"],
        "宏斌小米辣": ["小半辣", "小米辣"],
        "去皮蒜米": ["蒜米"],
        "绿豆芽": ["绿豆牙"],
    },
    "肉类": {},
}

# 同义词：别名 -> 标准名（仅当物料表只存在标准名、OCR 出现别名时才有用）
SYNONYMS = {}

_VARIANTS = None


def load_variants():
    """从 JSON 加载变体表；不存在则用默认 seed 创建"""
    global _VARIANTS
    if os.path.exists(VARIANTS_FILE):
        try:
            with open(VARIANTS_FILE, encoding="utf-8") as f:
                _VARIANTS = json.load(f)
            return _VARIANTS
        except (OSError, ValueError):
            pass
    _VARIANTS = json.loads(json.dumps(DEFAULT_VARIANTS))
    save_variants()
    return _VARIANTS


def save_variants():
    os.makedirs(os.path.dirname(VARIANTS_FILE), exist_ok=True)
    with open(VARIANTS_FILE, "w", encoding="utf-8") as f:
        json.dump(_VARIANTS, f, ensure_ascii=False, indent=2)


def variants_for(category):
    if _VARIANTS is None:
        load_variants()
    return _VARIANTS.get(category, {})


def add_variant(category, std, wrong):
    """回流：把 OCR 错名 wrong 归入标准名 std 的变体列表（幂等，持久化）"""
    if not std or not wrong or std == wrong:
        return
    if _VARIANTS is None:
        load_variants()
    std = str(std).strip()
    wrong = str(wrong).strip()
    if not std or not wrong or std == wrong:
        return
    lst = _VARIANTS.setdefault(category, {}).setdefault(std, [])
    if wrong not in lst:
        lst.append(wrong)
        save_variants()


def normalize(name):
    """去括号规格、空白；不删数字（有些品名本身含规格如「20升 食用油」）"""
    s = str(name).strip()
    s = re.sub(r"[（(].*?[)）]", "", s)
    s = re.sub(r"\s+", "", s)
    return s


def _pinyin_full(s):
    return "".join(lazy_pinyin(s))


def _pinyin_initials(s):
    return "".join(p[0] for p in lazy_pinyin(s))


def _strip_tail_num(s):
    """去掉品名末尾粘连的数字（如「盐5」「鸡骨1」），数字是数量不是品名"""
    return re.sub(r"\d+(\.\d+)?$", "", s)


def match(raw, candidates, category="蔬菜", units=None, unit=None):
    """匹配 OCR 品名到候选标准品名。
    candidates: 标准品名列表；units: {品名: 单位}（可选）；unit: OCR 识别的单位（可选）。
    返回 (best, confidence, source, top)，未命中时 best=None。"""
    candidates = [c for c in candidates if c and str(c).strip()]
    n = normalize(raw)
    if not n:
        return None, 0.0, "empty", []

    # 1) 精确匹配（含归一化后、去尾随数字后）
    n_stripped = _strip_tail_num(n)
    for c in candidates:
        if normalize(c) == n:
            return c, 1.0, "exact", [(c, 1.0)]
    for c in candidates:
        if normalize(c) == n_stripped:
            return c, 1.0, "exact_strip", [(c, 1.0)]

    # 2) 变体表精确命中
    for std, variants in variants_for(category).items():
        if std not in candidates:
            continue
        for v in variants:
            if normalize(v) == n:
                return std, 0.95, "variant", [(std, 0.95)]

    # 3) 同义词
    if n in SYNONYMS and SYNONYMS[n] in candidates:
        return SYNONYMS[n], 0.95, "synonym", [(SYNONYMS[n], 0.95)]

    # 4) 多路打分
    scored = []
    for c in candidates:
        cn = normalize(c)
        score = 0.0
        if _pinyin_full(n) == _pinyin_full(cn):
            score = max(score, 0.9)
        if len(n) >= 2 and _pinyin_initials(n) == _pinyin_initials(cn):
            score = max(score, 0.85)
        if n and cn and (n in cn or cn in n):
            score = max(score, 0.8)
        ratio = SequenceMatcher(None, n, cn).ratio()
        score = max(score, ratio)
        # 单位辅助：一致小幅加分（弱信号，不惩罚，避免单位 OCR 错误导致漏召回）
        if unit and units and units.get(c) and units.get(c) == unit:
            score = min(score + 0.05, 1.0)
        scored.append((c, round(min(score, 1.0), 3)))
    scored.sort(key=lambda x: -x[1])
    if not scored:
        return None, 0.0, "none", []
    best, s = scored[0]
    source = "pinyin_exact" if s >= 0.9 else "fuzzy"
    return best, s, source, scored[:5]
