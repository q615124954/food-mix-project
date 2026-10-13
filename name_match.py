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
  >= AUTO_CONF 且 source 不是 needs_review：自动纠；
  SUGGEST_CONF~AUTO_CONF：交 LLM 定夺；< SUGGEST_CONF：不纠（交人工）。
  source==needs_review：弱相似已经摸到自动线（如拼音首字母 + 单位把「香菜」打成「香葱」），
  调用方保留原名并提示人工，不得静默替换。
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
# 弱相似（拼音首字母、子串、一般编辑距离）即使用单位加分，也停在自动线以下。
# 这类高分仍可能把「香菜」误改成「香葱」，调用方应标存疑而不是静默替换。
REVIEW_SOURCE = "needs_review"
WEAK_SCORE_CAP = 0.89

# 默认变体表（首次运行时写入 JSON；seed 自 8 张菜类单实测 OCR 错例）
DEFAULT_VARIANTS = {
    "蔬菜": {
        "薄荷叶": ["莴青叶", "莲荷叶", "薄荷"],
        # 金标 8/11 香蕊 1斤4元、8/12 香蕊 2斤8元、8/13 香芯 3斤12元 都是香葱（8/12 香菜另有一行 2斤15元）
        "香葱": ["香蕊", "香芯"],
        "香菜": [],
        "金针菇": ["全针蕊", "金针"],
        "娃娃菜": ["搭娃亲"],
        "猪腱子肉": ["惜腱子肉", "腱肉"],
        "亿亩田大米": ["亿百日", "亿亩田"],
        "细面": ["细面条"],
        "蒜头": ["蒜末", "大蒜头"],
        "东古鸡精": ["鸡精"],
        "宏斌小米辣": ["小半辣", "小米辣"],
        "去皮蒜米": ["蒜米", "蒜半"],
        "绿豆芽": ["绿豆牙"],
    },
    "肉类": {},
}

# 同义词：别名 -> 标准名（仅当物料表只存在标准名、OCR 出现别名时才有用）
SYNONYMS = {}

# 旧版默认表里写错的变体：加载已有 JSON 时剔除（香蕊/香芯 是香葱，不是香菜）
WRONG_VARIANTS = {"蔬菜": {"香菜": ["香蕊", "香芯"], "猪腱子肉": ["腊肉"]}}

# 只给建议、不自动改的映射：OCR 名本身是真实品名（如「腊肉」），
# 金标里虽然对应「猪腱子肉」，但不能静默替换。命中时置信度压在自动线以下、
# source=needs_review，调用方保留原名、标存疑，并把建议品名给人工确认。
SUGGESTIONS = {"蔬菜": {"腊肉": "猪腱子肉"}}
SUGGEST_SCORE = 0.85


def suggest(raw, candidates, category="蔬菜"):
    """返回建议品名（必须在候选里，且原名本身不在候选里），没有返回 None"""
    n = normalize(raw)
    cands = [c for c in candidates if c and str(c).strip()]
    if any(normalize(c) == n for c in cands):
        return None
    std = (SUGGESTIONS.get(category) or {}).get(n)
    return std if std in cands else None

_VARIANTS = None


def _migrate(variants):
    """已落盘的变体表：剔除旧版错误映射，补上默认表新增的变体（不删用户回流学到的）。
    返回是否有改动。"""
    changed = False
    for cat, table in WRONG_VARIANTS.items():
        for std, wrongs in table.items():
            lst = variants.get(cat, {}).get(std)
            if lst:
                kept = [v for v in lst if v not in wrongs]
                if kept != lst:
                    variants[cat][std] = kept
                    changed = True
    for cat, table in DEFAULT_VARIANTS.items():
        dst = variants.setdefault(cat, {})
        for std, vs in table.items():
            lst = dst.setdefault(std, [])
            for v in vs:
                if v not in lst:
                    lst.append(v)
                    changed = True
    return changed


def load_variants():
    """从 JSON 加载变体表；不存在则用默认 seed 创建"""
    global _VARIANTS
    if os.path.exists(VARIANTS_FILE):
        try:
            with open(VARIANTS_FILE, encoding="utf-8") as f:
                _VARIANTS = json.load(f)
            if _migrate(_VARIANTS):
                save_variants()
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
    返回 (best, confidence, source, top)，未命中时 best=None。
    source 为 needs_review 时置信度已压到自动线以下，调用方应保留原名并提示人工。"""
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

    # 2.5) 建议映射：只提示不自动改
    sg = suggest(n, candidates, category)
    if sg:
        return sg, SUGGEST_SCORE, REVIEW_SOURCE, [(sg, SUGGEST_SCORE)]

    # 3) 同义词
    if n in SYNONYMS and SYNONYMS[n] in candidates:
        return SYNONYMS[n], 0.95, "synonym", [(SYNONYMS[n], 0.95)]

    # 4) 多路打分
    scored = []
    risky_high = set()
    for c in candidates:
        cn = normalize(c)
        score = 0.0
        pinyin_exact_hit = _pinyin_full(n) == _pinyin_full(cn)
        if pinyin_exact_hit:
            score = max(score, 0.9)
        if len(n) >= 2 and _pinyin_initials(n) == _pinyin_initials(cn):
            score = max(score, 0.85)
        if n and cn and (n in cn or cn in n):
            score = max(score, 0.8)
        ratio = SequenceMatcher(None, n, cn).ratio()
        score = max(score, ratio)
        # 全拼相同，或字符几乎相同，才算可以自动替换的强匹配。
        # 仅首字母相同（香菜 / 香葱都是 xc）是高相似、语义不同的弱匹配。
        strong = pinyin_exact_hit or ratio >= AUTO_CONF
        # 单位辅助：一致小幅加分（弱信号，不惩罚，避免单位 OCR 错误导致漏召回）。
        # 弱匹配加分后封顶在自动线以下，避免 0.85+0.05 变成 0.9 后静默替换。
        if unit and units and units.get(c) and units.get(c) == unit:
            score = min(score + 0.05, 1.0)
        rounded = round(min(score, 1.0), 3)
        if not strong and rounded >= AUTO_CONF:
            risky_high.add(c)
            rounded = WEAK_SCORE_CAP
        scored.append((c, rounded))
    scored.sort(key=lambda x: -x[1])
    if not scored:
        return None, 0.0, "none", []
    best, s = scored[0]
    if best in risky_high:
        source = REVIEW_SOURCE
    else:
        source = "pinyin_exact" if s >= AUTO_CONF else "fuzzy"
    return best, s, source, scored[:5]


def correction_action(raw, best, conf, source):
    """决定这次匹配要不要改品名。

    返回 (品名, 提示人工)。
    needs_review：高分仍可能误改（香菜→香葱一类），保留原名并提示人工。
    """
    if source == REVIEW_SOURCE:
        return raw, True
    if best and conf >= AUTO_CONF:
        return best, False
    return raw, False
