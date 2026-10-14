# -*- coding: utf-8 -*-
"""
品名归一化匹配（变体表持久化 + 单位辅助 + 分层召回）

默认路径不调大模型，按层处理：
  1) 固定别名：变体表 / 确认回流的别名（可带物料代码）
  2) 本地模糊：高分自动；中分只给建议，等人确认
  3) 分很低，或员工区 / 已知员工餐：交调用方标员工餐待确认
  4) 只有 2～3 个中分且彼此接近时，才把这几个候选交给 LLM（少见）

海天/老抽（含「消天」这类形近）一律不得自动归到生抽；拿不准就不要替。

返回 (best_name, confidence, source, top_candidates)。
  >= AUTO_CONF 且 source 不是 needs_review：可以自动纠；
  source==needs_review：弱相似摸到自动线（香菜/香葱），保留原名并提示人工。
分层动作见 layered_match()。
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
# 中分候选彼此相差不超过这个值，视为「分数接近」。2～3 个才值得问 LLM。
MID_TIE_GAP = 0.05
# 确认回流写在变体 JSON 顶层，不跟标准名列表混在一个类别里。
# 结构：{类别: {归一化 OCR 名: {std, code, ocr}}}
ALIAS_META = "_alias_meta"

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
# 读失败时的人话原因；空字符串表示没出错。写失败不记在这里。
_LOAD_ERROR = ""


def variants_load_error():
    """变体/别名表读失败时的说明。没加载过会先加载一次。空串表示可用。"""
    if _VARIANTS is None and not _LOAD_ERROR:
        load_variants()
    return _LOAD_ERROR


def _migrate(variants):
    """已落盘的变体表：剔除旧版错误映射，补上默认表新增的变体（不删用户回流学到的）。
    返回是否有改动。"""
    changed = False
    for cat, table in WRONG_VARIANTS.items():
        for std, wrongs in table.items():
            lst = variants.get(cat, {}).get(std)
            if isinstance(lst, list) and lst:
                kept = [v for v in lst if v not in wrongs]
                if kept != lst:
                    variants[cat][std] = kept
                    changed = True
    for cat, table in DEFAULT_VARIANTS.items():
        dst = variants.setdefault(cat, {})
        if not isinstance(dst, dict):
            variants[cat] = {}
            dst = variants[cat]
            changed = True
        for std, vs in table.items():
            lst = dst.setdefault(std, [])
            if not isinstance(lst, list):
                lst = []
                dst[std] = lst
                changed = True
            for v in vs:
                if v not in lst:
                    lst.append(v)
                    changed = True
    return changed


def load_variants():
    """从 JSON 加载变体表。

    文件不存在：用默认 seed 创建（第一次运行，不算失败）。
    文件损坏或读不出来：记一条加载错误，内存里用空表，**不覆盖**原文件，
    免得把已经学过的别名冲掉。调用方看到错误后不得把没对上的品名静默当成员工餐。
    """
    global _VARIANTS, _LOAD_ERROR
    if os.path.exists(VARIANTS_FILE):
        try:
            with open(VARIANTS_FILE, encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                raise ValueError("格式不是 JSON 对象")
            _VARIANTS = data
            _LOAD_ERROR = ""
            if _migrate(_VARIANTS):
                save_variants()
            return _VARIANTS
        except json.JSONDecodeError:
            detail = "文件不是合法的 JSON"
        except OSError as e:
            detail = f"无法读取（{e.strerror or e}）"
        except (ValueError, TypeError):
            detail = "格式不正确"
        _VARIANTS = {"蔬菜": {}, "肉类": {}}
        _LOAD_ERROR = (
            f"品名变体表加载失败：{detail}。"
            "已暂停按品名归员工餐，避免未纠错的品名被当成员工餐。"
            "请修复品名变体表后重新扫描。"
        )
        return _VARIANTS
    _VARIANTS = json.loads(json.dumps(DEFAULT_VARIANTS))
    _LOAD_ERROR = ""
    save_variants()
    return _VARIANTS


def save_variants():
    os.makedirs(os.path.dirname(VARIANTS_FILE), exist_ok=True)
    with open(VARIANTS_FILE, "w", encoding="utf-8") as f:
        json.dump(_VARIANTS, f, ensure_ascii=False, indent=2)


def variants_for(category):
    if _VARIANTS is None:
        load_variants()
    if not isinstance(_VARIANTS, dict) or str(category).startswith("_"):
        return {}
    table = _VARIANTS.get(category, {})
    return table if isinstance(table, dict) else {}


def add_variant(category, std, wrong, code=None):
    """回流：把 OCR 错名 wrong 归入标准名 std 的变体列表（幂等，持久化）。

    code 是物料代码。标准名以后改了字，只要代码还在，下次仍能对上。
    海天/老抽 → 生抽 这种禁止自动归并的配对，这里直接拒绝，不写进别名表。
    变体表没加载成功时不写盘，避免用空表覆盖原来的文件。
    """
    if _LOAD_ERROR:
        return False
    if not std or not wrong or std == wrong:
        return False
    if _VARIANTS is None:
        load_variants()
    if _LOAD_ERROR:
        return False
    std = str(std).strip()
    wrong = str(wrong).strip()
    if not std or not wrong or std == wrong:
        return False
    if forbidden_auto_map(wrong, std):
        return False
    changed = False
    lst = _VARIANTS.setdefault(category, {}).setdefault(std, [])
    if not isinstance(lst, list):
        lst = []
        _VARIANTS[category][std] = lst
    if wrong not in lst:
        lst.append(wrong)
        changed = True
    if code:
        meta = _VARIANTS.setdefault(ALIAS_META, {}).setdefault(category, {})
        rec = {"std": std, "code": str(code), "ocr": wrong}
        key = normalize(wrong)
        if meta.get(key) != rec:
            meta[key] = rec
            changed = True
    if changed:
        save_variants()
    return True


def alias_code(category, ocr_name):
    """确认回流记下的物料代码；没有则空串。"""
    if _VARIANTS is None:
        load_variants()
    if not isinstance(_VARIANTS, dict):
        return ""
    meta = (_VARIANTS.get(ALIAS_META) or {}).get(category) or {}
    if not isinstance(meta, dict):
        return ""
    rec = meta.get(normalize(ocr_name))
    if isinstance(rec, dict):
        return str(rec.get("code") or "")
    return ""


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


def _alias_hit(n, candidates, category, codes):
    """固定别名：先按标准名找变体，再按回流记下的物料代码找（标准名改字也能对上）。

    命中海天/老抽→生抽时当作没对上，交给后面的规则判成拿不准。
    """
    cand_set = set(candidates)
    for std, variants in variants_for(category).items():
        if not isinstance(variants, list) or std not in cand_set:
            continue
        for v in variants:
            if normalize(v) == n and not forbidden_auto_map(n, std):
                return std, 0.95, "variant"
    if not isinstance(_VARIANTS, dict):
        return None
    meta = (_VARIANTS.get(ALIAS_META) or {}).get(category) or {}
    if not isinstance(meta, dict):
        return None
    rec = meta.get(n)
    if not isinstance(rec, dict):
        return None
    code = str(rec.get("code") or "")
    std = rec.get("std")
    if code and codes:
        for c in candidates:
            if str(codes.get(c) or "") == code and not forbidden_auto_map(n, c):
                return c, 0.95, "alias"
    if std in cand_set and not forbidden_auto_map(n, std):
        return std, 0.95, "alias"
    return None


def match(raw, candidates, category="蔬菜", units=None, unit=None, codes=None):
    """匹配 OCR 品名到候选标准品名。
    candidates: 标准品名列表；units: {品名: 单位}（可选）；unit: OCR 识别的单位（可选）。
    codes: {标准品名: 物料代码}（可选），给确认回流的别名用。
    返回 (best, confidence, source, top)，未命中时 best=None。
    source 为 needs_review 时置信度已压到自动线以下，调用方应保留原名并提示人工。
    海天/老抽（含消天）对生抽的分数直接记 0，不会因为模糊分高就被选中。"""
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
        if normalize(c) == n_stripped and not forbidden_auto_map(n, c):
            return c, 1.0, "exact_strip", [(c, 1.0)]

    # 2) 固定别名（变体表 / 物料代码）
    hit = _alias_hit(n, candidates, category, codes)
    if hit:
        std, conf, src = hit
        return std, conf, src, [(std, conf)]

    # 2.5) 建议映射：只提示不自动改
    sg = suggest(n, candidates, category)
    if sg and not forbidden_auto_map(n, sg):
        return sg, SUGGEST_SCORE, REVIEW_SOURCE, [(sg, SUGGEST_SCORE)]

    # 3) 同义词
    if n in SYNONYMS and SYNONYMS[n] in candidates:
        return SYNONYMS[n], 0.95, "synonym", [(SYNONYMS[n], 0.95)]

    # 4) 多路打分
    scored = []
    risky_high = set()
    for c in candidates:
        cn = normalize(c)
        # 老抽/海天/消天 即使和生抽编辑距离很近，也不参与打分
        if forbidden_auto_map(n, c):
            scored.append((c, 0.0))
            continue
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


def forbidden_auto_map(raw, candidate):
    """海天/老抽（以及消天这类形近 OCR）不得自动归到生抽。

    原文自己就写了「生抽」（如海天金标生抽）时不拦，那是生抽不是老抽。
    拿不准的暗抽保持原名，交给调用方标员工餐待确认。
    """
    raw_n = normalize(raw)
    cand_n = normalize(candidate)
    if not raw_n or not cand_n:
        return False
    if "生抽" not in cand_n or "老抽" in cand_n:
        return False
    if "老抽" in raw_n:
        return True
    if ("海天" in raw_n or "消天" in raw_n) and "生抽" not in raw_n:
        return True
    return False


def similar_mid_candidates(top, raw=None):
    """中分段里和最高分接近的候选。

    最高分已经到自动线，或最高分低于建议线时返回空（不是「几个中分难分」）。
    被海天/老抽→生抽 禁掉的候选不计入。
    """
    usable = []
    for name, score in top or []:
        if score < SUGGEST_CONF or score >= AUTO_CONF:
            continue
        if raw and forbidden_auto_map(raw, name):
            continue
        usable.append((name, score))
    if not usable:
        return []
    best_score = usable[0][1]
    close = []
    for name, score in usable:
        if best_score - score <= MID_TIE_GAP:
            close.append((name, score))
        else:
            break
    return close


def decide(raw, best, conf, source, top):
    """把一次 match 结果收成分层动作。默认不表示要调 LLM。

    action:
      auto    固定别名或高分，直接改名
      suggest 只有一个中分（或中分太多），只建议，等人确认
      llm     2～3 个中分接近，才把这几个候选交给 LLM
      review  弱相似 / 内置建议（腊肉），保留原名
      keep    分太低或被禁，保留原名（调用方按员工餐待确认处理）
    """
    top = list(top or [])
    if best and forbidden_auto_map(raw, best):
        return {"action": "keep", "name": raw, "suggestion": "", "source": "banned",
                "confidence": 0.0, "llm_candidates": []}
    if source in ("exact", "exact_strip", "variant", "alias", "synonym") or (
            best and conf >= AUTO_CONF and source != REVIEW_SOURCE):
        return {"action": "auto", "name": best, "suggestion": "", "source": source,
                "confidence": conf, "llm_candidates": []}
    if source == REVIEW_SOURCE:
        sg = suggest(raw, [name for name, _s in top] or ([best] if best else [])) or ""
        if sg and forbidden_auto_map(raw, sg):
            sg = ""
        return {"action": "review", "name": raw, "suggestion": sg, "source": REVIEW_SOURCE,
                "confidence": conf, "llm_candidates": []}
    close = similar_mid_candidates(top, raw)
    if len(close) in (2, 3):
        return {"action": "llm", "name": raw, "suggestion": close[0][0], "source": "llm",
                "confidence": close[0][1],
                "llm_candidates": [name for name, _s in close]}
    if close:
        # 一个中分，或超过 3 个挤在一起：不调 LLM，只把最高的一个交给人确认
        return {"action": "suggest", "name": raw, "suggestion": close[0][0], "source": "suggest",
                "confidence": close[0][1], "llm_candidates": []}
    return {"action": "keep", "name": raw, "suggestion": "", "source": source or "low",
            "confidence": conf, "llm_candidates": []}


def layered_match(raw, candidates, category="蔬菜", units=None, unit=None, codes=None):
    """分层匹配。除 action=llm 外，调用方不应再调大模型。"""
    best, conf, source, top = match(
        raw, candidates, category=category, units=units, unit=unit, codes=codes)
    return decide(raw, best, conf, source, top)


def correction_action(raw, best, conf, source):
    """决定这次匹配要不要改品名。

    返回 (品名, 提示人工)。
    needs_review：高分仍可能误改（香菜→香葱一类），保留原名并提示人工。
    海天/老抽→生抽：哪怕置信度已经被抬到自动线以上，也不替换。
    """
    if best and forbidden_auto_map(raw, best):
        return raw, True
    if source in (REVIEW_SOURCE, "suggest", "banned"):
        return raw, True
    if best and conf >= AUTO_CONF:
        return best, False
    return raw, False
