# -*- coding: utf-8 -*-
"""
采购单审核工作台后端
流程：扫描文件夹 → 后台批量 OCR → 网页审核修正 → 确认后落地 xlsx
复用 pipeline_min 的 OCR/解析/单价推导/日期提取逻辑
"""
import os
import sys
import json
import re
import shutil
import time
import threading
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from contextlib import asynccontextmanager

CENT = Decimal("0.01")

# 让 webapp 能 import 到上级目录的 pipeline_min
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pipeline_min as core
import name_match
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from typing import List

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
STATE_FILE = os.path.join(BASE_DIR, "state.json")
SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")
LAND_FILENAME = "采购入库_落地.xlsx"
NAME_DICT_FILE = os.path.join(BASE_DIR, "product_dict.json")
_name_dict = {"蔬菜": {}, "肉类": {}}
# 品名归并映射表（物料明细同步用，别名→目标品名，按类别隔离）
NAME_MAP_FILE = os.path.join(BASE_DIR, "name_map.json")
_name_map = {"肉类": {}, "蔬菜": {}}
# 品名→物料代码映射表（物料代码主键同步用，财务品名→物料代码，按类别隔离）
CODE_MAP_FILE = os.path.join(BASE_DIR, "code_map.json")
_code_map = {"肉类": {}, "蔬菜": {}}
# 合法品名集合（来自物料明细表「日物料采购」sheet 的品名列，归一化后）
_legal_names = set()
# 合法品名有序列表（原始品名，用于 name_match 模糊匹配候选集）
_legal_name_list = []
# 合法品名单位映射 {品名: 单位}，用于阶段2单位辅助匹配
_legal_units = {}

# 类别 -> 目标表 sheet 名
CATEGORY_SHEET = {"肉类": "肉类采购", "蔬菜": "日物料采购"}

_lock = threading.Lock()
_land_lock = threading.Lock()
_ocr_lock = threading.Lock()   # 防重入：同一时间只允许一个 ocr_worker 运行
OCR_CONCURRENCY = 3            # 并发识别数（平台慢，别调太大以免触发限流）
_state = {"invoices": []}
# 可配置的输入/输出目录与物料明细表路径
_paths = {"input_dir": core.INPUT_DIR, "output_dir": core.OUTPUT_DIR, "material_file": ""}


def land_file():
    return os.path.join(_paths["output_dir"], LAND_FILENAME)


# ===== 路径配置读写 =====
def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, encoding="utf-8") as f:
                s = json.load(f)
            if s.get("input_dir"):
                _paths["input_dir"] = s["input_dir"]
            if s.get("output_dir"):
                _paths["output_dir"] = s["output_dir"]
            if s.get("material_file"):
                _paths["material_file"] = s["material_file"]
        except (OSError, ValueError):
            pass


def save_settings():
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(_paths, f, ensure_ascii=False, indent=2)


# ===== 品名归并映射表：持久化 =====
def load_name_map():
    global _name_map
    if os.path.exists(NAME_MAP_FILE):
        try:
            with open(NAME_MAP_FILE, encoding="utf-8") as f:
                _name_map = json.load(f)
        except (OSError, ValueError):
            _name_map = {"肉类": {}, "蔬菜": {}}


def save_name_map():
    with open(NAME_MAP_FILE, "w", encoding="utf-8") as f:
        json.dump(_name_map, f, ensure_ascii=False, indent=2)


# ===== 品名→物料代码映射表：持久化 =====
def load_code_map():
    global _code_map
    if os.path.exists(CODE_MAP_FILE):
        try:
            with open(CODE_MAP_FILE, encoding="utf-8") as f:
                _code_map = json.load(f)
        except (OSError, ValueError):
            _code_map = {"肉类": {}, "蔬菜": {}}


def save_code_map():
    with open(CODE_MAP_FILE, "w", encoding="utf-8") as f:
        json.dump(_code_map, f, ensure_ascii=False, indent=2)


# ===== 合法品名表 + 员工餐归属判定 =====
def normalize_pin(name):
    """品名归一：去括号规格（如「枸杞 500g」「雨靴（39、41码）」）与空白，便于集合匹配"""
    s = str(name).strip()
    s = re.sub(r"[（(].*?[)）]", "", s)
    return re.sub(r"\s+", "", s)


def load_legal_names():
    """从物料明细表的「日物料采购」sheet 读取合法品名（每次路径变更/启动时刷新）。
    同时维护：归一化集合（归属判定）+ 原始品名有序列表（模糊匹配候选集）+ 单位映射。"""
    global _legal_names, _legal_name_list, _legal_units
    _legal_names = set()
    _legal_name_list = []
    _legal_units = {}
    material = _paths.get("material_file", "")
    if not material or not os.path.exists(material):
        return
    try:
        from openpyxl import load_workbook
        wb = load_workbook(material, read_only=True, data_only=True)
        if "日物料采购" in wb.sheetnames:
            ws = wb["日物料采购"]
            for r in range(3, ws.max_row + 1):
                v = ws.cell(r, 4).value
                if v and str(v).strip():
                    raw = str(v).strip()
                    _legal_name_list.append(raw)
                    _legal_names.add(normalize_pin(raw))
                    unit = str(ws.cell(r, 5).value or "").strip()
                    if unit:
                        _legal_units[raw] = unit
        wb.close()
    except Exception:
        pass


def classify_item(name, category):
    """行级归属：蔬菜单按合法品名表判定（命中→物料，未命中→员工餐）；肉类单全部物料。"""
    if category == "蔬菜":
        return "员工餐" if (name and normalize_pin(name) not in _legal_names) else "物料"
    return "物料"


# ===== 品名纠错字典：持久化 + 回流学习 =====
def load_name_dict():
    global _name_dict
    if os.path.exists(NAME_DICT_FILE):
        try:
            with open(NAME_DICT_FILE, encoding="utf-8") as f:
                _name_dict = json.load(f)
        except (OSError, ValueError):
            _name_dict = {"蔬菜": {}, "肉类": {}}


def save_name_dict():
    with open(NAME_DICT_FILE, "w", encoding="utf-8") as f:
        json.dump(_name_dict, f, ensure_ascii=False, indent=2)


def collect_valid_names():
    """合法品名集合：纠错字典 value ∪ 归并映射表 key ∪ 已确认单的「品名」。
    用于回流学习安全网（判断 OCR 原品名是否本身已是合法品名，避免把合法品名
    之间的串位如 牛肉↔牛骨 学成全局映射），也用作存疑判定的「认定品名」白名单。"""
    names = set()
    for table in _name_dict.values():
        for v in table.values():
            if v:
                names.add(str(v).strip())
    for table in _name_map.values():
        for k in table.keys():
            if k:
                names.add(str(k).strip())
    for inv in _state.get("invoices", []):
        if inv.get("status") != "confirmed":
            continue
        for it in inv.get("items", []):
            v = it.get("品名")
            if v:
                names.add(str(v).strip())
    # 加入物料明细表的合法品名（原始 + 归一化两种形式），避免合法品名被误判为「存疑」
    names.update(_legal_name_list)
    names.update(_legal_names)
    names.discard("")
    return names


def learn_name(category, wrong, right):
    """回流学习：把「OCR 原品名 → 确认后品名」写入纠错字典（幂等，同值跳过）。
    安全网：
    1) 若 wrong（原品名）本身是合法品名，说明是合法名之间的串位歧义而非 OCR
       垃圾串，拒绝学习，避免下次把真正品名反向纠正；
    2) 成环检测：right 已是 key 且映射回 wrong（A→B 与 B→A 同存），拒绝。"""
    if not wrong or not right or wrong == right:
        return
    wrong = str(wrong).strip()
    right = str(right).strip()
    if not wrong or not right or wrong == right:
        return
    if wrong in collect_valid_names():
        return
    table = _name_dict.setdefault(category, {})
    if table.get(right) == wrong:
        return
    table[wrong] = right
    save_name_dict()
    # 同步回流到变体表（标准名 -> 错名），让模糊匹配下次直接命中该错法
    name_match.add_variant(category, right, wrong)


# ===== 状态读写 =====
def save_state():
    with _lock:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(_state, f, ensure_ascii=False, indent=2)


def load_state():
    global _state
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, encoding="utf-8") as f:
            _state = json.load(f)


def image_files():
    d = _paths["input_dir"]
    if not os.path.isdir(d):
        return []
    return sorted(
        f for f in os.listdir(d)
        if f.lower().endswith((".jpg", ".jpeg", ".png"))
    )


# ===== 复用 pipeline_min 做单张 OCR =====
# 易混品名对：OCR 常把这两个合法品名互相认串（香葱↔香菜等），同单同时出现时全部标红提示人工核对
CONFUSABLE_PAIRS = [("番茄", "香草"), ("蒜米", "去皮蒜米"), ("蒜米", "蒜头")]


def compute_suspects(date, items):
    """OCR 后规则式存疑判定（零成本，不调 LLM），返回存疑原因列表（空=正常）。
    同时给每个 item 打上字段级存疑标记「品名存疑 / 数量存疑 / 金额存疑」，
    供前端在对应字段后显示红色感叹号，快速定位问题数据。"""
    suspects = []
    if not date:
        suspects.append("日期未识别")
    if not items:
        suspects.append("未识别到内容")
        return suspects
    cnt_q = cnt_a = cnt_z = 0
    bad_names = []
    # 第一遍：逐行打字段级标记 + 聚合计数
    for it in items:
        g = it.get("勾稽")
        it["数量存疑"] = g in ("缺数量", "数量为0")
        it["金额存疑"] = g == "缺金额"
        if g == "缺数量":
            cnt_q += 1
        elif g == "缺金额":
            cnt_a += 1
        elif g == "数量为0":
            cnt_z += 1
        nm = str(it.get("品名", "")).strip()
        fmt_bad = bool(nm) and (len(nm) == 1 or re.search(r"\d", nm) is not None)
        # 纠错后的最终品名不在物料表合法集合 → 才判存疑（OCR 认错但已纠正成功则信任）
        not_legal = bool(nm) and bool(_legal_names) and normalize_pin(nm) not in _legal_names
        # 仅 LLM 中置信（有歧义）纠正才提示核对；高置信 fuzzy 纠正已可信
        machine_fixed = it.get("纠正来源", "") == "llm"
        it["品名存疑"] = fmt_bad or not_legal or machine_fixed
        if it["品名存疑"]:
            bad_names.append(nm)
    # 易混品名对：同单同时出现时，整对都标红（合法名之间串位 OCR 很难自查）
    names_set = {str(it.get("品名", "")).strip() for it in items}
    if names_set:
        for a, b in CONFUSABLE_PAIRS:
            if a in names_set and b in names_set:
                for it in items:
                    if str(it.get("品名", "")).strip() in (a, b):
                        it["品名存疑"] = True
                        nm = str(it.get("品名", "")).strip()
                        if nm not in bad_names:
                            bad_names.append(nm)
    # 重复品名也标品名存疑（同一单内重复品名多为误识别）
    name_counts = {}
    for it in items:
        nm = str(it.get("品名", "")).strip()
        if nm:
            name_counts[nm] = name_counts.get(nm, 0) + 1
    dups = [nm for nm, c in name_counts.items() if c > 1]
    if dups:
        for it in items:
            nm = str(it.get("品名", "")).strip()
            if name_counts.get(nm, 0) > 1:
                it["品名存疑"] = True
        suspects.append(f"品名重复（{'、'.join(dups)}）")
    # 汇总单据级存疑原因
    if cnt_q:
        suspects.append(f"{cnt_q} 行缺数量")
    if cnt_a:
        suspects.append(f"{cnt_a} 行缺金额")
    if cnt_z:
        suspects.append(f"{cnt_z} 行数量为0")
    if bad_names:
        shown = "、".join(bad_names[:3]) + ("…" if len(bad_names) > 3 else "")
        suspects.append(f"{len(bad_names)} 行品名存疑（{shown}）")
    return suspects


def _count_missing_rows(rows):
    """统计「有序号、品名为空、且后续列(金额/备注)也全空」的表格行数。
    这类行通常是补货段/下半部分手写内容 OCR 没识别出来（只认出了序号），
    用于勾稽兜底：发现漏识别时提示人工核对补录。"""
    cnt = 0
    for cells in rows:
        if not cells or all(c.strip() == "" for c in cells):
            continue
        if core.is_header_row(cells):
            continue
        first = (cells[0] or "").strip()
        if not re.fullmatch(r"\d+", first):
            continue
        name = (cells[1] or "").strip() if len(cells) > 1 else ""
        if name:
            continue
        tail = "".join((c or "") for c in cells[2:])
        if not re.search(r"\d", tail):
            cnt += 1
    return cnt


def ocr_one(filename):
    path = os.path.join(_paths["input_dir"], filename)
    text = core.call_ocr(path)
    date = core.extract_date(text, filename)
    rows = core.parse_table(text)
    items = core.parse_line_items(rows)
    items, _stat = core.derive_price(items)
    # 漏识别兜底：有序号但内容为空的行较多（阴影遮挡/补货段未识别）时，
    # 用阴影补偿后的增强图重试一次 OCR，解析出更多明细时采用增强结果
    if _count_missing_rows(rows) >= 2 or not items:
        try:
            text2 = core.call_ocr(path, enhance=True)
            rows2 = core.parse_table(text2)
            items2, _s2 = core.derive_price(core.parse_line_items(rows2))
            if len(items2) > len(items):
                text, rows, items = text2, rows2, items2
                date = core.extract_date(text2, filename)
        except Exception:
            pass  # 增强重试失败则保留原结果，交人工核对
    total = core.sum_amount(items)
    order_total = core.extract_order_total(rows)
    for it in items:
        it["日期"] = date
        raw = it.get("品名", "")
        cat = it.get("类别", "蔬菜")
        corrected, _ = core.correct_name(raw, cat, _name_dict)
        src = "dict" if corrected != raw else ""
        # 纠错字典未命中时，对蔬菜走「变体表+拼音+单位」模糊匹配（阶段1+2）
        if corrected == raw and cat == "蔬菜" and _legal_name_list:
            best, conf, _src, top = name_match.match(
                raw, _legal_name_list, category=cat,
                units=_legal_units, unit=it.get("单位"))
            if best and conf >= name_match.AUTO_CONF:
                corrected = best
                src = "fuzzy"
            elif best and conf >= name_match.SUGGEST_CONF:
                # 中置信歧义：LLM 在 top 候选里定夺（带数量/金额/单位上下文）
                try:
                    mp = map_product_name(
                        raw, cat, [t[0] for t in top], "DeepSeek-V4-Pro",
                        qty=it.get("数量"), amt=it.get("金额"),
                        candidate_units=_legal_units)
                    if mp.get("来源") in ("mapping", "ai") and mp.get("目标品名") in _legal_name_list:
                        corrected = mp["目标品名"]
                        src = "llm"
                except Exception:
                    pass  # LLM 失败则保留 OCR 原值，交人工
        it["原品名"] = raw
        it["品名"] = corrected
        it["纠正来源"] = src
    suspects = compute_suspects(date, items)
    missing = _count_missing_rows(rows)
    if missing:
        suspects.append(f"疑似漏识别 {missing} 行（有序号但内容为空，可能补货段未识别，请人工核对补录）")
    return {"date": date, "total": total, "order_total": order_total, "items": items, "suspects": suspects}


def build_invoice_index():
    """按文件夹现有图片建索引，保留已有的 OCR 结果与确认状态"""
    existing = {inv["filename"]: inv for inv in _state.get("invoices", [])}
    invoices = []
    for f in image_files():
        if f in existing:
            inv = existing[f]
            # 旧状态可能没有 total，用明细金额补算
            if inv.get("total") in (None, ""):
                inv["total"] = core.sum_amount(inv.get("items", []))
            # 每次重算 suspects（基于现有明细），保证规则更新后重启即生效
            if inv.get("ocr_done") and inv.get("items"):
                inv["suspects"] = compute_suspects(inv.get("date", ""), inv.get("items", []))
            elif "suspects" not in inv:
                inv["suspects"] = []
            invoices.append(inv)
        else:
            invoices.append({
                "filename": f,
                "status": "pending",      # pending / confirmed
                "ocr_done": False,
                "date": "",
                "total": "",
                "order_total": "",
                "items": [],
                "error": "",
                "suspects": [],
            })
    with _lock:
        _state["invoices"] = invoices
    save_state()


def _ocr_one_inv(inv):
    """并发任务：OCR 单张票据（含 3 次重试），只更新该 inv 自己的字段。
    每个 inv 由单个工作线程独占更新，线程间无共享写，故无需加锁。返回是否成功。"""
    last_err = ""
    for _attempt in range(3):
        try:
            res = ocr_one(inv["filename"])
            inv["date"] = res["date"]
            inv["total"] = res.get("total")
            inv["order_total"] = res.get("order_total")
            inv["items"] = res["items"]
            inv["suspects"] = res.get("suspects", [])
            inv["ocr_done"] = True
            inv["error"] = ""
            return True
        except Exception as e:
            last_err = str(e)
            time.sleep(3)
    inv["error"] = last_err
    return False


def ocr_worker():
    """后台线程：线程池并发 OCR 尚未识别的单。
    并发注意事项：
    1) _ocr_lock 防重入，避免 rescan/reset/启动 多次触发导致多组 worker 重复识别同一张；
    2) 各张单由独立工作线程处理（只写各自的 inv 字段），无共享写冲突；
    3) save_state 只在主线程串行调用，避免并发写坏 state.json。"""
    if not _ocr_lock.acquire(blocking=False):
        return  # 已有 worker 在跑，直接退出
    try:
        pending = [inv for inv in _state["invoices"] if not inv.get("ocr_done")]
        if not pending:
            return
        from concurrent.futures import ThreadPoolExecutor
        with ThreadPoolExecutor(max_workers=OCR_CONCURRENCY) as ex:
            futures = [ex.submit(_ocr_one_inv, inv) for inv in pending]
            for fut in futures:
                try:
                    fut.result()   # 按提交顺序等待（线程池内部并发执行）
                except Exception:
                    pass
                save_state()       # 每完成一张即落一次状态，前端轮询可及时看到
    finally:
        _ocr_lock.release()


# ===== 数值规范化：前端提交的数字可能是字符串 =====
def _num(v):
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return v


def normalize_item(it, category=None):
    """规范化明细。category 为确认时用户选择的类别。
    未传入时保留明细上已有的类别；两边都没有则不填，不默认成蔬菜。"""
    out = {
        "品名": it.get("品名", ""),
        "原品名": it.get("原品名", ""),
        "品名编辑": bool(it.get("品名编辑")),
        "数量": _num(it.get("数量")),
        "单价": _num(it.get("单价")),
        "金额": _num(it.get("金额")),
    }
    if category not in (None, ""):
        cat = str(category).strip()
    else:
        cat = str(it.get("类别") or "").strip()
    if cat:
        out["类别"] = cat
    return out


def confirmed_categories():
    """文件名 -> 确认时保存的类别。没保存过类别的旧单不在表里，调用方不得默认成蔬菜。"""
    found = {}
    for inv in _state.get("invoices", []):
        fn = inv.get("filename")
        cat = str(inv.get("category") or "").strip()
        if fn and cat:
            found[fn] = cat
    return found


def resolve_sync_category(saved, land_cat):
    """同步用类别：优先确认时保存的类别；没有保存值时沿用落地表上的类别。
    两边都没有时返回空串，调用方必须提示，不得默认成蔬菜。"""
    saved = str(saved or "").strip()
    if saved:
        return saved
    return str(land_cat or "").strip()


# ===== 物料明细同步 =====
def _excel_serial(date_str):
    """'2026-08-27' -> Excel 日期序列号"""
    dt = datetime.strptime(str(date_str)[:10], "%Y-%m-%d")
    return (dt - datetime(1899, 12, 30)).days


def _land_date_text(value):
    """落地表日期单元格转成 YYYY-MM-DD。空单元格返回空字符串。

    合并单元格的续行、以及整张单都没填日期时，openpyxl 读出来都是 None。
    这里只负责「这个单元格自己有没有日期」，不把上一张单的日期带过来。
    """
    if value is None:
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    return text[:10] if text else ""


# 日期为空时的提示。确认接口标成存疑并拒绝落地；同步接口逐行跳过、不写 Excel。
BLANK_DATE_SUSPECT = "日期为空，需人工补填"
BLANK_DATE_SYNC_REASON = "日期为空，需人工补填日期，未写入 Excel"


def _extract_json(text):
    """从模型回复中提取 JSON 对象文本"""
    if not text:
        return ""
    m = re.search(r"\{.*\}", text, flags=re.S)
    return m.group(0) if m else text


def map_product_name(raw_name, category, candidates, model, qty=None, price=None, amt=None, candidate_units=None):
    """
    品名归并：先查映射表（按类别隔离），未命中再调 LLM。
    候选集带单位注解 + few-shot + 价格分品类提示，提升归并准确率。
    返回 dict：目标品名 / 置信度 / 理由 / 来源(mapping|ai|ai_error|ai_invalid)
    """
    mapping = _name_map.get(category, {})
    if raw_name in mapping:
        return {"目标品名": mapping[raw_name], "置信度": 1.0, "理由": "映射表命中", "来源": "mapping"}
    # 品名本身已是候选集标准名 → 直接命中，省掉一次 LLM 调用（同步性能关键）
    if raw_name in candidates:
        return {"目标品名": raw_name, "置信度": 1.0, "理由": "精确匹配", "来源": "exact"}
    # 仅差空白/括号等归一化差异的标准名也直接命中
    norm = normalize_pin(raw_name)
    for c in candidates:
        if normalize_pin(c) == norm:
            return {"目标品名": c, "置信度": 1.0, "理由": "精确匹配", "来源": "exact"}
    # 候选集名称（带单位注解，帮助模型区分相似品名）
    if candidate_units:
        cand_str = "、".join(
            f"{n}({candidate_units.get(n)})" if candidate_units.get(n) else n
            for n in candidates
        )
    else:
        cand_str = "、".join(candidates)
    system = (
        "你是餐饮采购品名归并助手。把采购单上的品名归并到给定的标准品名之一。"
        "判断要点：1) 结合单价区分品类：单价约2~3元/斤的多为骨头，30元以上的多为肉；"
        "2) 优先看品名字面含义，再结合单价佐证；3) 目标品名必须从候选集原样选取，一字不差。"
        "只输出 JSON，不要解释、不要多余文字。"
    )
    ctx = []
    if qty is not None:
        ctx.append(f"数量={qty}")
    if price is not None:
        ctx.append(f"单价={price}")
    if amt is not None:
        ctx.append(f"金额={amt}")
    ctx_str = ("，".join(ctx)) if ctx else "无"
    user = (
        f"标准品名候选集：{cand_str}\n"
        f"待归并品名：{raw_name}\n"
        f"类别：{category}\n"
        f"参考信息：{ctx_str}\n"
        f'输出格式示例：{{"目标品名":"牛大骨","置信度":0.9,"理由":"单价2.48为骨类"}}\n'
        f'注意：目标品名只输出品名本身，不要带括号或单位。\n'
        f'只输出 JSON：{{"目标品名":"候选集之一","置信度":0到1,"理由":"简短说明"}}'
    )
    try:
        text = core.call_chat(
            [{"role": "system", "content": system}, {"role": "user", "content": user}],
            model=model,
        )
    except Exception as e:
        return {"目标品名": raw_name, "置信度": 0.0, "理由": f"AI调用失败: {e}", "来源": "ai_error"}
    try:
        obj = json.loads(_extract_json(text))
    except Exception:
        return {"目标品名": raw_name, "置信度": 0.0, "理由": "AI返回无法解析", "来源": "ai_error"}
    target = str(obj.get("目标品名", "")).strip()
    # 容错：模型可能把单位"(斤)"带进品名，剥离括号后再匹配
    target_plain = re.sub(r"[（(].*?[)）]", "", target).strip()
    if target_plain in candidates:
        target = target_plain
    elif target not in candidates:
        return {"目标品名": raw_name, "置信度": 0.0, "理由": f"目标不在候选集: {target}", "来源": "ai_invalid"}
    try:
        conf = float(obj.get("置信度", 0))
    except (TypeError, ValueError):
        conf = 0.0
    return {"目标品名": target, "置信度": conf, "理由": str(obj.get("理由", "")), "来源": "ai"}


def _code_name_lookup(code, name_code):
    """从物料代码反查标准品名（用于 code_map 命中时回填目标品名）"""
    for nm, cd in name_code.items():
        if cd == code:
            return nm
    return code


def map_product_code(raw_name, category, name_code, model, qty=None, amt=None):
    """品名归并到物料代码：三层兜底，逐级返回，仅最后一级调 LLM。
    1) code_map 精确命中(已归并过的财务品名→代码)；
    2) 品名精确/归一化命中候选标准品名；
    3) name_match 本地召回 top5 候选 + LLM 定夺(LLM 返回候选品名, 再映射代码)。
    返回 dict：物料代码 / 目标品名 / 置信度 / 理由 / 来源。
    来源=="ai" 且置信度>=0.8 时写回 code_map 持久化，供重复同步零 LLM 复用。"""
    code_map_cat = _code_map.get(category, {})
    # 1) code_map 命中（且代码仍在当前候选 sheet 中）
    if raw_name in code_map_cat and code_map_cat[raw_name] in set(name_code.values()):
        code = code_map_cat[raw_name]
        return {"物料代码": code, "目标品名": _code_name_lookup(code, name_code),
                "置信度": 1.0, "理由": "代码映射表命中", "来源": "code_map"}
    # 2) 品名精确/归一化命中候选标准品名
    if raw_name in name_code:
        return {"物料代码": name_code[raw_name], "目标品名": raw_name,
                "置信度": 1.0, "理由": "品名精确匹配", "来源": "exact"}
    norm = normalize_pin(raw_name)
    candidates = list(name_code.keys())
    for c in candidates:
        if normalize_pin(c) == norm:
            return {"物料代码": name_code[c], "目标品名": c,
                    "置信度": 1.0, "理由": "品名归一化匹配", "来源": "exact"}
    # 3) name_match 本地召回 top5 候选，缩短 prompt 后喂 LLM
    top_names = candidates[:5]
    try:
        _best, _conf, _src, top = name_match.match(raw_name, candidates, category=category)
        if top:
            top_names = [t[0] for t in top]
    except Exception:
        pass
    cand_str = "、".join(top_names)
    system = (
        "你是餐饮采购品名归并助手。把采购单上的品名归并到给定的标准品名之一。"
        "判断要点：1) 结合单价区分品类：单价约2~3元/斤的多为骨头，30元以上的多为肉；"
        "2) 优先看品名字面含义，再结合单价佐证；3) 目标品名必须从候选集原样选取，一字不差。"
        "只输出 JSON，不要解释、不要多余文字。"
    )
    ctx = []
    if qty is not None:
        ctx.append(f"数量={qty}")
    if amt is not None:
        ctx.append(f"金额={amt}")
    ctx_str = ("，".join(ctx)) if ctx else "无"
    user = (
        f"标准品名候选集：{cand_str}\n"
        f"待归并品名：{raw_name}\n"
        f"类别：{category}\n"
        f"参考信息：{ctx_str}\n"
        f'输出格式示例：{{"目标品名":"牛大骨","置信度":0.9,"理由":"单价2.48为骨类"}}\n'
        f'注意：目标品名只输出品名本身，不要带括号或单位。\n'
        f'只输出 JSON：{{"目标品名":"候选集之一","置信度":0到1,"理由":"简短说明"}}'
    )
    try:
        text = core.call_chat(
            [{"role": "system", "content": system}, {"role": "user", "content": user}],
            model=model,
        )
    except Exception as e:
        return {"物料代码": "", "目标品名": raw_name, "置信度": 0.0, "理由": f"AI调用失败: {e}", "来源": "ai_error"}
    try:
        obj = json.loads(_extract_json(text))
    except Exception:
        return {"物料代码": "", "目标品名": raw_name, "置信度": 0.0, "理由": "AI返回无法解析", "来源": "ai_error"}
    target = str(obj.get("目标品名", "")).strip()
    target_plain = re.sub(r"[（(].*?[)）]", "", target).strip()
    if target_plain in name_code:
        target = target_plain
    elif target not in name_code:
        return {"物料代码": "", "目标品名": raw_name, "置信度": 0.0, "理由": f"目标不在候选集: {target}", "来源": "ai_invalid"}
    try:
        conf = float(obj.get("置信度", 0))
    except (TypeError, ValueError):
        conf = 0.0
    code = name_code[target]
    # 高置信 LLM 归并写回持久化（一次归并永久复用）
    if conf >= 0.8:
        _code_map.setdefault(category, {})[raw_name] = code
        save_code_map()
    return {"物料代码": code, "目标品名": target, "置信度": conf,
            "理由": str(obj.get("理由", "")), "来源": "ai"}


def _parse_target(ws):
    """解析矩阵式目标表：返回 (日期序列号->数量列号, 品名->行号, 品名->单位, 物料代码->行号, 品名->物料代码)"""
    header_row = None
    for r in range(1, min(ws.max_row, 60) + 1):
        cells = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        has_pin = any(str(c).strip() == "品名" for c in cells[:6])
        has_date = any(isinstance(c, (int, float)) and 40000 < c < 60000 for c in cells)
        if has_pin and has_date:
            header_row = r
            break
    if header_row is None:
        raise ValueError("未找到目标表表头行")
    date_cols = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if isinstance(v, (int, float)) and 40000 < v < 60000:
            date_cols[int(v)] = c
    name_rows = {}
    name_units = {}
    code_rows = {}
    name_code = {}
    for r in range(header_row + 2, ws.max_row + 1):
        name = ws.cell(r, 4).value
        if name and str(name).strip() and "合计" not in str(name):
            nm = str(name).strip()
            name_rows[nm] = r
            name_units[nm] = str(ws.cell(r, 5).value or "").strip()
            code = str(ws.cell(r, 2).value or "").strip()
            if code:
                code_rows[code] = r
                name_code[nm] = code
    return date_cols, name_rows, name_units, code_rows, name_code


def _parse_staff_sheet(ws):
    """解析员工餐sheet：返回 (日期序列号->行号, 员工餐金额列号, 日物料金额列号, 日物料总计列号)。
    表头第2行：序号|日期|类别|金额(员工餐) … 序号|日期|类别|金额(日物料)|总计。
    日期单元格为 Excel 序列号(General 格式)；合计行(日期为空)自动跳过。"""
    header_row = None
    for r in range(1, min(ws.max_row, 10) + 1):
        vals = [str(ws.cell(r, c).value or "").strip() for c in range(1, ws.max_column + 1)]
        if "总计" in vals:
            header_row = r
            break
    if header_row is None:
        raise ValueError("未找到员工餐sheet表头行")
    amount_cols = []
    total_col = None
    date_col = None
    for c in range(1, ws.max_column + 1):
        v = str(ws.cell(header_row, c).value or "").strip()
        if v == "金额":
            amount_cols.append(c)
        elif v == "总计" and total_col is None:
            total_col = c
        elif v == "日期" and date_col is None:
            date_col = c
    if len(amount_cols) < 2 or date_col is None:
        raise ValueError("员工餐sheet缺少金额/日期列")
    staff_amt_col = min(amount_cols)   # 员工餐·金额
    daily_amt_col = max(amount_cols)   # 日物料·金额
    if total_col is None:
        total_col = daily_amt_col + 1  # 总计列默认紧跟日物料金额列右侧
    date_row_map = {}
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, date_col).value
        if isinstance(v, (int, float)):
            date_row_map[int(v)] = r
    return date_row_map, staff_amt_col, daily_amt_col, total_col


def _cell_number(total):
    """累加后的 Decimal 写回单元格：整数值保持 int，其余转 float。"""
    if total == total.to_integral_value():
        return int(total)
    return float(total)


def sync_to_material(model):
    """把落地表数据同步到物料明细表（写前自动备份，只写数量+金额，单价/合计靠公式）。

    同一天同一物料若出现在多张已确认单据中，数量和金额先按日期+物料汇总再写入。
    写入的是落地表全部已确认单据的合计，不在单元格旧值上累加，因此重复同步幂等。
    """
    land = land_file()
    if not os.path.exists(land):
        raise HTTPException(400, f"落地表不存在：{land}")
    material = _paths.get("material_file", "")
    if not material or not os.path.exists(material):
        raise HTTPException(400, "物料明细表路径无效，请先在页面配置")
    if not material.lower().endswith(".xlsx"):
        raise HTTPException(400, "物料明细表必须是 .xlsx 文件")

    from openpyxl import load_workbook
    lwb = load_workbook(land, data_only=True)
    lws = lwb["已落地"] if "已落地" in lwb.sheetnames else lwb.active
    # 动态解析表头，兼容有无「原品名」列；归并以「品名」为准（OCR→纠错→人工确认后的最终值）
    header = [c.value for c in lws[1]]
    col = {str(h).strip(): i for i, h in enumerate(header) if h}
    raw_col = col.get("原品名")
    rows = []
    cur_date = ""
    cur_cat = ""
    cur_total = ""
    cur_paid = ""
    # 用哨兵区分「还没读到任何单据」和「来源图片本身是空的」
    cur_src = object()
    saved_cats = confirmed_categories()
    for r in lws.iter_rows(min_row=2, values_only=True):
        name = r[col["品名"]] if "品名" in col else None
        if name is None:
            continue
        src = r[col["来源图片"]] if "来源图片" in col else None
        date = r[col["日期"]] if "日期" in col else ""
        # 日期/类别/合计按来源图片合并，只有每张单的首行留着值，续行是空单元格。
        # 同一张单的续行继承本单日期和类别；换了一张单则只认本单自己的值。
        # 本单日期为空时不得沿用上一张单，否则会写进错误的日期列并覆盖那天原有的值。
        # 本单类别为空时同样不得沿用上一张（尤其不能默认为蔬菜）。
        if src != cur_src:
            cur_src = src
            cur_date = _land_date_text(date)
            cur_cat = ""
        else:
            own_date = _land_date_text(date)
            if own_date:
                cur_date = own_date
        cat = r[col["类别"]] if "类别" in col else ""
        if cat:
            cur_cat = cat
        cur_cat = resolve_sync_category(saved_cats.get(cur_src), cur_cat)
        qty = r[col["数量"]] if "数量" in col else None
        amt = r[col["金额"]] if "金额" in col else None
        price = r[col["单价"]] if "单价" in col else None
        raw = r[raw_col] if raw_col is not None else None
        # 合计/实付合计按来源图片合并，只在首行有值，需向下继承
        total_v = r[col["合计"]] if "合计" in col else None
        if total_v not in (None, ""):
            cur_total = total_v
        paid_v = r[col["实付合计"]] if "实付合计" in col else None
        if paid_v not in (None, ""):
            cur_paid = paid_v
        # 归并以「品名」（用户确认后的最终值）为准，「原品名」仅作追溯/回流学习，
        # 不再作为归并输入，否则会忽略用户手动修正的品名（如把「牛肉」误归成「牛骨」）
        merge_name = str(name).strip() if name else (str(raw).strip() if raw else "")
        # 实付总额：优先「实付合计」，旧落地表无该列时回退到「合计」（即无优惠）
        paid = cur_paid if cur_paid not in (None, "") else cur_total
        rows.append({"date": cur_date, "cat": cur_cat, "name": merge_name,
                     "qty": qty, "price": price, "amt": amt, "paid": paid})
    if not rows:
        raise HTTPException(400, "落地表没有可同步的数据")

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = material + f".bak_{ts}.xlsx"
    shutil.copy2(material, backup)

    twb = load_workbook(material)
    sheet_cache = {}
    map_cache = {}
    results = []
    # 日期为空的单据不参与写入（含员工餐汇总），避免落到错误列或参与同一天累加。
    # 有日期但缺少类别的行同样跳过，不进入 cell_sums，也不默认归入蔬菜。
    dated_rows = []
    for row in rows:
        if not row["date"]:
            results.append({"日期": "", "落地品名": row["name"], "目标品名": "", "置信度": 0,
                            "理由": BLANK_DATE_SYNC_REASON, "来源": "error"})
            continue
        if not str(row.get("cat") or "").strip():
            results.append({"日期": row["date"], "落地品名": row["name"], "目标品名": "", "置信度": 0,
                            "理由": "缺少类别，未默认归入蔬菜，请重新打开该单据选择类别并确认", "来源": "error"})
            continue
        dated_rows.append(row)
    # 同一天同一物料可能来自多张已确认单据。先按目标单元格（日期列+物料行）
    # 汇总落地表全部已确认行，再一次性写入合计；不读取、不累加单元格旧值，重复同步幂等。
    cell_sums = {}
    # 员工餐sheet汇总：按日期聚合蔬菜单实付合计（同一天只有一张蔬菜单，不做累加）
    staff = {}
    for row in dated_rows:
        if row["cat"] == "蔬菜" and row.get("paid") not in (None, ""):
            staff.setdefault(row["date"], {"paid": Decimal(str(row["paid"])), "material": Decimal("0")})

    for row in dated_rows:
        primary = CATEGORY_SHEET.get(row["cat"])
        alt = "日物料采购" if primary == "肉类采购" else ("肉类采购" if primary == "日物料采购" else None)
        sheets_to_try = [s for s in (primary, alt) if s and s in twb.sheetnames]
        if not sheets_to_try:
            results.append({"日期": row["date"], "落地品名": row["name"], "目标品名": "", "置信度": 0,
                            "理由": f"未识别类别或缺少sheet：{row['cat']}", "来源": "error"})
            continue
        placed = False
        for sheet_name in sheets_to_try:
            if sheet_name not in sheet_cache:
                ws = twb[sheet_name]
                date_cols, name_rows, name_units, code_rows, name_code = _parse_target(ws)
                sheet_cache[sheet_name] = (ws, date_cols, name_rows, name_units, code_rows, name_code)
            ws, date_cols, name_rows, name_units, code_rows, name_code = sheet_cache[sheet_name]
            key = (sheet_name, row["name"])
            mp = map_cache.get(key)
            if mp is None:
                mp = map_product_code(row["name"], row["cat"], name_code, model,
                                      qty=row.get("qty"), amt=row.get("amt"))
                map_cache[key] = mp
            if not mp.get("物料代码") or mp["物料代码"] not in code_rows:
                continue  # 该 sheet 匹配不上，试备选
            # 置信度不足（<0.8）不落地，避免错误品名被放入错误位置
            if mp["置信度"] < 0.8:
                results.append({"日期": row["date"], "落地品名": row["name"], "目标品名": mp["目标品名"],
                                "置信度": mp["置信度"], "理由": f"置信度不足(低于0.8)，跳过不落地：{mp['理由']}", "来源": mp["来源"]})
                placed = True
                break
            tr = code_rows[mp["物料代码"]]
            col_qty = date_cols.get(_excel_serial(row["date"]))
            if col_qty is None:
                results.append({"日期": row["date"], "落地品名": row["name"], "目标品名": mp["目标品名"],
                                "置信度": mp["置信度"], "理由": f"目标表无该日期列：{row['date']}", "来源": mp["来源"]})
                placed = True
                break
            bucket = cell_sums.get((sheet_name, tr, col_qty))
            if bucket is None:
                bucket = {"ws": ws, "qty": None, "amt": None}
                cell_sums[(sheet_name, tr, col_qty)] = bucket
            if row["qty"] is not None:
                qty = Decimal(str(row["qty"]))
                bucket["qty"] = qty if bucket["qty"] is None else bucket["qty"] + qty
            if row["amt"] is not None:
                amt = Decimal(str(row["amt"]))
                bucket["amt"] = amt if bucket["amt"] is None else bucket["amt"] + amt
            # 累计该单成功落到「日物料采购」的物料金额，供员工餐sheet员工餐金额=实付−物料合计
            if sheet_name == "日物料采购" and row["date"] in staff and row["amt"] is not None:
                staff[row["date"]]["material"] += Decimal(str(row["amt"]))
            results.append({"日期": row["date"], "落地品名": row["name"], "目标品名": mp["目标品名"],
                            "置信度": mp["置信度"], "理由": mp["理由"], "来源": mp["来源"],
                            "数量": row["qty"], "金额": row["amt"]})
            placed = True
            break
        if not placed:
            results.append({"日期": row["date"], "落地品名": row["name"], "目标品名": "", "置信度": 0,
                            "理由": "品名在肉类/日物料表中均无法归并", "来源": "error"})

    for (_sheet, tr, col_qty), bucket in cell_sums.items():
        ws = bucket["ws"]
        if bucket["qty"] is not None:
            ws.cell(tr, col_qty).value = _cell_number(bucket["qty"])
        if bucket["amt"] is not None:
            ws.cell(tr, col_qty + 2).value = _cell_number(bucket["amt"])

    # 写入员工餐sheet（仅蔬菜单）：日物料金额列=物料金额合计，日物料总计列=实付总额(优惠后)，
    # 员工餐金额列=实付合计−物料金额合计（员工餐+物料=总计）
    if "员工餐" in twb.sheetnames:
        try:
            staff_ws = twb["员工餐"]
            date_row_map, staff_amt_col, daily_amt_col, daily_total_col = _parse_staff_sheet(staff_ws)
        except Exception as e:
            results.append({"日期": "", "落地品名": "", "目标品名": "", "置信度": 0,
                            "理由": f"员工餐sheet解析失败：{e}", "来源": "error"})
        else:
            for d, v in staff.items():
                serial = _excel_serial(d)
                r = date_row_map.get(serial)
                if r is None:
                    results.append({"日期": d, "落地品名": "", "目标品名": "", "置信度": 0,
                                    "理由": "员工餐sheet无该日期行", "来源": "error"})
                    continue
                paid = v["paid"].quantize(CENT, rounding=ROUND_HALF_UP)
                material_sum = v["material"].quantize(CENT, rounding=ROUND_HALF_UP)
                staff_amt = (paid - material_sum).quantize(CENT, rounding=ROUND_HALF_UP)
                staff_ws.cell(r, daily_amt_col).value = float(material_sum)   # 日物料·金额 = 物料金额合计
                staff_ws.cell(r, daily_total_col).value = float(paid)         # 日物料·总计 = 实付总额(优惠后)
                staff_ws.cell(r, staff_amt_col).value = float(staff_amt)      # 员工餐·金额 = 实付-物料
                note = ""
                if staff_amt < 0:
                    note = "，注意：员工餐金额为负（优惠额>员工餐部分，请人工核对）"
                results.append({"日期": d, "落地品名": "", "目标品名": "员工餐",
                                "置信度": 1, "来源": "mapping",
                                "理由": f"日物料金额={float(material_sum):.2f}，日物料总计=实付{float(paid):.2f}，员工餐金额={float(staff_amt):.2f}{note}"})

    twb.save(material)
    return results, backup


# ===== 落地写 xlsx =====
def _merge_by_source(ws):
    """按「来源图片」分组合并：日期/类别/合计/实付合计四列纵向合并并居中"""
    from openpyxl.styles import Alignment
    for mr in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(mr))
    if ws.max_row < 2:
        return
    header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    try:
        img_col = header.index("来源图片") + 1
        date_col = header.index("日期") + 1
        cat_col = header.index("类别") + 1
        total_col = header.index("合计") + 1
    except ValueError:
        return
    paid_col = header.index("实付合计") + 1 if "实付合计" in header else None
    center = Alignment(horizontal="center", vertical="center")
    start = 2
    while start <= ws.max_row:
        key = ws.cell(start, img_col).value
        end = start
        while end + 1 <= ws.max_row and ws.cell(end + 1, img_col).value == key:
            end += 1
        merge_cols = [date_col, cat_col, total_col] + ([paid_col] if paid_col else [])
        for col in merge_cols:
            if end > start:
                ws.merge_cells(start_row=start, start_column=col, end_row=end, end_column=col)
            ws.cell(start, col).alignment = center
        start = end + 1


def append_land(filename, date, category, total, paid_total, items):
    os.makedirs(_paths["output_dir"], exist_ok=True)
    HEADERS = ["日期", "类别", "品名", "数量", "单价", "金额", "来源图片", "合计", "实付合计"]
    land = land_file()
    with _land_lock:
        if os.path.exists(land):
            from openpyxl import load_workbook
            wb = load_workbook(land)
            ws = wb.active
            # 先取消所有合并，避免删除列/行时冲突
            for mr in list(ws.merged_cells.ranges):
                ws.unmerge_cells(str(mr))
            # 兼容旧表头：删除旧「原品名」列，统一为固定 9 列
            for c in range(ws.max_column, 0, -1):
                if ws.cell(1, c).value == "原品名":
                    ws.delete_cols(c)
            ws.cell(1, 8).value = "合计"
            ws.cell(1, 9).value = "实付合计"
            # 按来源图片去重：先删掉该图片已有的旧行，再覆盖式写入
            for r in range(ws.max_row, 1, -1):
                if ws.cell(r, 7).value == filename:
                    ws.delete_rows(r)
        else:
            from openpyxl import Workbook
            wb = Workbook()
            ws = wb.active
            ws.title = "已落地"
            ws.append(HEADERS)
        for it in items:
            ws.append([
                date, category, it.get("品名"), it.get("数量"),
                it.get("单价"), it.get("金额"), filename, total, paid_total,
            ])
        _merge_by_source(ws)
        wb.save(land)


# ===== FastAPI 应用 =====
@asynccontextmanager
async def lifespan(app: FastAPI):
    load_settings()
    load_name_dict()
    load_name_map()
    load_code_map()
    load_legal_names()
    load_state()
    build_invoice_index()
    threading.Thread(target=ocr_worker, daemon=True).start()
    yield


app = FastAPI(lifespan=lifespan)


class ConfirmPayload(BaseModel):
    date: str = ""
    category: str = ""
    items: List[dict] = Field(default_factory=list)
    paid_total: float | None = None   # 实付总额（优惠后）；None 时按明细加总处理


class PathsPayload(BaseModel):
    input_dir: str = ""
    output_dir: str = ""
    material_file: str = ""


class SyncPayload(BaseModel):
    model: str = "DeepSeek-V4-Pro"


@app.get("/")
def index():
    # no-cache：强制浏览器每次取最新 index.html，避免旧缓存页（旧 init 调 reset）清空确认状态
    return FileResponse(os.path.join(STATIC_DIR, "index.html"),
                        headers={"Cache-Control": "no-cache, no-store, must-revalidate"})


@app.get("/api/state")
def get_state():
    invoices = _state.get("invoices", [])
    total = len(invoices)
    ocr_done = sum(1 for i in invoices if i.get("ocr_done"))
    confirmed = sum(1 for i in invoices if i.get("status") == "confirmed")
    return {
        "invoices": invoices,
        "summary": {"total": total, "ocr_done": ocr_done, "confirmed": confirmed},
        "paths": {"input_dir": _paths["input_dir"], "output_dir": _paths["output_dir"],
                  "material_file": _paths["material_file"]},
    }


@app.get("/images/{filename}")
def get_image(filename: str):
    p = os.path.join(_paths["input_dir"], filename)
    if not os.path.exists(p):
        raise HTTPException(404, "图片不存在")
    return FileResponse(p)


def _mark_blank_date(filename):
    """确认时日期为空：标为存疑，保持未确认，不写落地表。"""
    for inv in _state["invoices"]:
        if inv["filename"] == filename:
            suspects = list(inv.get("suspects") or [])
            if BLANK_DATE_SUSPECT not in suspects:
                suspects.append(BLANK_DATE_SUSPECT)
            inv["suspects"] = suspects
            save_state()
            return
    raise HTTPException(404, "未找到该票据")


@app.post("/api/invoice/{filename}/confirm")
def confirm(filename: str, payload: ConfirmPayload):
    # 日期为空不得落地：既不继承别的单据日期，也不把空日期写进 Excel
    if not _land_date_text(payload.date):
        _mark_blank_date(filename)
        raise HTTPException(400, "日期为空，请人工补填日期后再确认（本次未写入 Excel）")
    category = str(payload.category or "").strip()
    if category not in CATEGORY_SHEET:
        raise HTTPException(400, "请选择类别（蔬菜或肉类）后再确认；缺少类别时不会默认写成蔬菜")
    items = [normalize_item(it, category) for it in payload.items]
    # 回流学习：仅当用户手动编辑过品名时，才把「OCR 原品名 → 确认后品名」写入纠错字典
    for it in items:
        if it.get("品名编辑"):
            learn_name(category, it.get("原品名", ""), it.get("品名", ""))
    # code_map 联动：本次确认的全部品名集合之外/品名被改动的旧映射条目失效，
    # 避免人工纠正品名后，旧的「旧品名→代码」映射继续误同步到新单据。
    confirmed_names = {str(it.get("品名", "")).strip() for it in items}
    code_map_cat = _code_map.setdefault(category, {})
    if any(it.get("品名编辑") for it in items):
        for it in items:
            if it.get("品名编辑"):
                raw_old = str(it.get("原品名", "")).strip()
                if raw_old and raw_old not in confirmed_names:
                    code_map_cat.pop(raw_old, None)
        save_code_map()
    # 合计 = 所有品名金额累加（不再区分员工餐/物料），所有明细行都落地
    total = core.sum_amount(items)
    # 实付总额（优惠后）：用户未填或为空时按明细加总（即无优惠）
    paid = payload.paid_total if payload.paid_total not in (None, "") else total
    for inv in _state["invoices"]:
        if inv["filename"] == filename:
            # 先写落地表，成功后再标已确认。WPS 占用导致写入失败时保持未确认，
            # 避免网页显示已确认但 Excel 里没有这张单。
            try:
                append_land(filename, payload.date, category, total, paid, items)
            except PermissionError:
                raise HTTPException(500, "落地文件被占用，请先关闭 WPS 后重试")
            except Exception as e:
                raise HTTPException(500, f"落地文件写入失败：{e}")
            inv["date"] = payload.date
            inv["category"] = category
            inv["total"] = total
            inv["paid_total"] = paid
            inv["items"] = items
            inv["suspects"] = []   # 已确认，清除存疑标记
            inv["status"] = "confirmed"
            save_state()
            return {"ok": True}
    raise HTTPException(404, "未找到该票据")


@app.post("/api/paths")
def set_paths(payload: PathsPayload):
    global _state
    in_dir = (payload.input_dir or "").strip().strip('"')
    out_dir = (payload.output_dir or "").strip().strip('"')
    mat_file = (payload.material_file or "").strip().strip('"')
    # 输入目录必须真实存在
    if not in_dir or not os.path.isdir(in_dir):
        raise HTTPException(400, f"采购单目录不存在或不是文件夹：{in_dir}")
    # 输出目录不存在则创建；存在但不是目录则报错
    if not out_dir:
        raise HTTPException(400, "请填写财务表输出目录")
    if os.path.exists(out_dir) and not os.path.isdir(out_dir):
        raise HTTPException(400, f"输出路径不是文件夹：{out_dir}")
    os.makedirs(out_dir, exist_ok=True)
    # 物料明细表：可选，填了则校验是存在的 .xlsx 文件
    if mat_file and not mat_file.lower().endswith(".xlsx"):
        raise HTTPException(400, "物料明细表必须是 .xlsx 文件")
    if mat_file and not os.path.exists(mat_file):
        raise HTTPException(400, f"物料明细表文件不存在：{mat_file}")

    with _lock:
        _paths["input_dir"] = os.path.normpath(in_dir)
        _paths["output_dir"] = os.path.normpath(out_dir)
        _paths["material_file"] = os.path.normpath(mat_file) if mat_file else ""
        _state = {"invoices": []}
    save_settings()
    load_legal_names()
    try:
        if os.path.exists(STATE_FILE):
            os.remove(STATE_FILE)
    except OSError:
        pass
    build_invoice_index()
    threading.Thread(target=ocr_worker, daemon=True).start()
    return {"ok": True, "paths": dict(_paths)}


@app.post("/api/rescan")
def rescan():
    build_invoice_index()
    threading.Thread(target=ocr_worker, daemon=True).start()
    return {"ok": True}


@app.post("/api/reset")
def reset():
    """重置审核状态并重新扫描文件夹（已落地文件保留，按图片去重覆盖）"""
    global _state
    with _lock:
        _state = {"invoices": []}
    try:
        if os.path.exists(STATE_FILE):
            os.remove(STATE_FILE)
    except OSError:
        pass
    build_invoice_index()
    threading.Thread(target=ocr_worker, daemon=True).start()
    return {"ok": True}


@app.post("/api/sync")
def sync(payload: SyncPayload):
    """执行落地表 → 物料明细表的同步映射（全自动，写前自动备份）"""
    try:
        results, backup = sync_to_material(payload.model)
    except PermissionError:
        raise HTTPException(500, "物料明细表被占用，请先关闭 WPS/Excel 后重试")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f"同步失败：{e}")
    return {"ok": True, "count": len(results), "backup": backup, "results": results}


# ===== 系统原生文件/文件夹选择器 =====
def _pick_dialog(kind, initial=""):
    """弹系统选择器：kind='dir' 选文件夹，'file' 选 .xlsx 文件；取消返回空串。
    initial 为初始定位目录（文件夹选择器取目录本身，文件选择器取其父目录）。
    用独立子进程弹 tkinter 对话框，避免 tkinter 在 web 服务线程里的线程/消息循环问题。"""
    import subprocess
    init_dir = ""
    if initial:
        if kind == "dir":
            init_dir = initial if os.path.isdir(initial) else ""
        else:
            d = os.path.dirname(initial)
            init_dir = d if os.path.isdir(d) else ""
    init_arg = f", initialdir={init_dir!r}" if init_dir else ""
    pick = (
        f"print(filedialog.askdirectory(title='选择文件夹'{init_arg}) or '')"
        if kind == "dir"
        else f"print(filedialog.askopenfilename(title='选择物料明细表', filetypes=[('Excel','*.xlsx')]{init_arg}) or '')"
    )
    code = (
        "import tkinter as tk;from tkinter import filedialog;"
        "r=tk.Tk();r.withdraw();r.attributes('-topmost',True);r.update();"
        + pick + ";r.destroy()"
    )
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    kwargs = {}
    if os.name == "nt":  # Windows：隐藏子进程控制台窗口
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        si.wShowWindow = 0
        kwargs["startupinfo"] = si
        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
    r = subprocess.run(
        [sys.executable, "-c", code],
        capture_output=True, text=True, encoding="utf-8",
        timeout=600, env=env, **kwargs,
    )
    if not r.stdout:
        return ""
    return r.stdout.strip().splitlines()[-1].strip()


@app.get("/api/pick")
def pick(kind: str = "dir", initial: str = ""):
    if kind not in ("dir", "file"):
        raise HTTPException(400, "kind 必须是 dir 或 file")
    try:
        return {"path": _pick_dialog(kind, initial)}
    except Exception as e:
        raise HTTPException(500, f"无法打开选择器：{e}")


app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")