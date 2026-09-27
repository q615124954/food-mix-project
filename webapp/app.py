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
import threading
from datetime import datetime
from contextlib import asynccontextmanager

# 让 webapp 能 import 到上级目录的 pipeline_min
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pipeline_min as core
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

# 类别 -> 目标表 sheet 名
CATEGORY_SHEET = {"肉类": "肉类采购", "蔬菜": "日物料采购"}

_lock = threading.Lock()
_land_lock = threading.Lock()
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
    """合法品名集合：纠错字典所有 value ∪ 已确认单的「品名」。
    用于回流学习安全网——判断一个 OCR 原品名是否本身已是合法品名，避免把
    合法品名之间的串位（如 牛肉↔牛骨）学成全局映射、反向污染字典。"""
    names = set()
    for table in _name_dict.values():
        for v in table.values():
            if v:
                names.add(str(v).strip())
    for inv in _state.get("invoices", []):
        if inv.get("status") != "confirmed":
            continue
        for it in inv.get("items", []):
            v = it.get("品名")
            if v:
                names.add(str(v).strip())
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
def ocr_one(filename):
    path = os.path.join(_paths["input_dir"], filename)
    text = core.call_ocr(path)
    date = core.extract_date(text, filename)
    rows = core.parse_table(text)
    items = core.parse_line_items(rows)
    items, _stat = core.derive_price(items)
    total = core.sum_amount(items)
    for it in items:
        it["日期"] = date
        raw = it.get("品名", "")
        corrected, _ = core.correct_name(raw, it.get("类别", "蔬菜"), _name_dict)
        it["原品名"] = raw
        it["品名"] = corrected
    return {"date": date, "total": total, "items": items}


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
            invoices.append(inv)
        else:
            invoices.append({
                "filename": f,
                "status": "pending",      # pending / confirmed
                "ocr_done": False,
                "date": "",
                "total": "",
                "items": [],
                "error": "",
            })
    with _lock:
        _state["invoices"] = invoices
    save_state()


def ocr_worker():
    """后台线程：逐个 OCR 尚未识别的单，边做边落状态"""
    for inv in _state["invoices"]:
        if inv.get("ocr_done"):
            continue
        try:
            res = ocr_one(inv["filename"])
            inv["date"] = res["date"]
            inv["total"] = res.get("total")
            inv["items"] = res["items"]
            inv["ocr_done"] = True
            inv["error"] = ""
        except Exception as e:
            inv["error"] = str(e)
        save_state()


# ===== 数值规范化：前端提交的数字可能是字符串 =====
def _num(v):
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return v


def normalize_item(it):
    return {
        "品名": it.get("品名", ""),
        "原品名": it.get("原品名", ""),
        "品名编辑": bool(it.get("品名编辑")),
        "数量": _num(it.get("数量")),
        "单价": _num(it.get("单价")),
        "金额": _num(it.get("金额")),
    }


# ===== 物料明细同步 =====
def _excel_serial(date_str):
    """'2026-08-27' -> Excel 日期序列号"""
    dt = datetime.strptime(str(date_str)[:10], "%Y-%m-%d")
    return (dt - datetime(1899, 12, 30)).days


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


def _parse_target(ws):
    """解析矩阵式目标表：返回 (日期序列号->数量列号, 品名->行号, 品名->单位)"""
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
    for r in range(header_row + 2, ws.max_row + 1):
        name = ws.cell(r, 4).value
        if name and str(name).strip() and "合计" not in str(name):
            nm = str(name).strip()
            name_rows[nm] = r
            name_units[nm] = str(ws.cell(r, 5).value or "").strip()
    return date_cols, name_rows, name_units


def sync_to_material(model):
    """把落地表数据同步到物料明细表（写前自动备份，只写数量+金额，单价/合计靠公式）"""
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
    for r in lws.iter_rows(min_row=2, values_only=True):
        name = r[col["品名"]] if "品名" in col else None
        if name is None:
            continue
        date = r[col["日期"]] if "日期" in col else ""
        if date:
            cur_date = date
        cat = r[col["类别"]] if "类别" in col else ""
        if cat:
            cur_cat = cat
        qty = r[col["数量"]] if "数量" in col else None
        amt = r[col["金额"]] if "金额" in col else None
        price = r[col["单价"]] if "单价" in col else None
        raw = r[raw_col] if raw_col is not None else None
        # 归并以「品名」（用户确认后的最终值）为准，「原品名」仅作追溯/回流学习，
        # 不再作为归并输入，否则会忽略用户手动修正的品名（如把「牛肉」误归成「牛骨」）
        merge_name = str(name).strip() if name else (str(raw).strip() if raw else "")
        rows.append({"date": str(cur_date)[:10], "cat": cur_cat, "name": merge_name,
                     "qty": qty, "price": price, "amt": amt})
    if not rows:
        raise HTTPException(400, "落地表没有可同步的数据")

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = material + f".bak_{ts}.xlsx"
    shutil.copy2(material, backup)

    twb = load_workbook(material)
    sheet_cache = {}
    map_cache = {}
    results = []

    for row in rows:
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
                date_cols, name_rows, name_units = _parse_target(ws)
                sheet_cache[sheet_name] = (ws, date_cols, name_rows, name_units)
            ws, date_cols, name_rows, name_units = sheet_cache[sheet_name]
            candidates = list(name_rows.keys())
            key = (sheet_name, row["name"])
            mp = map_cache.get(key)
            if mp is None:
                mp = map_product_name(row["name"], row["cat"], candidates, model,
                                      qty=row.get("qty"), price=row.get("price"), amt=row.get("amt"),
                                      candidate_units=name_units)
                map_cache[key] = mp
            if mp["目标品名"] not in name_rows:
                continue  # 该 sheet 匹配不上，试备选
            tr = name_rows[mp["目标品名"]]
            col_qty = date_cols.get(_excel_serial(row["date"]))
            if col_qty is None:
                results.append({"日期": row["date"], "落地品名": row["name"], "目标品名": mp["目标品名"],
                                "置信度": mp["置信度"], "理由": f"目标表无该日期列：{row['date']}", "来源": mp["来源"]})
                placed = True
                break
            if row["qty"] is not None:
                ws.cell(tr, col_qty).value = row["qty"]
            if row["amt"] is not None:
                ws.cell(tr, col_qty + 2).value = row["amt"]
            results.append({"日期": row["date"], "落地品名": row["name"], "目标品名": mp["目标品名"],
                            "置信度": mp["置信度"], "理由": mp["理由"], "来源": mp["来源"],
                            "数量": row["qty"], "金额": row["amt"]})
            placed = True
            break
        if not placed:
            results.append({"日期": row["date"], "落地品名": row["name"], "目标品名": "", "置信度": 0,
                            "理由": "品名在肉类/日物料表中均无法归并", "来源": "error"})

    twb.save(material)
    return results, backup


# ===== 落地写 xlsx =====
def _merge_by_source(ws):
    """按「来源图片」分组合并：日期/类别/合计三列纵向合并并居中"""
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
    center = Alignment(horizontal="center", vertical="center")
    start = 2
    while start <= ws.max_row:
        key = ws.cell(start, img_col).value
        end = start
        while end + 1 <= ws.max_row and ws.cell(end + 1, img_col).value == key:
            end += 1
        for col in (date_col, cat_col, total_col):
            if end > start:
                ws.merge_cells(start_row=start, start_column=col, end_row=end, end_column=col)
            ws.cell(start, col).alignment = center
        start = end + 1


def append_land(filename, date, category, total, items):
    os.makedirs(_paths["output_dir"], exist_ok=True)
    HEADERS = ["日期", "类别", "品名", "数量", "单价", "金额", "来源图片", "合计", "原品名"]
    land = land_file()
    with _land_lock:
        if os.path.exists(land):
            from openpyxl import load_workbook
            wb = load_workbook(land)
            ws = wb.active
            # 兼容旧表头：确保「合计」「原品名」列存在
            if ws.cell(1, 8).value != "合计":
                ws.cell(1, 8).value = "合计"
            if ws.cell(1, 9).value != "原品名":
                ws.cell(1, 9).value = "原品名"
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
                it.get("单价"), it.get("金额"), filename, total, it.get("原品名", ""),
            ])
        _merge_by_source(ws)
        wb.save(land)


# ===== FastAPI 应用 =====
@asynccontextmanager
async def lifespan(app: FastAPI):
    load_settings()
    load_name_dict()
    load_name_map()
    load_state()
    build_invoice_index()
    threading.Thread(target=ocr_worker, daemon=True).start()
    yield


app = FastAPI(lifespan=lifespan)


class ConfirmPayload(BaseModel):
    date: str = ""
    category: str = ""
    items: List[dict] = Field(default_factory=list)


class PathsPayload(BaseModel):
    input_dir: str = ""
    output_dir: str = ""
    material_file: str = ""


class SyncPayload(BaseModel):
    model: str = "DeepSeek-V4-Pro"


@app.get("/")
def index():
    return FileResponse(os.path.join(STATIC_DIR, "index.html"))


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


@app.post("/api/invoice/{filename}/confirm")
def confirm(filename: str, payload: ConfirmPayload):
    items = [normalize_item(it) for it in payload.items]
    # 回流学习：仅当用户手动编辑过品名时，才把「OCR 原品名 → 确认后品名」写入纠错字典
    for it in items:
        if it.get("品名编辑"):
            learn_name(payload.category, it.get("原品名", ""), it.get("品名", ""))
    for inv in _state["invoices"]:
        if inv["filename"] == filename:
            total = core.sum_amount(items)
            inv["date"] = payload.date
            inv["total"] = total
            inv["items"] = items
            inv["status"] = "confirmed"
            save_state()
            try:
                append_land(filename, payload.date, payload.category, total, items)
            except PermissionError:
                raise HTTPException(500, "落地文件被占用，请先关闭 WPS 后重试")
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


app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")