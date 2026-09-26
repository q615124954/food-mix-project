# -*- coding: utf-8 -*-
"""
采购单审核工作台后端
流程：扫描文件夹 → 后台批量 OCR → 网页审核修正 → 确认后落地 xlsx
复用 pipeline_min 的 OCR/解析/单价推导/日期提取逻辑
"""
import os
import sys
import json
import threading
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

_lock = threading.Lock()
_land_lock = threading.Lock()
_state = {"invoices": []}
# 可配置的输入/输出目录，默认取 pipeline_min 的常量
_paths = {"input_dir": core.INPUT_DIR, "output_dir": core.OUTPUT_DIR}


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
        except (OSError, ValueError):
            pass


def save_settings():
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(_paths, f, ensure_ascii=False, indent=2)


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


def learn_name(category, wrong, right):
    """回流学习：把「OCR 原品名 → 确认后品名」写入纠错字典（幂等，同值跳过）"""
    if not wrong or not right or wrong == right:
        return
    table = _name_dict.setdefault(category, {})
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
    HEADERS = ["日期", "类别", "品名", "数量", "单价", "金额", "来源图片", "合计"]
    land = land_file()
    with _land_lock:
        if os.path.exists(land):
            from openpyxl import load_workbook
            wb = load_workbook(land)
            ws = wb.active
            # 兼容旧表头：确保「合计」列存在（合计追加在第 8 列，第 7 列仍为来源图片）
            if ws.cell(1, 8).value != "合计":
                ws.cell(1, 8).value = "合计"
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
                it.get("单价"), it.get("金额"), filename, total,
            ])
        _merge_by_source(ws)
        wb.save(land)


# ===== FastAPI 应用 =====
@asynccontextmanager
async def lifespan(app: FastAPI):
    load_settings()
    load_name_dict()
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
        "paths": {"input_dir": _paths["input_dir"], "output_dir": _paths["output_dir"]},
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
    # 输入目录必须真实存在
    if not in_dir or not os.path.isdir(in_dir):
        raise HTTPException(400, f"采购单目录不存在或不是文件夹：{in_dir}")
    # 输出目录不存在则创建；存在但不是目录则报错
    if not out_dir:
        raise HTTPException(400, "请填写财务表输出目录")
    if os.path.exists(out_dir) and not os.path.isdir(out_dir):
        raise HTTPException(400, f"输出路径不是文件夹：{out_dir}")
    os.makedirs(out_dir, exist_ok=True)

    with _lock:
        _paths["input_dir"] = os.path.normpath(in_dir)
        _paths["output_dir"] = os.path.normpath(out_dir)
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


app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")