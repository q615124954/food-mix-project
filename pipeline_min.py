# -*- coding: utf-8 -*-
"""
最小可跑通路：采购单图片 → HunyuanOCR 识别 → 解析 → 写入暂存表
用途：验证"图 → 识别 → 入表"这条链路，后续再补勾稽/去重/轮询/白名单
"""
import os
import io
import re
import html as html_lib
from decimal import Decimal, ROUND_HALF_UP
import requests
from PIL import Image, ImageOps
from openpyxl import Workbook


def _load_env():
    """从项目根 .env 读取环境变量（无第三方依赖），已存在的环境变量优先"""
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if not os.path.exists(env_path):
        return
    with open(env_path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip())


_load_env()

# ===== 配置 =====
API_URL = "https://ai.gitee.com/v1/images/ocr"
API_KEY = os.environ.get("HUNYUAN_API_KEY", "")
MODEL = "HunyuanOCR"
MAX_SIZE = 2048          # 图片最长边上限
INPUT_DIR = r"D:\采购单"
OUTPUT_DIR = r"D:\财务表"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "采购入库_暂存.xlsx")


# ===== 1. 图片预处理 =====
def prepare_image(path):
    """按 EXIF 方向旋转校正 + 超限等比缩小，返回 (字节流, 文件名)"""
    img = Image.open(path)
    img = ImageOps.exif_transpose(img)  # 依据 EXIF Orientation 自动摆正
    w, h = img.size
    if max(w, h) > MAX_SIZE:
        scale = MAX_SIZE / max(w, h)
        img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
    buf = io.BytesIO()
    img.convert("RGB").save(buf, format="JPEG", quality=92)
    buf.seek(0)
    return buf, os.path.basename(path)


# ===== 2. 调用 OCR =====
def call_ocr(image_path):
    buf, name = prepare_image(image_path)
    files = {"image": (name, buf, "image/jpeg")}
    data = {"model": MODEL}
    r = requests.post(
        API_URL,
        headers={"Authorization": f"Bearer {API_KEY}"},
        files=files,
        data=data,
        timeout=180,
    )
    if r.status_code != 200:
        raise RuntimeError(f"OCR失败 {r.status_code}: {r.text[:200]}")
    return r.json().get("text_result") or r.json().get("text") or ""


# ===== 3. 解析 HTML 表格 =====
def parse_table(text):
    """
    从 OCR 输出中提取 <table> 的所有行，返回 [[cell, ...], ...]
    """
    rows = {}
    # OCR 输出的表格用 <tr>...<td> 组织，逐行切分
    trs = re.findall(r"<tr>(.*?)</tr>", text, flags=re.S)
    parsed = []
    for tr in trs:
        tds = re.findall(r"<td>(.*?)</td>", tr, flags=re.S)
        cells = [html_lib.unescape(td).strip() for td in tds]
        # 去掉纯空行
        if any(cells):
            parsed.append(cells)
    return parsed


HEADER_KEYWORDS = ("品名", "数量", "单价", "金额", "序号", "规格", "客户", "名称")


def is_header_row(cells):
    """表头行：同时命中多个列名关键字才算（避免把含'序号'字样的数据行误判为表头）"""
    joined = "".join(cells)
    hits = sum(1 for k in HEADER_KEYWORDS if k in joined)
    return hits >= 2


def extract_num(s):
    """从字符串提取第一个数字（含小数），失败返回 None"""
    if not s:
        return None
    m = re.search(r"-?\d+(?:\.\d+)?", s.replace("×", "x"))
    return float(m.group()) if m else None


def extract_date(text, filename):
    """从 OCR 文本提取「X月X日」，缺失回退文件名日期；都拿不到返回空串"""
    m = re.search(r"(\d{1,2})\s*月\s*(\d{1,2})\s*[日号]", text)
    if m:
        year = "2026"
        ym = re.search(r"(20\d{2})\s*年", text)
        if ym:
            year = ym.group(1)
        return f"{year}-{int(m.group(1)):02d}-{int(m.group(2)):02d}"
    fm = re.search(r"(20\d{2})(\d{2})(\d{2})", filename)
    if fm:
        return f"{fm.group(1)}-{fm.group(2)}-{fm.group(3)}"
    return ""


def parse_line_items(rows):
    """
    把表格行规整为 [品名, 数量, 金额]，单价由「金额÷数量」推导，不取 OCR 单价。
    支持三类单据：
    - 蔬菜单：4 列（品名/数量/单价/金额），只取品名、数量、金额；
    - 肉类单(复合数量)：5 列（序号/品名/空/"n×p"/金额），数量取复合数量前段、金额取金额列；
    - 肉类单(分开列)：5 列（序号/品名/数量/单价/金额），数量取数量列、金额取金额列。
    跳过重复表头、合计/件数等脏行。
    """
    items = []
    for cells in rows:
        if not cells or all(c.strip() == "" for c in cells):
            continue
        if is_header_row(cells):
            continue
        first = cells[0].strip() if cells else ""
        # 脏行：件数 / 合计 / 总计
        if first == "件数" or "合计" in first or "总计" in first:
            continue
        # 序号开头的单据（肉类/销货清单）：5 列，品名在 cells[1]，金额在 cells[4]
        # 注：OCR 偶把数据行序号"1"识别成"序号"，故 first=="序号" 也按数据行处理
        if re.fullmatch(r"\d+", first) or first == "序号":
            name = cells[1].strip() if len(cells) > 1 else ""
            if not name:
                continue
            amount = extract_num(cells[4]) if len(cells) > 4 else None
            # 数量：复合数量在 cells[3]（"n×p"），分开列则数量在 cells[2]
            if len(cells) > 3 and re.search(r"[×xX]", cells[3]):
                m = re.search(r"(\d+(?:\.\d+)?)\s*[×xX]", cells[3])
                qty = float(m.group(1)) if m else None
            else:
                qty = extract_num(cells[2]) if len(cells) > 2 else None
            items.append({"品名": name, "数量": qty, "金额": amount, "类别": "肉类"})
            continue
        # 品名开头的单据（蔬菜单）：4 列，品名在 cells[0]，数量 cells[1]，金额 cells[3]
        name = re.sub(r"\d+(\.\d+)?\s*$", "", cells[0]).strip()
        if not name:
            continue
        qty = extract_num(cells[1]) if len(cells) > 1 else None
        amount = extract_num(cells[3]) if len(cells) > 3 else None
        items.append({"品名": name, "数量": qty, "金额": amount, "类别": "蔬菜"})
    return items


# ===== 4. 单价推导 =====
CENT = Decimal("0.01")


def derive_price(items):
    """
    单价 = 金额 ÷ 数量，精确到分（ROUND_HALF_UP）。
    金额、数量严格取 OCR 原值，不纠正；金额或数量缺失则单价留空并标记。
    返回 (items, 统计)，item 新增「单价」「勾稽」字段。
    """
    stat = {"OK": 0, "缺金额": 0, "缺数量": 0, "数量为0": 0}
    for it in items:
        qty = it.get("数量")
        amount = it.get("金额")
        it["单价"] = None
        if qty is None:
            it["勾稽"] = "缺数量"
            stat["缺数量"] += 1
        elif qty == 0:
            it["勾稽"] = "数量为0"
            stat["数量为0"] += 1
        elif amount is None:
            it["勾稽"] = "缺金额"
            stat["缺金额"] += 1
        else:
            price = (Decimal(str(amount)) / Decimal(str(qty))).quantize(CENT, rounding=ROUND_HALF_UP)
            it["单价"] = float(price)
            it["勾稽"] = "OK"
            stat["OK"] += 1
    return items, stat


# ===== 5. 写入暂存表 =====
def write_staging(records):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "暂存区"
    headers = ["日期", "类别", "品名", "数量", "单价", "金额", "来源图片"]
    ws.append(headers)
    rows_out = []
    for rec in records:
        for item in rec["items"]:
            rows_out.append([
                item.get("日期", ""),
                item.get("类别", ""),
                item["品名"],
                item["数量"],
                item["单价"],
                item["金额"],
                rec["image"],
            ])
    # 按 日期 → 类别 → 品名 排序，同一天的数据连续不混淆
    rows_out.sort(key=lambda r: (r[0], r[1], r[2] or ""))
    for r in rows_out:
        ws.append(r)
    wb.save(OUTPUT_FILE)
    print(f"已写入: {OUTPUT_FILE}")


# ===== 主流程 =====
def main():
    images = sorted(
        os.path.join(INPUT_DIR, f)
        for f in os.listdir(INPUT_DIR)
        if f.lower().endswith((".jpg", ".jpeg", ".png"))
    )
    print(f"发现 {len(images)} 张图片")
    records = []
    total_stat = {"OK": 0, "缺金额": 0, "缺数量": 0, "数量为0": 0}
    for img in images:
        print(f"\n处理: {os.path.basename(img)}")
        try:
            text = call_ocr(img)
            date = extract_date(text, os.path.basename(img))
            rows = parse_table(text)
            items = parse_line_items(rows)
            items, stat = derive_price(items)
            for it in items:
                it["日期"] = date
            for k in total_stat:
                total_stat[k] += stat[k]
            print(f"  日期={date} | 解析出 {len(items)} 行 | 单价推导 OK={stat['OK']} 缺金额={stat['缺金额']} 缺数量={stat['缺数量']} 数量为0={stat['数量为0']}")
            records.append({"image": os.path.basename(img), "items": items, "raw": text})
        except Exception as e:
            print(f"  失败: {e}")
    if records:
        write_staging(records)
        print(f"\n汇总: OK={total_stat['OK']} 缺金额={total_stat['缺金额']} 缺数量={total_stat['缺数量']} 数量为0={total_stat['数量为0']}")


if __name__ == "__main__":
    main()