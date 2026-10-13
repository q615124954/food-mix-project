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
from PIL import Image, ImageOps, ImageChops, ImageFilter
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
CHAT_URL = "https://ai.gitee.com/v1/chat/completions"
# 混元与模力方舟是同一把密钥：优先 HUNYUAN_API_KEY，为空再用 MOLLI_ARK_API_KEY
API_KEY = (
    os.environ.get("HUNYUAN_API_KEY", "").strip()
    or os.environ.get("MOLLI_ARK_API_KEY", "").strip()
)
MODEL = "HunyuanOCR"
MAX_SIZE = 2048          # 图片最长边上限
INPUT_DIR = r"D:\采购单"
OUTPUT_DIR = r"D:\财务表"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "采购入库_暂存.xlsx")


# ===== 1. 图片预处理 =====
def _shadow_compensate(img):
    """阴影/光照补偿：灰度图减去大半径高斯模糊估计出的纸面背景亮度，
    保留比背景暗的手写笔画，再拉伸对比度。
    用于拍摄时被阴影（手/手机投影）遮挡的区域，二次识别可显著提升可读性。"""
    try:
        gray = img.convert("L")
        bg = gray.filter(ImageFilter.GaussianBlur(radius=40))
        diff = ImageChops.subtract(gray, bg)
        # 对比度拉伸：笔画更黑、纸面更白
        diff = diff.point(lambda p: min(255, int(p * 2.0)))
        return diff.convert("RGB")
    except Exception:
        return img


def prepare_image(path, enhance=False):
    """按 EXIF 方向旋转校正 + 超限等比缩小，返回 (字节流, 文件名)。
    enhance=True 时先做阴影补偿（用于阴影遮挡导致漏识别后的重试）。"""
    img = Image.open(path)
    img = ImageOps.exif_transpose(img)  # 依据 EXIF Orientation 自动摆正
    w, h = img.size
    if max(w, h) > MAX_SIZE:
        scale = MAX_SIZE / max(w, h)
        img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
    if enhance:
        img = _shadow_compensate(img)
    buf = io.BytesIO()
    img.convert("RGB").save(buf, format="JPEG", quality=92)
    buf.seek(0)
    return buf, os.path.basename(path)


def call_ocr(image_path, timeout=600, enhance=False):
    """调用 HunyuanOCR 专用接口识别图片，返回 HTML 表格文本。
    注：模力方舟 OCR 平台侧偶发极慢（实测 400+ 秒/张），超时须放宽到 10 分钟。"""
    buf, name = prepare_image(image_path, enhance)
    files = {"image": (name, buf, "image/jpeg")}
    data = {"model": MODEL}
    r = requests.post(
        API_URL,
        headers={"Authorization": f"Bearer {API_KEY}"},
        files=files,
        data=data,
        timeout=timeout,
    )
    if r.status_code != 200:
        raise RuntimeError(f"OCR失败 {r.status_code}: {r.text[:200]}")
    return r.json().get("text_result") or r.json().get("text") or ""


# ===== 2.5 文本 LLM 调用（品名归并等） =====
def call_chat(messages, model="DeepSeek-V4-Pro", temperature=0.0, max_tokens=512):
    """调用模力方舟 OpenAI 兼容 chat/completions，返回回复文本"""
    r = requests.post(
        CHAT_URL,
        headers={"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"},
        json={"model": model, "messages": messages, "temperature": temperature, "max_tokens": max_tokens},
        timeout=120,
    )
    if r.status_code != 200:
        raise RuntimeError(f"LLM失败 {r.status_code}: {r.text[:200]}")
    data = r.json()
    try:
        return data["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError):
        return ""


# ===== 3. 解析 HTML 表格 =====
def parse_table(text):
    """
    从 HunyuanOCR 输出的 HTML 中提取 <table> 的所有行，返回 [[cell, ...], ...]
    """
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


def clean_name(name):
    """品名清洗：去尾部粘连的数量数字（如「薄荷4」→「薄荷」）。
    只删末尾纯数字，保留品名中间的规格数字（如「20升 食用油」「鲜味生抽 1.9升」）。"""
    if not name:
        return ""
    return re.sub(r"\d+(\.\d+)?\s*$", "", name).strip()


def name_or_raw(cell):
    """品名清洗后为空（OCR 把品名认成纯数字，如手写「盐」被认成「10」）时，
    保留 OCR 原值不再整行丢弃；纯数字品名会由存疑规则标红，人工修正后可回流学习。"""
    name = clean_name(cell)
    if name:
        return name
    return (cell or "").strip()


def extract_date(text, filename):
    """从 OCR 文本提取「X月X日」，缺失回退文件名日期；都拿不到返回空串"""
    m = re.search(r"(\d{1,2})\s*月\s*(\d{1,2})\s*[日号]", text)
    if m:
        year = "2026"
        ym = re.search(r"(20\d{2})\s*年", text)
        if ym:
            year = ym.group(1)
        return f"{year}-{int(m.group(1)):02d}-{int(m.group(2)):02d}"
    # 注意：文件名里的日期是微信保存时间（如 20260902），并非单据日期，
    # 回退到文件名会得到错误日期（如把 8 月的单记成 9 月），故识别失败时留空，交人工在前端填写。
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
        # 蔬菜单(7列)：序号|品名|单位|数量|单价|金额|备注。
        # 序号列可为空（补货段重新编号），品名在 cells[1]，数量 cells[3]，金额 cells[5]。
        # 数量列为空 → 小计/总计/签名等脏行，跳过（其金额落在金额列或备注列，非明细行）。
        # 单位列 cells[2] 偶为「50斤」「1斤」数字+单位粘连，剥离数字只留单位字，供阶段2辅助匹配。
        if len(cells) >= 7:
            name = name_or_raw(cells[1]) if len(cells) > 1 else ""
            if not name:
                continue
            qty = extract_num(cells[3]) if len(cells) > 3 else None
            if qty is None:
                continue
            unit = re.sub(r"[\d.\s×xX]", "", cells[2]) if len(cells) > 2 else ""
            amount = extract_num(cells[5]) if len(cells) > 5 else None
            items.append({"品名": name, "数量": qty, "金额": amount, "类别": "蔬菜", "单位": unit,
                          "票面单价": extract_num(cells[4]) if len(cells) > 4 else None})
            continue
        # 序号开头的单据（肉类/销货清单）：5 列，品名在 cells[1]，金额在 cells[4]
        # 注：OCR 偶把数据行序号"1"识别成"序号"，故 first=="序号" 也按数据行处理
        if re.fullmatch(r"\d+", first) or first == "序号":
            name = name_or_raw(cells[1]) if len(cells) > 1 else ""
            if not name:
                continue
            amount = extract_num(cells[4]) if len(cells) > 4 else None
            # 数量：优先识别「复合数量」(数字×数字，如 11.9×33 / 20×44)，可能落在
            # 数量列或品名列后一列；否则按「分开列」取数量列。
            # 注意：单价列偶被 OCR 成「×25」(乘号在数字前)，不能误判为复合数量。
            comp = re.compile(r"(\d+(?:\.\d+)?)\s*[×xX]\s*\d")
            qty = None
            for idx in (3, 2):
                if len(cells) > idx:
                    m = comp.search(cells[idx])
                    if m:
                        qty = float(m.group(1))
                        break
            if qty is None and len(cells) > 2:
                qty = extract_num(cells[2])
            items.append({"品名": name, "数量": qty, "金额": amount, "类别": "肉类"})
            continue
        # 品名开头的单据（蔬菜单）：4 列，品名在 cells[0]，数量 cells[1]，金额 cells[3]
        name = name_or_raw(cells[0])
        if not name:
            continue
        qty = extract_num(cells[1]) if len(cells) > 1 else None
        amount = extract_num(cells[3]) if len(cells) > 3 else None
        items.append({"品名": name, "数量": qty, "金额": amount, "类别": "蔬菜",
                      "票面单价": extract_num(cells[2]) if len(cells) > 2 else None})
    return items


def extract_order_total(rows):
    """从解析后的表格行中找「合计/总计/件数」行，提取单据级实付总额(float, 到分)。
    蔬菜单底部金额总计常有手动优惠，此值即优惠后实付；识别不到返回 None。
    只读 rows，不改动 parse_line_items 的明细提取逻辑。"""
    for cells in rows:
        if not cells or all((c or "").strip() == "" for c in cells):
            continue
        first = (cells[0] or "").strip() if cells else ""
        if first != "件数" and "合计" not in first and "总计" not in first:
            continue
        # 金额通常落在金额列（蔬菜单 7 列格式为 cells[5]）或末尾列，从右往左找第一个数字
        for c in reversed(cells):
            v = extract_num((c or "").strip())
            if v is not None:
                return round(v, 2)
    return None


# ===== 4. 单价推导 =====
CENT = Decimal("0.01")


def sum_amount(items):
    """对明细「金额」列求和（Decimal，缺失按 0 跳过），返回 float；无任何有效金额时返回 None"""
    total = Decimal("0")
    has_value = False
    for it in items:
        v = it.get("金额")
        if v is None or v == "":
            continue
        try:
            total += Decimal(str(v))
            has_value = True
        except (ValueError, TypeError):
            continue
    if not has_value:
        return None
    return float(total.quantize(CENT, rounding=ROUND_HALF_UP))


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


# ===== 4.4 单价异常（数量/金额存疑） =====
# 单价是「金额÷数量」反算的，数量丢小数点（5.8→58）时反算单价跟着变小，
# 「数量×单价≠金额」永远不会触发。所以改为拿外部参照比，偏离超过 PRICE_RATIO_LIMIT 倍才判：
#   1) 历史单价（物料表该品名各日期 金额÷数量 的中位数）有值时以它为准：
#      反算单价与历史单价相差超过 3 倍 → 存疑；在 3 倍以内 → 不判
#      （票面单价列常被 OCR 丢小数点，如 3.9 认成 39，此时历史单价证明数量/金额没问题）。
#   2) 没有历史单价、票面有单价列：数量×票面单价 与 金额 相差超过 3 倍 → 存疑。
#   3) 两者都没有 → 不判，避免误报。
# 手写单常把金额抹零（3×3.8=11.4 记 12），所以阈值按倍数不按百分比。
# 只标记交人工核对，不自动改数量/金额。
PRICE_RATIO_LIMIT = 3.0


def _ratio_off(a, b):
    """a、b 均为正数时返回 max(a/b, b/a)"""
    r = a / b
    return max(r, 1 / r)


def price_anomaly(item, hist_price=None):
    """返回存疑说明（字符串）或 None。item 需要 数量/金额，可带 票面单价。"""
    try:
        qty = float(item.get("数量"))
        amount = float(item.get("金额"))
    except (TypeError, ValueError):
        return None
    if qty <= 0 or amount <= 0:
        return None
    price = amount / qty
    if hist_price and hist_price > 0:
        off = _ratio_off(price, hist_price)
        if off > PRICE_RATIO_LIMIT:
            return (f"单价{price:.2f}（金额{amount:g}÷数量{qty:g}）与历史单价{hist_price:.2f}"
                    f"相差{off:.1f}倍，请核对数量/金额（可能丢了小数点）")
        return None
    try:
        ticket = float(item.get("票面单价"))
    except (TypeError, ValueError):
        ticket = None
    if ticket and ticket > 0:
        expect = qty * ticket
        off = _ratio_off(expect, amount)
        if off > PRICE_RATIO_LIMIT:
            return (f"数量×票面单价={qty:g}×{ticket:g}={expect:g}，与金额{amount:g}"
                    f"相差{off:.1f}倍，请核对数量/金额（可能丢了小数点）")
    return None


# ===== 4.5 品名纠错 =====
def correct_name(name, category, dictionary):
    """
    品名纠错：仅精确匹配（错名 → 正确名）。
    dictionary 形如 {"蔬菜": {错误名: 正确名}, "肉类": {...}}，按类别隔离避免跨类误纠。
    未命中则原样返回，不做改动。
    """
    if not name or not dictionary:
        return name, False
    table = dictionary.get(category) or {}
    if name in table:
        return table[name], True
    return name, False


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