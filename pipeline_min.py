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
            qty = extract_num(cells[3]) if len(cells) > 3 else None
            if qty is None:
                continue
            # 品名为空但数量、金额都有：手写品名没认出来，保留整行（品名留空），
            # 由存疑规则标红交人工补，不能整行丢掉让金额漏账。
            unit = re.sub(r"[\d.\s×xX]", "", cells[2]) if len(cells) > 2 else ""
            amount = extract_num(cells[5]) if len(cells) > 5 else None
            items.append({"品名": name, "数量": qty, "金额": amount, "类别": "蔬菜", "单位": unit})
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
        qty = extract_num(cells[1]) if len(cells) > 1 else None
        amount = extract_num(cells[3]) if len(cells) > 3 else None
        if not name and (qty is None or amount is None):
            continue
        items.append({"品名": name, "数量": qty, "金额": amount, "类别": "蔬菜"})
    return items


_TOTAL_TAG = re.compile(r"^(?:合计|总计|小计|总额|计)\s*[:：]?\s*(\d+(?:\.\d+)?)\s*$")


def _is_total_label(s):
    """合计行的标签：合计/总计/件数，或手写「总…」被 OCR 认歪（如 8/13「总没」）"""
    s = (s or "").strip()
    return s == "件数" or "合计" in s or "总计" in s or (s.startswith("总") and len(s) <= 3)


def extract_order_total(rows):
    """提取单据级实付总额(float, 到分)；识别不到返回 None（交人工填，不猜）。

    手写单上合计的写法有三种：
    1) 首列是「合计/总计/件数」，金额在末尾列；
    2) 7 列单的品名列写「总…」（OCR 常认歪成「总没」），数量列空，金额在金额列；
    3) 备注列手写「计110」「计:213」「合计 141」（补货后会再写一次新的合计）。
    单据中出现多次合计时取最后一次（补货后的最终合计）。"""
    found = None
    for cells in rows:
        if not cells or all((c or "").strip() == "" for c in cells):
            continue
        cells = [(c or "").strip() for c in cells]
        first = cells[0]
        if first == "件数" or "合计" in first or "总计" in first:
            for c in reversed(cells):
                v = extract_num(c)
                if v is not None:
                    found = round(v, 2)
                    break
            continue
        if len(cells) >= 7 and _is_total_label(cells[1]) and extract_num(cells[3]) is None:
            v = extract_num(cells[5]) if cells[5] else extract_num(cells[6])
            if v is not None:
                found = round(v, 2)
        for c in cells[1:]:
            m = _TOTAL_TAG.match(c)
            if m:
                found = round(float(m.group(1)), 2)
    return found


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