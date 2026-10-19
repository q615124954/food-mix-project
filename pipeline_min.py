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
from PIL import Image, ImageOps, ImageChops, ImageFilter, ImageMath
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


# 原图一个字都没认出来时，按这个顺序各识一次，取内容最全的结果。
# shadow：上面的灰度相减（旧行为）。粉纸上的蓝笔在亮度里几乎看不见，相减还会把
# 整块手影压成黑，所以后面两档改成「除以背景」平场。
# red：蓝笔在红通道最暗、粉纸最亮；红减蓝有反差时用它，否则退回红通道。
# flat：不看颜色，直接把亮度除掉大块阴影。黑笔被手影压扁时用。
OCR_RETRY_PASSES = ("shadow", "red", "flat")

# 红减蓝几乎是一张空图（黑笔、或照片已经是灰度）时不要拿噪声当墨水。
_INK_PLANE_VAR = 100


def _channel_variance(channel):
    hist = channel.histogram()
    total = sum(hist) or 1
    mean = sum(i * hist[i] for i in range(256)) / total
    return sum(hist[i] * (i - mean) ** 2 for i in range(256)) / total


def _flat_field(channel):
    """除以模糊背景，消掉手/手机投影。笔画比纸面窄，半径盖过笔画、盖不住整块阴影。"""
    w, h = channel.size
    radius = max(12, min(24, min(w, h) // 24))
    bg = channel.filter(ImageFilter.GaussianBlur(radius=radius))
    flat = ImageMath.lambda_eval(
        lambda a: a["convert"](
            a["min"](a["ch"] * 255 / a["max"](a["bg"], 1), 255), "L"),
        ch=channel,
        bg=bg,
    )
    return flat


def _ink_plane(img):
    """粉/紫复写纸上的蓝笔：红减蓝。差值没有反差时退回红通道（黑笔、灰度图）。"""
    red, _green, blue = img.convert("RGB").split()
    diff = ImageChops.subtract(red, blue)
    if _channel_variance(diff) < _INK_PLANE_VAR:
        return red
    return diff


def _red_ink_enhance(img):
    plane = _flat_field(_ink_plane(img))
    plane = plane.filter(ImageFilter.UnsharpMask(radius=1.2, percent=140, threshold=2))
    return plane.convert("RGB")


def _flat_field_enhance(img):
    plane = _flat_field(ImageOps.grayscale(img))
    plane = plane.filter(ImageFilter.SHARPEN)
    return plane.convert("RGB")


def enhance_image(img, strategy):
    """strategy：True/'shadow' 原阴影补偿；'red' 粉纸蓝笔；'flat' 亮度平场。
    不认识的值原图返回，避免一次预处理失败把图涂掉。"""
    try:
        if strategy in (True, "shadow"):
            return _shadow_compensate(img)
        if strategy == "red":
            return _red_ink_enhance(img)
        if strategy == "flat":
            return _flat_field_enhance(img)
    except Exception:
        return img
    return img


def prepare_image(path, enhance=False):
    """按 EXIF 方向旋转校正 + 超限等比缩小，返回 (字节流, 文件名)。
    enhance 为 True/'shadow'/'red'/'flat' 时先做对应预处理（漏识别后的重试）。"""
    img = Image.open(path)
    img = ImageOps.exif_transpose(img)  # 依据 EXIF Orientation 自动摆正
    w, h = img.size
    if max(w, h) > MAX_SIZE:
        scale = MAX_SIZE / max(w, h)
        img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
    if enhance:
        img = enhance_image(img, enhance)
    buf = io.BytesIO()
    img.convert("RGB").save(buf, format="JPEG", quality=92)
    buf.seek(0)
    return buf, os.path.basename(path)


def _image_to_jpeg_buf(img, name="image.jpg"):
    """最长边超过 MAX_SIZE 时等比缩小，返回 (JPEG 字节流, 文件名)。"""
    img = img.convert("RGB")
    w, h = img.size
    if max(w, h) > MAX_SIZE:
        scale = MAX_SIZE / float(max(w, h))
        img = img.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=92)
    buf.seek(0)
    return buf, name


def _post_ocr_buf(buf, name, timeout):
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
    body = r.json()
    return body.get("text_result") or body.get("text") or ""


def call_ocr(image_path, timeout=600, enhance=False):
    """调用 HunyuanOCR 专用接口识别图片，返回 HTML 表格文本。
    注：模力方舟 OCR 平台侧偶发极慢（实测 400+ 秒/张），超时须放宽到 10 分钟。"""
    buf, name = prepare_image(image_path, enhance)
    return _post_ocr_buf(buf, name, timeout)


def looks_like_long_screenshot(path):
    """手机长截图：窄、且高度超过 OCR 最长边。手写单照片通常更宽，不走这里。

    打不开文件时返回 False，调用方继续用整图 OCR。"""
    try:
        img = Image.open(path)
        img = ImageOps.exif_transpose(img)
        w, h = img.size
    except Exception:
        return False
    if w <= 0 or h <= 0:
        return False
    return w <= 1600 and h > MAX_SIZE and (h / float(w)) >= 2.4


def iter_vertical_slices(img, max_edge=None, overlap=0.15):
    """把长图切成竖向重叠条，每条最长边不超过 max_edge，避免整图缩成一条窄图。"""
    max_edge = MAX_SIZE if max_edge is None else int(max_edge)
    w, h = img.size
    if w <= 0 or h <= 0:
        return
    slice_h = h if w > max_edge else min(h, max_edge)
    if h <= slice_h:
        yield img
        return
    step = max(1, int(slice_h * (1 - overlap)))
    y = 0
    while y < h:
        if h - y < slice_h and y > 0:
            y = h - slice_h
        bottom = min(h, y + slice_h)
        yield img.crop((0, y, w, bottom))
        if bottom >= h:
            break
        y += step


def join_ocr_slices(texts):
    """按切片顺序拼接 OCR 文本。相邻切片若有完全相同的行重叠，只保留一份。"""
    acc = []
    for text in texts or []:
        lines = [ln.strip() for ln in str(text or "").splitlines() if ln.strip()]
        if not lines:
            continue
        if not acc:
            acc = lines
            continue
        max_k = min(len(acc), len(lines), 12)
        cut = 0
        for k in range(max_k, 0, -1):
            if acc[-k:] == lines[:k]:
                cut = k
                break
        acc.extend(lines[cut:])
    deduped = []
    for ln in acc:
        if deduped and deduped[-1] == ln:
            continue
        deduped.append(ln)
    return "\n".join(deduped)


def call_ocr_tiled(image_path, timeout=600, overlap=0.15):
    """长截图分条识别再拼接。某一条失败则整次抛出，调用方回退整图 OCR。

    只拼接模型返回的文字，不根据原图像素补数量或金额。"""
    img = Image.open(image_path)
    img = ImageOps.exif_transpose(img)
    texts = []
    for i, crop in enumerate(iter_vertical_slices(img, overlap=overlap)):
        buf, name = _image_to_jpeg_buf(crop, f"slice-{i}.jpg")
        texts.append(_post_ocr_buf(buf, name, timeout))
    return join_ocr_slices(texts)


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
    肉类品名为空仍保留：有数量或金额时交人工补品名；只剩印刷序号（粉纸+手影整格认空）
    也留空行，复核页才能对照原图填。跳过重复表头、合计/件数等脏行。
    """
    items = []
    staff_zone = False   # 7 列蔬菜单：物料小计 /「员工：」分区之后的行都在员工区
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
                if is_staff_marker(cells):
                    staff_zone = True
                continue
            # 品名为空但数量、金额都有：手写品名没认出来，保留整行（品名留空），
            # 由存疑规则标红交人工补，不能整行丢掉让金额漏账。
            unit = re.sub(r"[\d.\s×xX]", "", cells[2]) if len(cells) > 2 else ""
            amount = extract_num(cells[5]) if len(cells) > 5 else None
            items.append({"品名": name, "数量": qty, "金额": amount, "类别": "蔬菜", "单位": unit,
                          "票面单价": extract_num(cells[4]) if len(cells) > 4 else None,
                          "员工区": staff_zone})
            continue
        # 序号开头的单据（肉类/销货清单）：5 列，品名在 cells[1]，金额在 cells[4]
        # 注：OCR 偶把数据行序号"1"识别成"序号"，故 first=="序号" 也按数据行处理
        if re.fullmatch(r"\d+", first) or first == "序号":
            name = name_or_raw(cells[1]) if len(cells) > 1 else ""
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
            # 品名为空不丢行。有数量/金额时跟 7 列蔬菜一样留给人工补品名；
            # 数量金额也空、只剩序号时同样留空行（粉纸复写被手影盖住，OCR 只框住表格）。
            items.append({"品名": name, "数量": qty, "金额": amount, "类别": "肉类"})
            continue
        # 品名开头的单据（蔬菜单）：4 列，品名在 cells[0]，数量 cells[1]，金额 cells[3]
        name = name_or_raw(cells[0])
        qty = extract_num(cells[1]) if len(cells) > 1 else None
        amount = extract_num(cells[3]) if len(cells) > 3 else None
        if not name and (qty is None or amount is None):
            continue
        items.append({"品名": name, "数量": qty, "金额": amount, "类别": "蔬菜",
                      "票面单价": extract_num(cells[2]) if len(cells) > 2 else None})
    return items


STAFF_HEADER_RE = re.compile(r"员\s*工")


def is_staff_marker(cells):
    """7 列蔬菜单的员工分区标记行（数量列为空）：
    1) 品名格写「员工」；
    2) 品名格为空或是合计类标签（OCR 常把「员工：」认成「总没」），且金额/备注列有手写数（物料小计，如「计110」「71」「168」）。
    票面写法是：物料段 → 物料小计 →「员工：」→ 员工餐行 → 员工小计 → 总计 → 优惠 → 实付。"""
    if len(cells) < 7:
        return False
    name = (cells[1] or "").strip()
    if extract_num(cells[3] or "") is not None:
        return False
    if STAFF_HEADER_RE.search(name):
        return True
    has_num = any(extract_num((c or "").strip()) is not None for c in cells[5:])
    return has_num and (not name or _is_total_label(name))


def content_item_score(items):
    """(有内容的行数, 已填字段数)。品名、数量、金额都空的占位行不计。"""
    rows_with = 0
    fields = 0
    for it in items or []:
        name = str(it.get("品名") or "").strip()
        has_qty = it.get("数量") not in (None, "")
        has_amt = it.get("金额") not in (None, "")
        n = int(bool(name)) + int(has_qty) + int(has_amt)
        if n:
            rows_with += 1
            fields += n
    return rows_with, fields


def ocr_parse_richer(new_items, new_rows, old_items, old_rows):
    """增强结果是否比当前结果更值得采用。

    先比有内容的明细行，再比填上的字段。两边都没有内容时，保留还认得出表格的那次：
    增强把整张表认丢时，不要盖掉只有序号的骨架（骨架会留成可手填的空行）。
    """
    new_key = content_item_score(new_items) + (1 if new_rows else 0,)
    old_key = content_item_score(old_items) + (1 if old_rows else 0,)
    return new_key > old_key


def count_staff_zone_missing(rows):
    """员工区里「有序号、其余全空」的行数：员工区手写内容疑似没识别出来（如 8/12 被阴影遮住）。"""
    zone = False
    cnt = 0
    for cells in rows:
        if not cells or len(cells) < 7 or is_header_row(cells):
            continue
        cells = [(c or "").strip() for c in cells]
        if not zone:
            zone = is_staff_marker(cells)
            continue
        if re.fullmatch(r"\d+", cells[0]) and not cells[1] and not re.search(r"\d", "".join(cells[2:])):
            cnt += 1
    return cnt


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


# ===== 3.5 美菜 App 截图（纯文本，无 HTML 表） =====
# 实收行：收 ￥金额/数量单位。订/下单数只是对照，缺「收」时不拿它填数量。
_MEICAI_AMT_QTY_RE = re.compile(
    r"[￥¥]\s*(\d+(?:\.\d+)?)\s*[/／]\s*(\d+(?:\.\d+)?)\s*"
    r"(斤|公斤|千克|kg|KG|托|件|箱|包|袋|盒|个|份|板|只|条|瓶|桶|提)?"
)
_MEICAI_PRICE_TAIL_RE = re.compile(
    r"[￥¥]\s*\d+(?:\.\d+)?\s*[/／]\s*\d+(?:\.\d+)?\s*"
    r"[斤公斤千克kgKG托件箱包袋盒个份板只条瓶桶提]*"
)
_MEICAI_UNIT_TAIL_RE = re.compile(
    r"(?:斤|公斤|千克|kg|KG|托(?:\s*[（(][^)）]{0,24}[)）])?|件|箱|包|袋|盒|个|份|板|只|条|瓶|桶|提)\s*$"
)
_MEICAI_COORD_RE = re.compile(r"^\(\d+\s*,\s*\d+\)\s*,\s*\(\d+\s*,\s*\d+\)$")
_MEICAI_SPEC_TOKENS = {
    "普通", "带皮", "去皮", "纯瘦", "优质", "水洗", "中码", "大码", "小码",
    "中号", "大号", "小号", "精品", "特级", "一级", "二级", "精选", "带泥",
    "新鲜", "鲜", "带叶", "去根",
}
_MEICAI_PAY_WORDS = (
    "下单", "支付", "合计", "欠款", "直降", "优惠", "补交", "补款", "补足",
    "多退少补", "总额", "商品总价", "商品金额", "付款", "实付", "签收",
    "司机", "物流", "包裹", "复制", "金额",
)
MEICAI_MISSING_RECV = "未识别到实收（收），请对照原图填写数量和金额，不要用下单数"


def meicai_missing_recv_summary(n):
    return f"{n} 行未识别到实收，请对照原图补数量和金额（不要用下单数）"


def is_meicai_ocr(text):
    """美菜/订单截图的纯文本。手写单是 HTML 表，关键字对不上，仍走原解析。

    拿不准时看这些词：美菜、订单号、送达/配送时间、订/收 成对出现、多退少补。
    分数不够就当手写单，不新开入口。"""
    if not text or not str(text).strip():
        return False
    raw = str(text)
    score = 0
    if "美菜" in raw:
        score += 2
    if re.search(r"订单号\s*[:：]?\s*\d{8,}", raw):
        score += 2
    if re.search(r"(?:送达|配送)时间\s*20\d{2}", raw):
        score += 2
    if re.search(r"(?:^|\n)\s*收\s*[￥¥]", raw) and re.search(r"(?:^|\n)\s*订\s*[￥¥]", raw):
        score += 2
    if "多退少补" in raw or "下单实付" in raw:
        score += 1
    return score >= 4


def _meicai_money(num):
    return float(Decimal(str(num)).quantize(CENT, rounding=ROUND_HALF_UP))


def _clean_meicai_name(raw):
    """去掉店铺前缀、规格套话和价格，留下还能对别名的品名。"""
    s = str(raw or "")
    s = re.sub(r"[\[【][^\]】]{0,40}[\]】]", "", s)
    s = _MEICAI_PRICE_TAIL_RE.sub(" ", s)
    s = re.sub(r"\s+", " ", s).strip()
    s = _MEICAI_UNIT_TAIL_RE.sub("", s).strip()
    tokens = [t for t in s.split() if t]
    kept = []
    for t in tokens:
        if t in _MEICAI_SPEC_TOKENS:
            continue
        kept.append(t)
    if kept:
        def _suffix(token):
            core_tok = re.sub(r"[（(].*?[)）]", "", token)
            return core_tok[-2:] if len(core_tok) >= 2 else core_tok
        base = _suffix(kept[0])
        filtered = [kept[0]]
        for t in kept[1:]:
            if len(base) >= 2 and _suffix(t) == base:
                continue
            filtered.append(t)
        kept = filtered
    name = " ".join(kept).strip()
    if name:
        return name
    fallback = re.sub(r"[\[【][^\]】]{0,40}[\]】]", "", str(raw or ""))
    fallback = _MEICAI_PRICE_TAIL_RE.sub(" ", fallback)
    return re.sub(r"\s+", " ", fallback).strip()


def _meicai_goods_cutoff(line):
    s = line.strip()
    if "付款信息" in s:
        return True
    if re.match(r"共\s*\d+\s*类", s):
        return True
    return False


def _parse_meicai_amt_qty(line):
    m = _MEICAI_AMT_QTY_RE.search(line or "")
    if not m:
        return None
    amt = _meicai_money(m.group(1))
    qty = _meicai_money(m.group(2))
    unit = m.group(3) or ""
    if unit.upper() == "KG":
        unit = "kg"
    return amt, qty, unit


def extract_meicai_date(text):
    m = re.search(
        r"(?:送达|配送)时间\s*(20\d{2})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})",
        text or "",
    )
    if not m:
        return ""
    return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"


def extract_meicai_paid(text):
    """结算后合计。优先「商品多退少补合计」，其次顶栏单独的金额。

    不用下单实付、本单商品支付总额、历史欠款合计。"""
    raw = text or ""
    m = re.search(r"商品多退少补合计\s*[￥¥]\s*(\d+(?:\.\d+)?)", raw)
    if m:
        return _meicai_money(m.group(1))
    m = re.search(r"结算后合计\s*[￥¥]\s*(\d+(?:\.\d+)?)", raw)
    if m:
        return _meicai_money(m.group(1))
    head = raw.split("下单实付", 1)[0]
    m = re.search(r"(?:^|\n)\s*[￥¥]\s*(\d+(?:\.\d+)?)\s*(?:\n|$)", head)
    if m:
        return _meicai_money(m.group(1))
    return None


def extract_meicai_order_ids(text):
    ids = []
    for m in re.finditer(r"订单号\s*[:：]\s*(\d{6,})", text or ""):
        oid = m.group(1)
        if oid not in ids:
            ids.append(oid)
    return ids


def _blank_meicai_item(name, order_id, unit=""):
    return {
        "品名": name or "",
        "数量": None,
        "金额": None,
        "单位": unit or "",
        "类别": "",
        "渠道": "美菜",
        "订单号": order_id or "",
        "实收缺失": False,
        "票面单价": None,
    }


def _meicai_line_kind(line):
    s = re.sub(r"^#+\s*", "", (line or "").strip())
    if not s or _MEICAI_COORD_RE.match(s):
        return "noise", s
    if s in ("复制", "等额收货", "已完成") or s.startswith("您的订单"):
        return "noise", s
    oid = re.search(r"订单号\s*[:：]\s*(\d{6,})", s)
    if oid:
        return "oid", oid.group(1)
    if s.startswith("订"):
        return "order", s
    if s.startswith("收"):
        return "recv", s
    if s.startswith("+"):
        return "noise", s
    if any(k in s for k in _MEICAI_PAY_WORDS) and _parse_meicai_amt_qty(s) is None:
        return "pay", s
    if _parse_meicai_amt_qty(s):
        return "title", s
    if re.search(r"[\u4e00-\u9fff]", s):
        return "name", s
    return "noise", s


def parse_meicai_text(text):
    """美菜订单截图 OCR 文本 → 一张采购单。

    一图里多个订单号仍是一张单，订单号只作溯源。
    明细只认货品的「收」行。直降、历史欠款、补差不单独成行；
    补差已经含在实收金额里。没有「收」就空着数量和金额，标实收缺失，
    不用订货数填。
    """
    date = extract_meicai_date(text)
    paid = extract_meicai_paid(text)
    order_ids = extract_meicai_order_ids(text)
    items = []
    current = None
    pending_name = ""
    current_oid = order_ids[0] if order_ids else ""
    # 只有出现过订单号 / 订 / 收 / 带价格的品名行，才把落单的品名行交给人工补实收。
    # 避免手写单抬头里的中文被误收成一行。
    seen_goods = False

    def _push(item):
        if item is None:
            return
        if item.get("数量") is None or item.get("金额") is None:
            item["数量"] = None
            item["金额"] = None
            item["实收缺失"] = True
        items.append(item)

    for raw_line in str(text or "").splitlines():
        if _meicai_goods_cutoff(raw_line):
            break
        kind, payload = _meicai_line_kind(raw_line)
        if kind in ("noise", "pay"):
            continue
        if kind == "oid":
            seen_goods = True
            current_oid = payload
            continue
        if kind == "order":
            seen_goods = True
            continue
        if kind == "name":
            pending_name = (pending_name + " " + payload).strip() if pending_name else payload
            continue
        if kind == "title":
            seen_goods = True
            _push(current)
            parsed = _parse_meicai_amt_qty(payload)
            unit = parsed[2] if parsed else ""
            current = _blank_meicai_item(_clean_meicai_name(payload), current_oid, unit)
            pending_name = ""
            continue
        if kind == "recv":
            seen_goods = True
            parsed = _parse_meicai_amt_qty(payload)
            if current is None:
                name = _clean_meicai_name(pending_name) if pending_name else ""
                current = _blank_meicai_item(name, current_oid)
                pending_name = ""
            if not parsed:
                current["实收缺失"] = True
                current["数量"] = None
                current["金额"] = None
                continue
            amt, qty, unit = parsed
            current["金额"] = amt
            current["数量"] = qty
            if unit:
                current["单位"] = unit
            current["实收缺失"] = False
            continue
    if current is not None:
        _push(current)
    elif pending_name and seen_goods:
        orphan = _blank_meicai_item(_clean_meicai_name(pending_name), current_oid)
        orphan["实收缺失"] = True
        items.append(orphan)

    suspects = []
    missing = sum(1 for it in items if it.get("实收缺失"))
    if missing:
        suspects.append(meicai_missing_recv_summary(missing))
    if paid is None:
        suspects.append("未识别到结算后合计（商品多退少补合计），请对照原图填写实付")
    note = ("订单号 " + "、".join(order_ids)) if order_ids else ""
    return {
        "kind": "美菜",
        "date": date,
        "paid": paid,
        "order_total": paid,
        "order_ids": order_ids,
        "note": note,
        "items": items,
        "suspects": suspects,
    }


def _meicai_line_sig(items):
    sig = []
    for it in items or []:
        sig.append((
            str(it.get("品名") or ""),
            it.get("数量"),
            it.get("金额"),
            str(it.get("订单号") or ""),
        ))
    return tuple(sig)


def _meicai_order_sig(items, order_id):
    rows = [it for it in items or [] if str(it.get("订单号") or "") == str(order_id)]
    if not rows:
        rows = list(items or [])
    return _meicai_line_sig(rows)


def merge_meicai_docs(docs):
    """同一配送日的多张美菜截图并成一张采购单。

    货品行按截图顺序拼接，不按品名合并。
    实付 = 各张「结算后合计」之和。
    订单号去重。完全相同的一张截图不重复计入。
    同一订单号明细对不上时保留两边的行，并记入 suspects。
    没有配送日的单不和别的单并。
    """
    groups = []
    index = {}
    for doc in docs or []:
        doc = doc or {}
        date = str(doc.get("date") or "")
        if not date:
            groups.append({"date": "", "docs": [doc]})
            continue
        bucket = index.get(date)
        if bucket is None:
            bucket = {"date": date, "docs": []}
            index[date] = bucket
            groups.append(bucket)
        bucket["docs"].append(doc)

    merged = []
    for group in groups:
        seen_orders = {}
        seen_docs = set()
        items = []
        order_ids = []
        suspects = []
        paid = Decimal("0")
        has_paid = False
        for doc in group["docs"]:
            doc_items = list(doc.get("items") or [])
            ids = []
            for oid in doc.get("order_ids") or []:
                oid = str(oid).strip()
                if oid and oid not in ids:
                    ids.append(oid)
            sig = (tuple(ids), _meicai_line_sig(doc_items), doc.get("paid"))
            if sig in seen_docs:
                continue
            seen_docs.add(sig)
            conflicts = []
            for oid in ids:
                osig = _meicai_order_sig(doc_items, oid)
                prev = seen_orders.get(oid)
                if prev is None:
                    seen_orders[oid] = osig
                elif prev != osig:
                    conflicts.append(oid)
                    seen_orders[oid] = osig
            if conflicts:
                suspects.append(
                    "订单号 " + "、".join(conflicts) + " 在多张截图中明细不一致，请核对是否重复入账"
                )
            for oid in ids:
                if oid not in order_ids:
                    order_ids.append(oid)
            items.extend(doc_items)
            if doc.get("paid") not in (None, ""):
                paid += Decimal(str(doc["paid"]))
                has_paid = True
            for tip in doc.get("suspects") or []:
                if tip and tip not in suspects:
                    suspects.append(tip)
        paid_val = float(paid.quantize(CENT, rounding=ROUND_HALF_UP)) if has_paid else None
        note = ("订单号 " + "、".join(order_ids)) if order_ids else ""
        merged.append({
            "kind": "美菜",
            "date": group["date"],
            "paid": paid_val,
            "order_total": paid_val,
            "order_ids": order_ids,
            "note": note,
            "items": items,
            "suspects": suspects,
        })
    return merged


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


def ticket_price_reliable(items, min_rows=3, bad_share=0.5):
    """整张单的票面单价列是否可信。手写单价常被 OCR 整列丢小数点（8/12：3.9→39、5.9→59），
    此时多数行「数量×票面单价」都和金额差约 10 倍，说明错的是单价列而不是数量。
    有票面单价的行不少于 min_rows，且超过 bad_share 比例对不上 → 判整列不可信，
    本单不再按票面单价判数量/金额（历史单价照常判）。"""
    total = bad = 0
    for it in items:
        try:
            qty, amount, ticket = float(it.get("数量")), float(it.get("金额")), float(it.get("票面单价"))
        except (TypeError, ValueError):
            continue
        if qty <= 0 or amount <= 0 or ticket <= 0:
            continue
        total += 1
        if _ratio_off(qty * ticket, amount) > PRICE_RATIO_LIMIT:
            bad += 1
    if total < min_rows:
        return True
    return bad / total <= bad_share


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