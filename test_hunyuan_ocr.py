# 测试 HunyuanOCR（Gitee AI 模力方舟）识别手写采购单
# 正确端点: POST /v1/images/ocr (multipart/form-data)
# 注意: 图片最大 2048x2048，超出需先缩小
import requests
import os
import io
from PIL import Image


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

API_URL = "https://ai.gitee.com/v1/images/ocr"
# 混元与模力方舟是同一把密钥：优先 HUNYUAN_API_KEY，为空再用 MOLLI_ARK_API_KEY
API_KEY = (
    os.environ.get("HUNYUAN_API_KEY", "").strip()
    or os.environ.get("MOLLI_ARK_API_KEY", "").strip()
)
MAX_SIZE = 2048

IMAGES = [
    r"d:\采购单\微信图片_20260902091938_1729_4.jpg",
    r"d:\采购单\微信图片_20260902092811_1842_4.jpg",
]

headers = {"Authorization": f"Bearer {API_KEY}"}


def prepare_image(path):
    """若图片超过 2048 则等比缩小，返回 (字节流, 文件名)"""
    img = Image.open(path)
    w, h = img.size
    if max(w, h) > MAX_SIZE:
        scale = MAX_SIZE / max(w, h)
        img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
        print(f"  原图 {w}x{h} -> 缩小为 {img.size[0]}x{img.size[1]}")
    buf = io.BytesIO()
    img.convert("RGB").save(buf, format="JPEG", quality=92)
    buf.seek(0)
    return buf, os.path.basename(path)


for img_path in IMAGES:
    print("=" * 70)
    print(f"图片: {os.path.basename(img_path)}")
    print("=" * 70)
    try:
        buf, name = prepare_image(img_path)
        files = {"image": (name, buf, "image/jpeg")}
        data = {"model": "HunyuanOCR"}
        r = requests.post(API_URL, headers=headers, files=files, data=data, timeout=180)
        print(f"状态码: {r.status_code}")
        if r.status_code == 200:
            result = r.json()
            text = result.get("text_result") or result.get("text") or str(result)
            print("识别结果:")
            print(text)
        else:
            print(f"错误: {r.text[:800]}")
    except Exception as e:
        print(f"异常: {e}")
    print()
