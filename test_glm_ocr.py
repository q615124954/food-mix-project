# -*- coding: utf-8 -*-
import base64, json, sys, os, requests


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

API_KEY = os.environ.get("GLM_API_KEY", "")
URL = "https://open.bigmodel.cn/api/paas/v4/chat/completions"

image_path = r"D:\采购单\微信图片_20260902091938_1729_4.jpg"
with open(image_path, "rb") as f:
    b64 = base64.b64encode(f.read()).decode()

prompt = """你是财务单据录入助手。请识别这张手写采购单图片（图片可能旋转了90度，请先自行判断方向）。
这是一个明细表格，请逐行提取数据并输出 JSON。

要求：
1. 字段：序号、品名、数量(保留原始计算式，如"20×44")、单价、金额、备注
2. 只输出 JSON，不要任何解释或前后缀文字
3. 手写品名尽量辨认；备注（如"货未付""打折""送菜"等）单独放到"备注"字段
4. 如果某行金额与 数量×单价 不符，也照实记录原始数字，不要替我修正

输出格式：
{"明细": [{"序号": "", "品名": "", "数量": "", "单价": "", "金额": "", "备注": ""}]}
"""

payload = {
    "model": "glm-5.3-flash",
    "messages": [
        {
            "role": "user",
            "content": [
                {"type": "image_url", "image_url": {"url": "data:image/jpeg;base64," + b64}},
                {"type": "text", "text": prompt},
            ],
        }
    ],
    "temperature": 0.1,
}

headers = {"Authorization": "Bearer " + API_KEY, "Content-Type": "application/json"}

try:
    resp = requests.post(URL, headers=headers, json=payload, timeout=120)
    print("HTTP", resp.status_code)
    data = resp.json()
    print(json.dumps(data, ensure_ascii=False, indent=2))
except Exception as e:
    print("ERROR:", repr(e))