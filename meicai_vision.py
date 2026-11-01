# -*- coding: utf-8 -*-
"""美菜电子小票：模力方舟视觉模型 → 复核页明细。

默认模型与对照脚本的回退名相同：Qwen3.5-Flash。
环境变量 MEICAI_VISION_MODEL=Qwen3-VL-Plus 时改用 VL-Plus。别的值（含 Qwen3.6-Plus）仍用 Flash。
密钥与 HunyuanOCR 同一把：先 HUNYUAN_API_KEY，空则 MOLLI_ARK_API_KEY。不另起密钥名。

手写采购单不进这个模块。
Shot3 煲汤袋实收（常见认成 3/14.2，金标是 1/4.8）是已知缺口，这里不改数字。
"""
from __future__ import annotations

import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(ROOT, "scripts")
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)

import vision_meicai_compare as harness

FLASH_MODEL = harness.FALLBACK_MODEL  # Qwen3.5-Flash
PLUS_MODEL = "Qwen3-VL-Plus"
_PLUS_NAMES = {"qwen3-vl-plus", "vl-plus", "plus"}
_FLASH_NAMES = {"", "qwen3.5-flash", "qwen3-vl-flash", "flash"}


def selected_model():
    """复核页美菜视觉模型。未设置或无法识别时是 Qwen3.5-Flash。"""
    raw = os.environ.get("MEICAI_VISION_MODEL", "")
    key = str(raw or "").strip().lower().replace("_", "-").replace(" ", "")
    if key in _PLUS_NAMES:
        return PLUS_MODEL
    if key in _FLASH_NAMES:
        return FLASH_MODEL
    return FLASH_MODEL


def review_doc_from_model_text(text):
    """模型原文 → 复核页文档。不是明细 JSON 时抛错，不把坐标框当成 0 行。"""
    kind, data = harness.classify_model_text(text)
    if kind == "boxes":
        raise ValueError("视觉模型只返回坐标框，未给出复核明细")
    if kind != "items":
        raise ValueError("视觉模型未返回复核明细 JSON")
    return harness.finalize_vision_doc(data)


def recognize(image_path, caller=None, timeout=600):
    """把整张美菜截图交给 chat/completions，返回复核页文档。

    caller 缺省时走对照脚本里的 call_vision（强制 JSON，接口不接受时再去掉 response_format）。
    没有密钥或接口失败时抛错，调用方不要回退到 HunyuanOCR 正文解析。
    """
    key, _source = harness.resolve_api_key()
    if not key:
        raise RuntimeError("未读取到 HUNYUAN_API_KEY 或 MOLLI_ARK_API_KEY，无法识别美菜截图")
    model = selected_model()
    post = caller or harness.call_vision
    raw = post(key, model, image_path, timeout)
    if not isinstance(raw, dict):
        raise RuntimeError("视觉接口没有返回结果")
    if raw.get("error"):
        raise RuntimeError(str(raw.get("error")))
    text = raw.get("text") or ""
    if not str(text).strip():
        raise RuntimeError("视觉模型返回空内容")
    doc = review_doc_from_model_text(text)
    doc["model"] = model
    return doc
