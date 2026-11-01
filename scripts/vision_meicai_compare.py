#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""美菜电子小票：视觉模型对照实验。

只给人工跑。不接入 webapp，不改复核页默认。
手写采购单仍走 HunyuanOCR。本脚本的基线也是 HunyuanOCR
`/v1/images/ocr` 加 pipeline_min.parse_meicai_text。

视觉模型走模力方舟 `POST /v1/chat/completions`，图片用 OpenAI 式
image_url。要求模型直接交复核页那套明细 JSON。
若 OCR 模型只回坐标框，标「接口形态不合格」，不把框换算成 0 命中。
这种情况下可再打一发回退模型，默认名 Qwen3.5-Flash。

没有 HUNYUAN_API_KEY / MOLLI_ARK_API_KEY 时不发请求。
对照表仍打印，行命中、实付一致、编造行、耗时留空。
"""
from __future__ import annotations

import argparse
import base64
import json
import os
import re
import sys
import time
from decimal import Decimal, ROUND_HALF_UP

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import pipeline_min as core
import requests

SCHEMA = "vision_meicai_compare.v1"
DEFAULT_MODELS = ("Qwen3.5-OCR", "Qwen3-VL-Plus")
FALLBACK_MODEL = "Qwen3.5-Flash"
BASELINE_MODEL = "HunyuanOCR"
LABEL_BAD_INTERFACE = "接口形态不合格"
CENT = Decimal("0.01")

# 复核页明细（pipeline_min._blank_meicai_item + derive_price）。
VISION_PROMPT = """你在录入一张美菜电子小票。只输出一个 JSON 对象，不要 Markdown，不要解释。

明细只认货品的「收」（实收数量和实收金额）。不要把「订」的数量和金额当成实收。
直降、历史欠款、补差、广告语不要单独成行。补差已经含在实收金额里。
没有「收」的货品仍要列出：数量和金额用 null，实收缺失 true。不要用订货数填空。
paid 用「商品多退少补合计」或「结算后合计」。没有则 null。不要用下单实付。

字段与复核页货品行一致：
{
  "date": "YYYY-MM-DD 或空串",
  "paid": null,
  "order_ids": [],
  "items": [
    {
      "品名": "",
      "数量": null,
      "金额": null,
      "单位": "",
      "类别": "",
      "渠道": "美菜",
      "订单号": "",
      "实收缺失": false,
      "票面单价": null
    }
  ]
}
"""

_FENCE_RE = re.compile(r"^```(?:json)?\s*|\s*```$", re.I)
_COORD_RE = re.compile(
    r"\(\s*-?\d+(?:\.\d+)?\s*,\s*-?\d+(?:\.\d+)?\s*\)"
    r"|\[\s*-?\d+(?:\.\d+)?(?:\s*,\s*-?\d+(?:\.\d+)?){3,}\s*\]"
)
_BOX_FIELDS = (
    "bbox", "box", "polygon", "poly", "points",
    "bounding_box", "boundingBox", "quad", "position",
)
_BOX_LIST_KEYS = (
    "boxes", "bboxes", "detections", "regions", "words", "layout",
    "ocr", "ocr_result", "result", "data", "output", "lines", "blocks", "cells",
)
_ITEM_NAME_KEYS = ("品名", "name", "商品", "商品名")


def resolve_api_key():
    """与 pipeline_min 相同：先 HUNYUAN_API_KEY，空则 MOLLI_ARK_API_KEY。

    返回 (key, source_name)。没有密钥时 key 为空串。不打印密钥。
    """
    core._load_env()
    hunyuan = os.environ.get("HUNYUAN_API_KEY", "").strip()
    molli = os.environ.get("MOLLI_ARK_API_KEY", "").strip()
    if hunyuan:
        return hunyuan, "HUNYUAN_API_KEY"
    if molli:
        return molli, "MOLLI_ARK_API_KEY"
    return "", ""


def default_pairs():
    exp = os.path.join(ROOT, "samples", "vision_exp")
    gold = os.path.join(ROOT, "samples", "gold")
    return [
        (os.path.join(exp, "shot1.jpg"), os.path.join(gold, "meicai_shot1.json")),
        (os.path.join(exp, "shot2.jpg"), os.path.join(gold, "meicai_shot2.json")),
        (os.path.join(exp, "shot3.jpg"), os.path.join(exp, "fixture_shot3.json")),
    ]


def display_path(path):
    try:
        rel = os.path.relpath(path, ROOT)
    except ValueError:
        return path
    if rel.startswith(".."):
        return path
    return rel.replace(os.sep, "/")


def num_or_none(value):
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "").replace("￥", "").replace("¥", "")
    if not text or text.lower() in {"null", "none", "无", "-", "nan"}:
        return None
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    return float(match.group())


def as_bool(value, default):
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return bool(value)
    text = str(value).strip().lower()
    if text in {"1", "true", "yes", "是"}:
        return True
    if text in {"0", "false", "no", "否"}:
        return False
    return default


def cents(value):
    if value is None:
        return None
    return Decimal(str(value)).quantize(CENT, rounding=ROUND_HALF_UP)


def same_cents(left, right):
    if left is None and right is None:
        return True
    if left is None or right is None:
        return False
    return cents(left) == cents(right)


def norm_name(value):
    return re.sub(r"\s+", " ", str(value or "").strip())


def norm_unit(value):
    unit = str(value or "").strip()
    if unit.upper() == "KG":
        return "kg"
    return unit


def units_compatible(left, right):
    a, b = norm_unit(left), norm_unit(right)
    if not a or not b:
        return True
    return a == b


def clip(text, limit=400):
    raw = re.sub(r"\s+", " ", str(text or "")).strip()
    if len(raw) <= limit:
        return raw
    return raw[:limit] + "…"


def extract_json(text):
    if not text or not str(text).strip():
        return None
    raw = _FENCE_RE.sub("", str(text).strip()).strip()
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        pass
    decoder = json.JSONDecoder()
    for opener in ("{", "["):
        start = raw.find(opener)
        if start < 0:
            continue
        try:
            obj, _end = decoder.raw_decode(raw[start:])
            return obj
        except json.JSONDecodeError:
            continue
    return None


def _looks_like_item(obj):
    return isinstance(obj, dict) and any(key in obj for key in _ITEM_NAME_KEYS)


def _is_box_object(obj):
    if not isinstance(obj, dict) or _looks_like_item(obj):
        return False
    return any(key in obj for key in _BOX_FIELDS)


def is_boxes_payload(data):
    """坐标框接口：有框、没有复核页品名行。"""
    if isinstance(data, list):
        dicts = [item for item in data if isinstance(item, dict)]
        if not dicts or any(_looks_like_item(item) for item in dicts):
            return False
        boxed = sum(1 for item in dicts if _is_box_object(item))
        return boxed >= 1 and boxed >= int(0.6 * len(dicts))
    if not isinstance(data, dict) or _looks_like_item(data):
        return False
    if any(key in data for key in _BOX_FIELDS) and "items" not in data and "明细" not in data:
        return True
    for key in _BOX_LIST_KEYS:
        value = data.get(key)
        if isinstance(value, list) and value and is_boxes_payload(value):
            return True
    return False


def _items_are_boxes(items):
    dicts = [item for item in items if isinstance(item, dict)]
    if not items or not dicts:
        return False
    if any(_looks_like_item(item) for item in dicts):
        return False
    return all(_is_box_object(item) for item in dicts)


def looks_like_coordinate_dump(text):
    raw = str(text or "")
    if "品名" in raw or '"items"' in raw:
        return False
    return len(_COORD_RE.findall(raw)) >= 3


def classify_model_text(text):
    """返回 (kind, data)。kind 为 items / boxes / text / other。"""
    data = extract_json(text)
    if isinstance(data, dict):
        for key in ("items", "明细"):
            items = data.get(key)
            if isinstance(items, list):
                if _items_are_boxes(items):
                    return "boxes", data
                return "items", data
    if data is not None and is_boxes_payload(data):
        return "boxes", data
    if isinstance(data, list) and any(_looks_like_item(item) for item in data):
        return "items", {"items": data}
    if isinstance(data, dict) and _looks_like_item(data):
        return "items", {"items": [data]}
    # 纯数字数组也会被 json 解出来，但仍是坐标而不是货品行。
    if looks_like_coordinate_dump(text):
        return "boxes", data
    if data is None:
        return "text", None
    return "other", data


def doc_paid(data):
    if not isinstance(data, dict):
        return None
    for key in ("paid", "order_total", "实付", "结算后合计"):
        if key in data and data.get(key) not in ("", None):
            return num_or_none(data.get(key))
    return None


def doc_order_ids(data):
    raw = data.get("order_ids")
    if raw is None:
        raw = data.get("订单号") or []
    if isinstance(raw, str):
        raw = [raw]
    ids = []
    for item in raw or []:
        text = str(item).strip()
        if text and text not in ids:
            ids.append(text)
    return ids


def coerce_item(raw):
    if not isinstance(raw, dict):
        return None
    name = ""
    for key in _ITEM_NAME_KEYS:
        if raw.get(key) not in (None, ""):
            name = raw.get(key)
            break
    qty = num_or_none(raw.get("数量", raw.get("qty", raw.get("实收数量"))))
    amount = num_or_none(raw.get("金额", raw.get("amt", raw.get("amount", raw.get("实收金额")))))
    missing_default = qty is None or amount is None
    if "实收缺失" in raw:
        missing = as_bool(raw.get("实收缺失"), missing_default)
    else:
        missing = missing_default
    return {
        "品名": norm_name(name),
        "数量": qty,
        "金额": amount,
        "单位": norm_unit(raw.get("单位", raw.get("unit"))),
        "类别": str(raw.get("类别") or ""),
        "渠道": str(raw.get("渠道") or "美菜"),
        "订单号": str(raw.get("订单号") or raw.get("order_id") or ""),
        "实收缺失": missing,
        "票面单价": num_or_none(raw.get("票面单价")),
    }


def finalize_vision_doc(data):
    """模型 JSON → 与 parse_meicai_text 同一套货品行，再按品名+单位合并。"""
    data = data or {}
    items = []
    for raw in data.get("items") or []:
        item = coerce_item(raw)
        if item is not None:
            items.append(item)
    items = core.aggregate_meicai_items(items)
    paid = doc_paid(data)
    date = str(data.get("date") or data.get("delivery_date") or "")
    order_ids = doc_order_ids(data)
    suspects = []
    missing = sum(1 for item in items if item.get("实收缺失"))
    if missing:
        suspects.append(core.meicai_missing_recv_summary(missing))
    if paid is None:
        suspects.append("未识别到结算后合计（商品多退少补合计），请对照原图填写实付")
    doc = {
        "kind": "美菜",
        "date": date,
        "paid": paid,
        "order_total": paid,
        "order_ids": order_ids,
        "note": ("订单号 " + "、".join(order_ids)) if order_ids else "",
        "items": items,
        "suspects": suspects,
    }
    return stamp_review_items(doc)


def baseline_doc_from_text(text):
    """当前美菜解析。品名清洗、别名、按品名合并都在 parse_meicai_text 里。"""
    parsed = core.parse_meicai_text(text or "")
    core.derive_price(parsed["items"])
    return stamp_review_items(parsed)


def stamp_review_items(doc):
    date = doc.get("date") or ""
    for item in doc.get("items") or []:
        item["日期"] = date
        if not item.get("渠道"):
            item["渠道"] = "美菜"
    return doc


def load_gold(path):
    with open(path, encoding="utf-8") as handle:
        data = json.load(handle)
    if not isinstance(data, dict):
        raise ValueError("金标 JSON 须是对象")
    rows = data.get("日物料采购_复核后期望行")
    if isinstance(rows, list):
        source = rows
        lines = [_gold_line(row, review_row=True) for row in source]
    else:
        lines = [_gold_line(row, review_row=False) for row in (data.get("items") or [])]
    return {
        "path": path,
        "lines": lines,
        "paid": doc_paid(data),
        "date": str(data.get("date") or data.get("delivery_date") or ""),
        "order_ids": doc_order_ids(data),
    }


def _gold_line(row, review_row):
    row = row or {}
    if review_row:
        qty = row.get("qty") if "qty" in row else row.get("数量")
        amount = row.get("amt") if "amt" in row else row.get("金额")
        name = row.get("name") if "name" in row else row.get("品名")
        unit = row.get("unit") if "unit" in row else row.get("单位")
    else:
        qty = row.get("数量") if "数量" in row else row.get("qty")
        amount = row.get("金额") if "金额" in row else row.get("amt")
        name = row.get("品名") if "品名" in row else row.get("name")
        unit = row.get("单位") if "单位" in row else row.get("unit")
    return {
        "品名": norm_name(name),
        "单位": norm_unit(unit),
        "数量": num_or_none(qty),
        "金额": num_or_none(amount),
    }


def line_has_substance(line):
    if norm_name(line.get("品名")):
        return True
    return line.get("数量") is not None or line.get("金额") is not None


def lines_match(gold_line, pred_line):
    if norm_name(gold_line.get("品名")) != norm_name(pred_line.get("品名")):
        return False
    if not units_compatible(gold_line.get("单位"), pred_line.get("单位")):
        return False
    return (
        same_cents(gold_line.get("数量"), pred_line.get("数量"))
        and same_cents(gold_line.get("金额"), pred_line.get("金额"))
    )


def public_line(line):
    return {
        "品名": norm_name(line.get("品名")),
        "数量": num_or_none(line.get("数量")),
        "金额": num_or_none(line.get("金额")),
        "单位": norm_unit(line.get("单位")),
    }


def compare_items(items, pred_paid, gold):
    """一对一：品名相同、单位不冲突、数量和金额到分一致算命中。

    没对上且有品名或有数的预测行计为编造。空行不计。
    """
    if gold is None:
        return {
            "line_hits": None,
            "gold_lines": None,
            "missed_lines": None,
            "paid_match": None,
            "gold_paid": None,
            "fabricated_lines": None,
            "fabricated": None,
            "missed": None,
        }
    used = set()
    missed = []
    hits = 0
    for gold_line in gold["lines"]:
        found = None
        for index, pred in enumerate(items or []):
            if index in used:
                continue
            if lines_match(gold_line, pred):
                found = index
                break
        if found is None:
            missed.append(public_line(gold_line))
        else:
            used.add(found)
            hits += 1
    fabricated = []
    for index, pred in enumerate(items or []):
        if index in used or not line_has_substance(pred):
            continue
        fabricated.append(public_line(pred))
    if gold.get("paid") is None or pred_paid is None:
        paid_match = None if gold.get("paid") is None else False
    else:
        paid_match = bool(same_cents(gold.get("paid"), pred_paid))
    return {
        "line_hits": hits,
        "gold_lines": len(gold["lines"]),
        "missed_lines": len(missed),
        "paid_match": paid_match,
        "gold_paid": gold.get("paid"),
        "fabricated_lines": len(fabricated),
        "fabricated": fabricated,
        "missed": missed,
    }


def try_load_gold(path):
    if not path:
        return None, "未提供金标路径"
    if not os.path.isfile(path):
        return None, "金标文件不存在"
    try:
        return load_gold(path), ""
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        return None, "金标无法读取: %s" % exc


def shell_run(role, model, endpoint, status, gold, note="", error="", route=""):
    return {
        "role": role,
        "model": model,
        "endpoint": endpoint,
        "status": status,
        "label": LABEL_BAD_INTERFACE if status == LABEL_BAD_INTERFACE else "",
        "interface": None,
        "line_hits": None,
        "gold_lines": None if gold is None else len(gold["lines"]),
        "missed_lines": None,
        "paid": None,
        "gold_paid": None if gold is None else gold.get("paid"),
        "paid_match": None,
        "fabricated_lines": None,
        "latency_ms": None,
        "items": None,
        "fabricated": None,
        "missed": None,
        "http_status": None,
        "error": error or "",
        "note": note or "",
        "route": route or "",
        "raw_excerpt": "",
    }


def apply_score(row, doc, gold):
    metrics = compare_items(doc.get("items") or [], doc.get("paid"), gold)
    row.update(metrics)
    row["paid"] = doc.get("paid")
    row["items"] = doc.get("items") or []
    row["interface"] = "items"
    row["status"] = "ok"
    row["label"] = ""
    return row


def hunyuan_meicai_text(image_path, timeout):
    """复核页同一条选择：窄长图先原宽分条，整图「收」多出至少 3 行才换整图。

    两次都是 HunyuanOCR `/v1/images/ocr`。这里不改 pipeline_min.MODEL。
    """
    if core.looks_like_long_screenshot(image_path):
        tiled = ""
        try:
            tiled = core.call_ocr_tiled(image_path, timeout=timeout) or ""
        except Exception:
            tiled = ""
        if str(tiled).strip():
            full = ""
            try:
                full = core.call_ocr(image_path, timeout=timeout) or ""
            except Exception:
                full = ""
            if core.shrunk_full_beats_tiled(tiled, full):
                return full, "hunyuan_full_frame"
            return tiled, "hunyuan_tiled"
    text = core.call_ocr(image_path, timeout=timeout) or ""
    return text, "hunyuan_full_frame"


def call_baseline(image_path, timeout, api_key):
    if api_key:
        core.API_KEY = api_key
    started = time.perf_counter()
    try:
        text, route = hunyuan_meicai_text(image_path, timeout)
        return {
            "text": text,
            "route": route,
            "latency_ms": int((time.perf_counter() - started) * 1000),
            "error": "",
        }
    except Exception as exc:
        return {
            "text": "",
            "route": "",
            "latency_ms": int((time.perf_counter() - started) * 1000),
            "error": clip(exc, 300),
        }


def _mime(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".png":
        return "image/png"
    if ext == ".webp":
        return "image/webp"
    if ext == ".gif":
        return "image/gif"
    return "image/jpeg"


def image_data_url(path):
    with open(path, "rb") as handle:
        payload = base64.b64encode(handle.read()).decode("ascii")
    return "data:%s;base64,%s" % (_mime(path), payload)


def vision_body(model, image_path, force_json):
    body = {
        "model": model,
        "temperature": 0,
        "max_tokens": 4096,
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "image_url",
                        "image_url": {"url": image_data_url(image_path)},
                    },
                    {"type": "text", "text": VISION_PROMPT},
                ],
            }
        ],
    }
    if force_json:
        body["response_format"] = {"type": "json_object"}
    return body


def response_format_rejected(status, body):
    if status not in (400, 422):
        return False
    low = (body or "").lower()
    return "response_format" in low or "json_object" in low


def message_text(payload):
    if not isinstance(payload, dict):
        return ""
    choices = payload.get("choices") or []
    if not choices or not isinstance(choices[0], dict):
        return ""
    message = choices[0].get("message") or {}
    content = message.get("content")
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts = []
        for part in content:
            if isinstance(part, str):
                parts.append(part)
            elif isinstance(part, dict):
                parts.append(str(part.get("text") or part.get("content") or ""))
        return "\n".join(part for part in parts if part)
    if content is None:
        return ""
    return str(content)


def call_vision(api_key, model, image_path, timeout):
    """OpenAI 式多模态 chat/completions。失败不抛出，交给对照行。"""
    started = time.perf_counter()
    url = core.CHAT_URL
    headers = {
        "Authorization": "Bearer " + api_key,
        "Content-Type": "application/json",
    }
    force_json = True
    try:
        response = requests.post(
            url,
            headers=headers,
            json=vision_body(model, image_path, force_json),
            timeout=timeout,
        )
        if response_format_rejected(response.status_code, response.text):
            force_json = False
            response = requests.post(
                url,
                headers=headers,
                json=vision_body(model, image_path, False),
                timeout=timeout,
            )
        latency = int((time.perf_counter() - started) * 1000)
        if response.status_code != 200:
            return {
                "text": "",
                "latency_ms": latency,
                "http_status": response.status_code,
                "error": "HTTP %s: %s" % (response.status_code, clip(response.text, 300)),
                "json_mode": force_json,
            }
        try:
            payload = response.json()
        except ValueError:
            return {
                "text": "",
                "latency_ms": latency,
                "http_status": response.status_code,
                "error": "响应不是 JSON",
                "json_mode": force_json,
            }
        return {
            "text": message_text(payload),
            "latency_ms": latency,
            "http_status": response.status_code,
            "error": "",
            "json_mode": force_json,
        }
    except Exception as exc:
        return {
            "text": "",
            "latency_ms": int((time.perf_counter() - started) * 1000),
            "http_status": None,
            "error": clip(exc, 300),
            "json_mode": force_json,
        }


def _join_note(*parts):
    return " ".join(part for part in parts if part)


def run_baseline_row(image, gold, skip_status, timeout, api_key, caller):
    note = "基线：HunyuanOCR /v1/images/ocr + parse_meicai_text。不改生产默认。"
    if skip_status:
        extra = "未发请求。"
        if not os.path.isfile(image):
            extra += " 图片不存在。"
        return shell_run(
            "baseline", BASELINE_MODEL, core.API_URL, skip_status, gold,
            note=_join_note(note, extra),
        )
    if not os.path.isfile(image):
        return shell_run(
            "baseline", BASELINE_MODEL, core.API_URL, "missing_image", gold,
            note=_join_note(note, "图片不存在，未发请求。"),
        )
    raw = caller(image, timeout, api_key)
    row = shell_run("baseline", BASELINE_MODEL, core.API_URL, "error", gold, note=note)
    row["latency_ms"] = raw.get("latency_ms")
    row["route"] = raw.get("route") or ""
    row["raw_excerpt"] = clip(raw.get("text") or "")
    row["error"] = raw.get("error") or ""
    if raw.get("error"):
        row["note"] = _join_note(note, "调用失败。")
        return row
    doc = baseline_doc_from_text(raw.get("text") or "")
    apply_score(row, doc, gold)
    row["route"] = raw.get("route") or ""
    row["note"] = _join_note(note, "route=%s" % (row["route"] or "unknown"))
    if gold is None:
        row["note"] = _join_note(row["note"], "无金标，不计行命中。")
    if not core.is_meicai_ocr(raw.get("text") or "") and not doc.get("items"):
        row["note"] = _join_note(row["note"], "文本未判成美菜。")
    return row


def interpret_vision(raw, gold):
    if raw.get("error"):
        return "error", None, raw.get("error") or ""
    text = raw.get("text") or ""
    if not str(text).strip():
        return "error", None, "空回复"
    kind, data = classify_model_text(text)
    if kind == "boxes":
        return LABEL_BAD_INTERFACE, None, "只返回坐标框，未按复核页明细 JSON 计分。"
    if kind != "items":
        return "非约定JSON", None, "回复里没有 items 明细。不把自由文本当成命中。"
    return "ok", finalize_vision_doc(data), ""


def run_vision_row(model, role, image, gold, skip_status, timeout, api_key, caller, note):
    if skip_status:
        extra = "未发请求。"
        if not os.path.isfile(image):
            extra += " 图片不存在。"
        return shell_run(
            role, model, core.CHAT_URL, skip_status, gold, note=_join_note(note, extra),
        )
    if not os.path.isfile(image):
        return shell_run(
            role, model, core.CHAT_URL, "missing_image", gold,
            note=_join_note(note, "图片不存在，未发请求。"),
        )
    raw = caller(api_key, model, image, timeout)
    row = shell_run(role, model, core.CHAT_URL, "error", gold, note=note)
    row["latency_ms"] = raw.get("latency_ms")
    row["http_status"] = raw.get("http_status")
    row["raw_excerpt"] = clip(raw.get("text") or "")
    status, doc, detail = interpret_vision(raw, gold)
    row["status"] = status
    row["error"] = raw.get("error") or ""
    if status == "ok":
        apply_score(row, doc, gold)
        row["note"] = note if gold is not None else _join_note(note, "无金标，不计行命中。")
        return row
    if status == LABEL_BAD_INTERFACE:
        row["label"] = LABEL_BAD_INTERFACE
        row["interface"] = "boxes"
        row["note"] = _join_note(note, detail)
        return row
    if status == "非约定JSON":
        row["interface"] = classify_model_text(raw.get("text") or "")[0]
        row["note"] = _join_note(note, detail)
        return row
    row["note"] = _join_note(note, detail)
    return row


def ocr_model_name(model):
    return "ocr" in str(model or "").lower()


def run_report(
    pairs,
    models,
    api_key="",
    key_source="",
    dry_run=False,
    fallback_model=FALLBACK_MODEL,
    use_fallback=True,
    include_baseline=True,
    ocr_timeout=600,
    vision_timeout=180,
    baseline_caller=None,
    vision_caller=None,
):
    """跑对照。无密钥或 dry_run 时不调用 caller。"""
    if baseline_caller is None:
        baseline_caller = call_baseline
    if vision_caller is None:
        vision_caller = call_vision
    if not api_key:
        skip_status = "skipped_no_api_key"
    elif dry_run:
        skip_status = "dry_run"
    else:
        skip_status = ""

    samples = []
    any_live = False
    for image, gold_path in pairs:
        gold, gold_error = try_load_gold(gold_path)
        runs = []
        if include_baseline:
            runs.append(run_baseline_row(
                image, gold, skip_status, ocr_timeout, api_key, baseline_caller,
            ))
        box_models = []
        for model in models:
            row = run_vision_row(
                model, "vision", image, gold, skip_status, vision_timeout, api_key,
                vision_caller,
                note="视觉：chat/completions，图片走 image_url，要求复核页 JSON。",
            )
            runs.append(row)
            if row["status"] == LABEL_BAD_INTERFACE and ocr_model_name(model):
                box_models.append(model)
        runs.append(_fallback_row(
            image, gold, skip_status, vision_timeout, api_key, vision_caller,
            fallback_model, use_fallback, box_models, models,
        ))
        if any(run.get("latency_ms") is not None for run in runs):
            any_live = True
        samples.append({
            "image": display_path(image),
            "gold": display_path(gold_path),
            "gold_error": gold_error,
            "runs": runs,
        })

    return {
        "schema": SCHEMA,
        "experiment": "meicai_electronic_receipt_vision_compare",
        "production_default_model": core.MODEL,
        "vision_is_production_default": False,
        "handwriting_ocr": BASELINE_MODEL,
        "ocr_url": core.API_URL,
        "chat_url": core.CHAT_URL,
        "api_key_present": bool(api_key),
        "key_source": key_source if api_key else "",
        "contains_live_measurements": any_live,
        "models": list(models),
        "fallback_model": fallback_model,
        "fallback_enabled": bool(use_fallback),
        "note": _report_note(bool(api_key), bool(dry_run), any_live),
        "samples": samples,
    }


def _fallback_row(
    image, gold, skip_status, timeout, api_key, caller,
    fallback_model, use_fallback, box_models, models,
):
    slot = "回退槽 %s。仅当 OCR 模型返回坐标框时才请求。" % fallback_model
    if skip_status:
        return shell_run(
            "fallback", fallback_model, core.CHAT_URL, skip_status, gold,
            note=_join_note(slot, "未发请求。"),
        )
    if not use_fallback:
        return shell_run(
            "fallback", fallback_model, core.CHAT_URL, "not_run", gold,
            note=_join_note(slot, "已关闭回退。"),
        )
    if not box_models:
        return shell_run(
            "fallback", fallback_model, core.CHAT_URL, "not_run", gold,
            note=_join_note(slot, "没有 OCR 模型返回坐标框。"),
        )
    if fallback_model in models:
        return shell_run(
            "fallback", fallback_model, core.CHAT_URL, "not_run", gold,
            note=_join_note(slot, "该模型已在对照名单中，不重复调用。"),
        )
    triggered = "、".join(box_models)
    return run_vision_row(
        fallback_model, "fallback", image, gold, "", timeout, api_key, caller,
        note="回退 %s，因 %s %s。" % (fallback_model, triggered, LABEL_BAD_INTERFACE),
    )


def _report_note(has_key, dry_run, any_live):
    if not has_key:
        return (
            "未读取到 HUNYUAN_API_KEY 或 MOLLI_ARK_API_KEY，未调用模力方舟。"
            "行命中、实付一致、编造行、耗时留空，不是实测 0。"
            "生产复核仍走 HunyuanOCR。实跑对照由运维持密钥后重跑本脚本填写，仓库不预写数字。"
        )
    if dry_run:
        return "dry-run：不发请求。对照数字留空，不是实测 0。"
    if not any_live:
        return "这次没有完成在线调用。空着的格子不是实测 0。"
    return (
        "含在线调用结果。坐标框只标接口形态不合格，不换算成行命中。"
        "手写单不在本实验范围，仍走 HunyuanOCR。"
    )


def fmt_cell(value):
    if value is None:
        return ""
    return str(value)


def fmt_bool(value):
    if value is None:
        return ""
    return "是" if value else "否"


def render_markdown(report):
    lines = [
        "# 美菜电子小票视觉模型对照",
        "",
        report.get("note") or "",
        "",
        "生产默认模型仍是 **%s**（`/v1/images/ocr`）。本表来自实验脚本，不会写回复核页。"
        % report.get("production_default_model"),
        "",
    ]
    if not report.get("contains_live_measurements"):
        lines.append("本次没有在线测量。行命中、实付一致、编造行、耗时留空，不是 0。")
        lines.append("")
    lines.append("| 样本 | 角色 | 模型 | 状态 | 行命中 | 金标行 | 实付一致 | 编造行 | 耗时ms |")
    lines.append("| --- | --- | --- | --- | --- | --- | --- | --- | --- |")
    for sample in report.get("samples") or []:
        name = os.path.basename(sample.get("image") or "")
        for run in sample.get("runs") or []:
            lines.append("| %s |" % " | ".join([
                name,
                str(run.get("role") or ""),
                str(run.get("model") or ""),
                str(run.get("status") or ""),
                fmt_cell(run.get("line_hits")),
                fmt_cell(run.get("gold_lines")),
                fmt_bool(run.get("paid_match")),
                fmt_cell(run.get("fabricated_lines")),
                fmt_cell(run.get("latency_ms")),
            ]))
    lines.append("")
    lines.append(
        "schema: `%s`。金标行数和 JSON 里的 gold_paid 来自金标文件，不是模型输出。"
        % report.get("schema")
    )
    lines.append("")
    return "\n".join(lines)


def write_text(path, text):
    directory = os.path.dirname(os.path.abspath(path))
    if directory:
        os.makedirs(directory, exist_ok=True)
    with open(path, "w", encoding="utf-8") as handle:
        handle.write(text)


def parse_pairs(args):
    raw = list(args or [])
    if not raw:
        return default_pairs()
    if len(raw) % 2 != 0:
        raise SystemExit("图片和金标须成对出现：image gold image gold ...")
    pairs = []
    for index in range(0, len(raw), 2):
        pairs.append((raw[index], raw[index + 1]))
    return pairs


def build_models(extra, no_default):
    models = [] if no_default else list(DEFAULT_MODELS)
    for name in extra or []:
        name = str(name).strip()
        if name and name not in models:
            models.append(name)
    if not models:
        raise SystemExit("模型名单为空。去掉 --no-default-models，或用 --model 指定。")
    return models


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="美菜电子小票视觉模型对照（实验脚本，不改生产默认）",
    )
    parser.add_argument(
        "pairs", nargs="*",
        help="成对路径：图片 金标。省略则用 samples/vision_exp 的三张 Shot",
    )
    parser.add_argument(
        "--model", action="append", default=[],
        help="在默认 Qwen3.5-OCR、Qwen3-VL-Plus 之外追加模型，可重复",
    )
    parser.add_argument(
        "--no-default-models", action="store_true",
        help="不用默认视觉模型，只跑 --model",
    )
    parser.add_argument(
        "--fallback-model", default=FALLBACK_MODEL,
        help="OCR 模型只回坐标框时改打的模型，默认 Qwen3.5-Flash",
    )
    parser.add_argument(
        "--no-fallback", action="store_true",
        help="坐标框结果不再请求回退模型",
    )
    parser.add_argument("--no-baseline", action="store_true", help="不跑 HunyuanOCR 基线")
    parser.add_argument(
        "--dry-run", action="store_true",
        help="不发请求，只输出空对照表（有密钥也不调用）",
    )
    parser.add_argument("--json", dest="json_path", help="把 JSON 对照写到这个路径")
    parser.add_argument("--md", dest="md_path", help="把 Markdown 对照写到这个路径")
    parser.add_argument("--timeout", type=float, default=180, help="视觉请求超时秒数")
    parser.add_argument("--ocr-timeout", type=float, default=600, help="HunyuanOCR 超时秒数")
    args = parser.parse_args(argv)

    api_key, key_source = resolve_api_key()
    report = run_report(
        parse_pairs(args.pairs),
        build_models(args.model, args.no_default_models),
        api_key=api_key,
        key_source=key_source,
        dry_run=args.dry_run,
        fallback_model=args.fallback_model,
        use_fallback=not args.no_fallback,
        include_baseline=not args.no_baseline,
        ocr_timeout=args.ocr_timeout,
        vision_timeout=args.timeout,
    )
    markdown = render_markdown(report)
    sys.stdout.write(markdown)
    if args.json_path:
        write_text(args.json_path, json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    if args.md_path:
        write_text(args.md_path, markdown)
    return 0


if __name__ == "__main__":
    sys.exit(main())
