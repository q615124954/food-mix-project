from __future__ import annotations

import json
import re

from foodmix.errors import AppError
from foodmix.numbers import norm_date, parse_decimal

ALLOWED_CATEGORY = {"蔬菜", "肉类"}


def parse_json_blob(text: str) -> dict:
    raw = str(text or "").strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        start = raw.find("{")
        end = raw.rfind("}")
        if start < 0 or end <= start:
            raise AppError("视觉模型没有返回 JSON")
        try:
            data = json.loads(raw[start:end + 1])
        except json.JSONDecodeError as exc:
            raise AppError("视觉模型返回的 JSON 无法解析") from exc
    if not isinstance(data, dict):
        raise AppError("视觉模型返回的不是 JSON 对象")
    return data


def _line_category(row: dict) -> str:
    if "category" in row:
        raw = row.get("category")
    elif "行类别" in row:
        raw = row.get("行类别")
    else:
        raw = row.get("类别", "")
    category = str(raw or "").strip()
    if category not in ALLOWED_CATEGORY:
        return ""
    return category


def parse_extraction(text: str) -> dict:
    """解析视觉模型的行。整单上的类别字段不会套到每一行。"""
    data = parse_json_blob(text)
    rows = data.get("lines")
    if rows is None:
        rows = data.get("明细") or []
    if not isinstance(rows, list):
        raise AppError("视觉模型的 lines 不是数组")
    lines = []
    for index, row in enumerate(rows, start=1):
        if not isinstance(row, dict):
            continue
        name = row.get("name") if "name" in row else row.get("品名")
        unit = row.get("unit") if "unit" in row else row.get("单位")
        qty_raw = row.get("qty") if "qty" in row else row.get("数量")
        amount_raw = row.get("amount") if "amount" in row else row.get("金额")
        qty = parse_decimal(qty_raw)
        amount = parse_decimal(amount_raw)
        issues = []
        if not str(name or "").strip():
            issues.append("缺少品名")
        if qty_raw not in (None, "") and qty is None:
            issues.append("数量无法解析")
        if amount_raw not in (None, "") and amount is None:
            issues.append("金额无法解析")
        lines.append({
            "line_no": index,
            "raw_name": str(name or "").strip(),
            "qty": qty,
            "unit": str(unit or "").strip(),
            "amount": amount,
            "category": _line_category(row),
            "issues": issues,
        })
    return {
        "date": norm_date(data.get("date", data.get("日期"))),
        "supplier": str(data.get("supplier") or data.get("供应商") or "").strip(),
        "lines": lines,
    }
