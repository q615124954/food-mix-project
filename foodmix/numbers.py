from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

EXCEL_EPOCH = date(1899, 12, 30)
CENT = Decimal("0.01")
QTY_QUANT = Decimal("0.0001")


def parse_decimal(value) -> Decimal | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, int):
        return Decimal(value)
    if isinstance(value, float):
        return Decimal(str(value))
    text = str(value).strip().replace(",", "").replace("，", "")
    if not text:
        return None
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def dump_decimal(value: Decimal | None) -> str | None:
    if value is None:
        return None
    text = format(value, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


def cell_number(value: Decimal, *, money: bool = False):
    quantum = CENT if money else QTY_QUANT
    quantized = value.quantize(quantum, rounding=ROUND_HALF_UP)
    if quantized == quantized.to_integral_value():
        return int(quantized)
    return float(quantized)


def norm_name(value) -> str:
    return "".join(str(value or "").split())


def norm_date(value) -> str:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value or "").strip()
    if not text:
        return ""
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    return ""


def excel_serial(iso_date: str) -> int | None:
    try:
        day = datetime.strptime(iso_date, "%Y-%m-%d").date()
    except ValueError:
        return None
    return (day - EXCEL_EPOCH).days
