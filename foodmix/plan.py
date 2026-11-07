from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal

from foodmix.catalog import Catalog
from foodmix.numbers import excel_serial
from foodmix.settings import SheetLayout
from foodmix.units import UnitRule, convert_quantity


@dataclass(frozen=True)
class BillLine:
    invoice_id: str
    date: str
    raw_name: str
    name: str
    qty: Decimal | None
    unit: str
    amount: Decimal | None
    category: str


@dataclass(frozen=True)
class Skip:
    invoice_id: str
    date: str
    raw_name: str
    name: str
    category: str
    reason: str
    sheet: str = ""


@dataclass
class Write:
    sheet: str
    row: int
    qty_col: int
    amt_col: int
    name: str
    code: str
    unit: str
    date: str
    qty: Decimal | None
    amount: Decimal | None
    invoice_ids: list[str] = field(default_factory=list)


def _skip(line: BillLine, reason: str, sheet: str = "") -> Skip:
    return Skip(line.invoice_id, line.date, line.raw_name, line.name, line.category, reason, sheet)


def build_plan(
    lines: list[BillLine],
    catalog: Catalog,
    rules: list[UnitRule],
    layout: SheetLayout,
) -> tuple[list[Write], list[Skip]]:
    """按行类别选表，再按日期 + 物料行汇总。不读单元格旧值。"""
    buckets: dict[tuple, dict] = {}
    order: list[tuple] = []
    skipped: list[Skip] = []
    for line in lines:
        sheet = layout.category_sheets.get(line.category or "")
        if line.category not in {"蔬菜", "肉类"} or not sheet or sheet in layout.never_write_sheets:
            if line.category not in {"蔬菜", "肉类"}:
                reason = "行类别不是蔬菜或肉类，未写入"
            else:
                reason = f"表「{sheet or line.category}」一期不写"
            skipped.append(_skip(line, reason, sheet or ""))
            continue
        if sheet not in catalog.sheets:
            skipped.append(_skip(line, f"物料表缺少「{sheet}」", sheet))
            continue
        if not line.date:
            skipped.append(_skip(line, "缺少日期，未写入", sheet))
            continue
        item = catalog.find(sheet, line.name)
        if item is None:
            label = line.name or line.raw_name or "（空品名）"
            skipped.append(_skip(line, f"「{label}」在「{sheet}」中找不到，跳过不写", sheet))
            continue
        serial = excel_serial(line.date)
        date_col = catalog.sheets[sheet].dates.get(serial) if serial is not None else None
        if date_col is None:
            skipped.append(_skip(line, f"「{sheet}」没有日期 {line.date} 的数量/金额列，未写入", sheet))
            continue
        converted = convert_quantity(item.name, line.qty, line.unit, item.unit, rules)
        if not converted.ok:
            skipped.append(_skip(line, converted.reason, sheet))
            continue
        key = (sheet, item.row, date_col.qty_col)
        bucket = buckets.get(key)
        if bucket is None:
            bucket = {
                "sheet": sheet,
                "row": item.row,
                "qty_col": date_col.qty_col,
                "amt_col": date_col.amt_col,
                "name": item.name,
                "code": item.code,
                "unit": item.unit,
                "date": line.date,
                "qty": None,
                "amount": None,
                "invoice_ids": [],
            }
            buckets[key] = bucket
            order.append(key)
        if converted.qty is not None:
            bucket["qty"] = converted.qty if bucket["qty"] is None else bucket["qty"] + converted.qty
        if line.amount is not None:
            bucket["amount"] = line.amount if bucket["amount"] is None else bucket["amount"] + line.amount
        if line.invoice_id not in bucket["invoice_ids"]:
            bucket["invoice_ids"].append(line.invoice_id)
    writes = [Write(**buckets[key]) for key in order]
    return writes, skipped
