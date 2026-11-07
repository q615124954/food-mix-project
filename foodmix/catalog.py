from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

from foodmix.errors import AppError
from foodmix.numbers import EXCEL_EPOCH, norm_name
from foodmix.settings import SheetLayout


@dataclass(frozen=True)
class DateCol:
    serial: int
    qty_col: int
    amt_col: int


@dataclass(frozen=True)
class Item:
    sheet: str
    row: int
    code: str
    name: str
    unit: str
    key: str


@dataclass
class ParsedSheet:
    sheet: str
    header_row: int
    dates: dict[int, DateCol]
    items: list[Item] = field(default_factory=list)


@dataclass
class Catalog:
    items: list[Item]
    sheets: dict[str, ParsedSheet]
    missing_sheets: list[str]
    source: str

    def find(self, sheet: str, name: str) -> Item | None:
        key = norm_name(name)
        if not key:
            return None
        for item in self.items:
            if item.sheet == sheet and item.key == key:
                return item
        return None

    def names_for(self, category: str, layout: SheetLayout) -> list[Item]:
        sheet = layout.category_sheets.get(category)
        if not sheet:
            return []
        return [item for item in self.items if item.sheet == sheet]


def _serial_of(value) -> int | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, datetime):
        return (value.date() - EXCEL_EPOCH).days
    if isinstance(value, date):
        return (value - EXCEL_EPOCH).days
    if isinstance(value, (int, float)) and 40000 < float(value) < 60000:
        return int(value)
    return None


def _date_cells(ws, row: int) -> list[tuple[int, int]]:
    found = []
    for col in range(1, (ws.max_column or 0) + 1):
        serial = _serial_of(ws.cell(row, col).value)
        if serial is not None:
            found.append((col, serial))
    return found


def _parse_sheet(ws, sheet_name: str, layout: SheetLayout) -> ParsedSheet:
    header_row = None
    for row in range(1, min(ws.max_row or 1, 40) + 1):
        labels = [str(ws.cell(row, col).value or "").strip() for col in range(1, 7)]
        if "品名" in labels and _date_cells(ws, row):
            header_row = row
            break
    if header_row is None:
        raise AppError(f"「{sheet_name}」找不到同时带品名和日期列的表头")
    dates: dict[int, DateCol] = {}
    sub = header_row + 1
    for col, serial in _date_cells(ws, header_row):
        qty_label = str(ws.cell(sub, col).value or "").strip()
        amt_col = col + layout.amount_column_offset
        amt_label = str(ws.cell(sub, amt_col).value or "").strip()
        if qty_label != "数量" or amt_label != "金额":
            continue
        dates[serial] = DateCol(serial=serial, qty_col=col, amt_col=amt_col)
    items: list[Item] = []
    seen: set[str] = set()
    for row in range(header_row + 2, (ws.max_row or 0) + 1):
        raw_name = ws.cell(row, layout.name_column).value
        if raw_name is None:
            continue
        name = str(raw_name).strip()
        if not name or name in {"品名", "名称"} or "合计" in name:
            continue
        key = norm_name(name)
        if key in seen:
            continue
        seen.add(key)
        code_value = ws.cell(row, layout.code_column).value
        unit_value = ws.cell(row, layout.unit_column).value
        items.append(Item(
            sheet=sheet_name,
            row=row,
            code="" if code_value is None else str(code_value).strip(),
            name=name,
            unit="" if unit_value is None else str(unit_value).strip(),
            key=key,
        ))
    return ParsedSheet(sheet=sheet_name, header_row=header_row, dates=dates, items=items)


def load_catalog(path: Path, layout: SheetLayout) -> Catalog:
    if not path.is_file():
        raise AppError(f"物料明细表不存在：{path}")
    if path.suffix.lower() != ".xlsx":
        raise AppError("物料明细表必须是 .xlsx")
    workbook = load_workbook(path, data_only=False)
    try:
        items: list[Item] = []
        sheets: dict[str, ParsedSheet] = {}
        missing: list[str] = []
        for sheet_name in layout.writable_sheets:
            if sheet_name in layout.never_write_sheets:
                continue
            if sheet_name not in workbook.sheetnames:
                missing.append(sheet_name)
                continue
            parsed = _parse_sheet(workbook[sheet_name], sheet_name, layout)
            sheets[sheet_name] = parsed
            items.extend(parsed.items)
        return Catalog(items=items, sheets=sheets, missing_sheets=missing, source=str(path))
    finally:
        workbook.close()
