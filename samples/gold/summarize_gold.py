# -*- coding: utf-8 -*-
"""读取金标物料表，统计 2026-08-11 至 08-13 有值的格子并打印摘要。

只依赖 Python 标准库，不调用 OCR，也不读取 API 密钥。
"""
import json
import sys
import zipfile
from datetime import date
from pathlib import Path
from xml.etree import ElementTree as ET

HERE = Path(__file__).resolve().parent
XLSX = HERE / "material_gold.xlsx"
BASELINE = HERE / "baseline_aug11_13.json"
OVERRIDES = HERE / "baseline_overrides.json"
RECEIPTS = (
    "receipt_2026-08-11.jpg",
    "receipt_2026-08-12.jpg",
    "receipt_2026-08-13.jpg",
)
TARGET_DATES = (date(2026, 8, 11), date(2026, 8, 12), date(2026, 8, 13))
EXCEL_EPOCH = date(1899, 12, 30)
TARGET_SERIALS = {(day - EXCEL_EPOCH).days: day for day in TARGET_DATES}
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
JPEG_MAGIC = b"\xff\xd8\xff"


def _fail(message):
    print("FAIL " + message, file=sys.stderr)
    return 1


def _col_index(letters):
    number = 0
    for char in letters:
        number = number * 26 + ord(char) - 64
    return number


def _split_ref(ref):
    letters = "".join(char for char in ref if char.isalpha())
    row = int("".join(char for char in ref if char.isdigit()))
    return _col_index(letters), row


def _cell_text(cell, shared_strings):
    kind = cell.attrib.get("t")
    if kind == "inlineStr":
        return "".join(node.text or "" for node in cell.iter(NS + "t"))
    value = cell.find(NS + "v")
    if value is None or value.text in (None, ""):
        return None
    if kind == "s":
        return shared_strings[int(value.text)]
    return value.text


def _load_shared_strings(archive):
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    strings = []
    for item in root.findall(NS + "si"):
        strings.append("".join(node.text or "" for node in item.iter(NS + "t")))
    return strings


def load_sheets(path):
    """返回 {表名: {(列号, 行号): 文本}}。列号、行号都从 1 开始。"""
    with zipfile.ZipFile(path) as archive:
        shared_strings = _load_shared_strings(archive)
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {}
        for rel in rels:
            targets[rel.attrib.get("Id")] = rel.attrib.get("Target", "").lstrip("/")
        sheets = {}
        for sheet in workbook.findall(NS + "sheets/" + NS + "sheet"):
            rel_id = sheet.attrib.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            )
            target = targets[rel_id]
            if not target.startswith("xl/"):
                target = "xl/" + target
            root = ET.fromstring(archive.read(target))
            grid = {}
            for cell in root.findall(NS + "sheetData/" + NS + "row/" + NS + "c"):
                text = _cell_text(cell, shared_strings)
                if text not in (None, ""):
                    grid[_split_ref(cell.attrib["r"])] = text
            sheets[sheet.attrib["name"]] = grid
    return sheets


def _as_float(text):
    try:
        return float(text)
    except (TypeError, ValueError):
        return None


def _serial_date(text):
    number = _as_float(text)
    if number is None:
        return None
    serial = int(number)
    if abs(number - serial) > 1e-6:
        return None
    return TARGET_SERIALS.get(serial)


def date_groups(grid):
    """日期在表头、下方一格为「数量」的三列（数量/单价/金额）。"""
    groups = []
    for (col, row), text in grid.items():
        found = _serial_date(text)
        if found is None or grid.get((col, row + 1)) != "数量":
            continue
        groups.append((found, col, row))
    return groups


def _name_column(grid, header_row):
    for (col, row), text in grid.items():
        if row == header_row and text == "品名":
            return col
    return 4


def sheet_rows(grid):
    """按日期列出有数量或金额的品名行，并统计三列里有值的格子。"""
    grouped = {day: [] for day in TARGET_DATES}
    counts = {day: 0 for day in TARGET_DATES}
    for found, col, header_row in date_groups(grid):
        name_col = _name_column(grid, header_row)
        rows = sorted(
            {
                row
                for (cell_col, row) in grid
                if col <= cell_col <= col + 2 and row > header_row + 1
            }
        )
        for row in rows:
            values = [grid.get((col + offset, row)) for offset in range(3)]
            filled = [value for value in values if value not in (None, "")]
            counts[found] += len(filled)
            qty, _price, amt = values
            if qty in (None, "") and amt in (None, ""):
                continue
            grouped[found].append(
                {
                    "name": grid.get((name_col, row), ""),
                    "qty": None if qty in (None, "") else float(qty),
                    "amt": None if amt in (None, "") else float(amt),
                }
            )
    return grouped, counts


def _close(left, right):
    if left is None or right is None:
        return left is right
    return abs(float(left) - float(right)) <= 1e-6


def _item_key(item):
    qty = None if item.get("qty") is None else round(float(item["qty"]), 6)
    amt = None if item.get("amt") is None else round(float(item["amt"]), 6)
    return (item.get("sheet", ""), item["date"], item["name"], qty, amt)


def _excel_items(sheets):
    items = []
    for sheet_name in ("肉类采购", "日物料采购"):
        grouped, _counts = sheet_rows(sheets[sheet_name])
        for day, rows in grouped.items():
            for row in rows:
                items.append(
                    {
                        "sheet": sheet_name,
                        "date": day.isoformat(),
                        "name": row["name"],
                        "qty": row["qty"],
                        "amt": row["amt"],
                    }
                )
    return items


def _print_sheet(name, counts, rows):
    print(name)
    total = 0
    for day in TARGET_DATES:
        filled = counts[day]
        total += filled
        names = "、".join(row["name"] for row in rows[day]) or "（无）"
        print(
            "  %s  有值格子 %d  有记录行 %d  %s"
            % (day.isoformat(), filled, len(rows[day]), names)
        )
    print("  小计 %d" % total)
    return total


def _check_baseline(excel_items):
    baseline = json.loads(BASELINE.read_text(encoding="utf-8"))
    expected = []
    for sheet_name in ("肉类采购", "日物料采购"):
        for row in baseline[sheet_name]:
            expected.append(
                {
                    "sheet": sheet_name,
                    "date": row["date"],
                    "name": row["name"],
                    "qty": row["qty"],
                    "amt": row["amt"],
                }
            )
    left = sorted(_item_key(item) for item in excel_items)
    right = sorted(_item_key(item) for item in expected)
    if left != right:
        missing = [item for item in right if item not in left]
        extra = [item for item in left if item not in right]
        print("金标有、baseline 无: %s" % (extra[:5],), file=sys.stderr)
        print("baseline 有、金标无: %s" % (missing[:5],), file=sys.stderr)
        return _fail("material_gold.xlsx 与 baseline_aug11_13.json 的数量/金额不一致")
    meals = baseline["员工餐"]
    if len(meals) != 3:
        return _fail("baseline 员工餐应有 3 天")
    print("baseline_aug11_13.json 日物料/肉类数量与金额和金标表一致，员工餐 %d 天" % len(meals))
    return 0


def _check_overrides():
    payload = json.loads(OVERRIDES.read_text(encoding="utf-8"))
    rows = payload.get("日物料采购") or []
    matched = [
        row
        for row in rows
        if row.get("date") == "2026-08-13" and row.get("name") == "鸡毛菜"
    ]
    if len(matched) != 1:
        return _fail("baseline_overrides.json 应有一条 8/13 鸡毛菜")
    row = matched[0]
    if not _close(row.get("qty"), 5.8) or not _close(row.get("amt"), 28):
        return _fail("overrides 中鸡毛菜应为数量 5.8、金额 28")
    print("baseline_overrides.json 鸡毛菜数量 5.8，金额 28")
    return 0


def _check_receipts():
    for name in RECEIPTS:
        path = HERE / name
        if not path.is_file():
            return _fail("缺少采购单 " + name)
        if path.read_bytes()[:3] != JPEG_MAGIC:
            return _fail(name + " 不是 JPEG")
        print("采购单 %s  %d 字节" % (name, path.stat().st_size))
    return 0


def main():
    for path in (XLSX, BASELINE, OVERRIDES):
        if not path.is_file():
            return _fail("缺少 " + path.name)
    sheets = load_sheets(XLSX)
    for name in ("肉类采购", "日物料采购"):
        if name not in sheets:
            return _fail("金标表缺少工作表 " + name)

    print("金标 %s" % XLSX.name)
    grand = 0
    daily_rows = None
    for name in ("日物料采购", "肉类采购"):
        rows, counts = sheet_rows(sheets[name])
        grand += _print_sheet(name, counts, rows)
        if name == "日物料采购":
            daily_rows = rows
    print("日期列有值格子合计 %d" % grand)

    chicken = [
        row for row in daily_rows[date(2026, 8, 13)] if row["name"] == "鸡毛菜"
    ]
    if len(chicken) != 1 or not _close(chicken[0]["qty"], 5.8) or not _close(chicken[0]["amt"], 28):
        return _fail("金标表 8/13 鸡毛菜应为数量 5.8、金额 28")
    print("8/13 鸡毛菜 数量 5.8  金额 28")

    status = _check_baseline(_excel_items(sheets))
    status = _check_overrides() or status
    status = _check_receipts() or status
    if status:
        return status
    print("样本可读")
    return 0


if __name__ == "__main__":
    sys.exit(main())
