from __future__ import annotations

import re
import shutil
from copy import copy
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import column_index_from_string, get_column_letter

from foodmix.catalog import inspect_dates
from foodmix.errors import AppError
from foodmix.numbers import cell_number, parse_decimal
from foodmix.plan import Write
from foodmix.settings import SheetLayout

# 公式里的 A1 引用。引号内的文字不改，避免把 LOG10 这类函数名当成单元格。
_CELL_REF = re.compile(r'(?<![A-Za-z0-9_])(\$?)([A-Za-z]{1,3})(\$?)(\d+)')
_QUOTED = re.compile(r'("(?:[^"]|"")*")')


def _same(got, expected) -> bool:
    parsed = parse_decimal(got)
    wanted = parse_decimal(expected)
    if parsed is None or wanted is None:
        return got == expected
    return parsed == wanted


def _snapshot_forbidden(workbook, layout: SheetLayout) -> dict[str, list]:
    snap = {}
    for sheet_name in layout.never_write_sheets:
        if sheet_name not in workbook.sheetnames:
            continue
        ws = workbook[sheet_name]
        snap[sheet_name] = [
            [ws.cell(row, col).value for col in range(1, (ws.max_column or 1) + 1)]
            for row in range(1, (ws.max_row or 1) + 1)
        ]
    return snap


def _forbidden_unchanged(path: Path, layout: SheetLayout, before: dict[str, list]) -> None:
    workbook = load_workbook(path, data_only=False)
    try:
        after = _snapshot_forbidden(workbook, layout)
    finally:
        workbook.close()
    if after != before:
        raise AppError("同步改动了一期禁止写入的表，已回滚")


def _shift_formula(formula: str, insert_at: int, width: int) -> str:
    def shift_piece(piece: str) -> str:
        def repl(match: re.Match) -> str:
            col = column_index_from_string(match.group(2).upper())
            if col < insert_at:
                return match.group(0)
            letter = get_column_letter(col + width)
            return f"{match.group(1)}{letter}{match.group(3)}{match.group(4)}"

        return _CELL_REF.sub(repl, piece)

    parts = _QUOTED.split(formula)
    return "".join(part if part.startswith('"') else shift_piece(part) for part in parts)


def _shift_formulas(ws, insert_at: int, width: int) -> None:
    for cell in list(ws._cells.values()):
        if isinstance(cell, MergedCell) or cell.data_type != "f" or not isinstance(cell.value, str):
            continue
        updated = _shift_formula(cell.value, insert_at, width)
        if updated != cell.value:
            cell.value = updated


def _retarget_merges(ws, insert_at: int, width: int) -> None:
    """insert_cols 只挪单元格，合并区域仍指着旧坐标。整段在插入点右侧的跟着右移，跨过插入点的加宽。"""
    kept = []
    for merged in list(ws.merged_cells.ranges):
        ws.merged_cells.remove(merged)
        if merged.min_col >= insert_at:
            merged.shift(col_shift=width)
        elif merged.max_col >= insert_at:
            merged.max_col += width
        kept.append(merged)
    for merged in kept:
        ws.merged_cells.add(merged)


def _merge_covers(ws, row: int, col: int) -> bool:
    for merged in ws.merged_cells.ranges:
        if merged.min_row <= row <= merged.max_row and merged.min_col <= col <= merged.max_col:
            return True
    return False


def _capture_style(cell) -> dict:
    return {
        "font": copy(cell.font),
        "fill": copy(cell.fill),
        "border": copy(cell.border),
        "alignment": copy(cell.alignment),
        "number_format": cell.number_format,
    }


def _apply_style(cell, style: dict) -> None:
    cell.font = style["font"]
    cell.fill = style["fill"]
    cell.border = style["border"]
    cell.alignment = style["alignment"]
    cell.number_format = style["number_format"]


def _subheader_label(offset: int, layout: SheetLayout) -> str:
    if offset == 0:
        return "数量"
    if offset == layout.amount_column_offset:
        return "金额"
    return "单价"


def _insert_date_group(ws, sheet_name: str, serial: int, iso_date: str, layout: SheetLayout) -> int | None:
    """缺这天的日期组时，按日期升序插入三列。已经有则返回 None。"""
    header_row, dates, broken = inspect_dates(ws, sheet_name, layout)
    if serial in dates:
        return None
    if serial in broken:
        raise AppError(f"「{sheet_name}」已有日期 {iso_date}，但下一行不是「数量 / 金额」，无法写入")
    block = layout.amount_column_offset + 1
    if block < 2:
        raise AppError("表布局的金额列偏移无效，无法加列")
    ordered = sorted(dates.values(), key=lambda item: (item.serial, item.qty_col))
    if not ordered:
        raise AppError(f"「{sheet_name}」找不到可对齐的日期列，无法自动加列")
    columns = [item.qty_col for item in ordered]
    if columns != sorted(columns):
        raise AppError(f"「{sheet_name}」的日期列不是从左到右按日期排列，拒绝加列以免拆开原有日期")
    for prev, nxt in zip(columns, columns[1:]):
        if nxt < prev + block:
            raise AppError(f"「{sheet_name}」的日期列重叠，拒绝加列")
    insert_at = None
    template = ordered[-1]
    for item in ordered:
        if item.serial > serial:
            insert_at = item.qty_col
            template = item
            break
    if insert_at is None:
        insert_at = ordered[-1].qty_col + block
    occupied = [item.qty_col for item in ordered] + list(broken.values())
    for start in occupied:
        if start < insert_at < start + block:
            raise AppError(f"「{sheet_name}」加列会拆开已有的日期列，已停止")

    sub = header_row + 1
    header_style = _capture_style(ws.cell(header_row, template.qty_col))
    sub_styles = [_capture_style(ws.cell(sub, template.qty_col + offset)) for offset in range(block)]
    widths = []
    for offset in range(block):
        letter = get_column_letter(template.qty_col + offset)
        widths.append(ws.column_dimensions[letter].width if letter in ws.column_dimensions else None)

    ws.insert_cols(insert_at, block)
    _retarget_merges(ws, insert_at, block)
    _shift_formulas(ws, insert_at, block)
    for offset in range(block):
        if _merge_covers(ws, header_row, insert_at + offset) or _merge_covers(ws, sub, insert_at + offset):
            raise AppError(f"「{sheet_name}」日期 {iso_date} 的加列位置落在合并单元格里，无法写入")

    header_cell = ws.cell(header_row, insert_at, serial)
    _apply_style(header_cell, header_style)
    for offset in range(block):
        label_cell = ws.cell(sub, insert_at + offset, _subheader_label(offset, layout))
        _apply_style(label_cell, sub_styles[offset])
        if widths[offset] is not None:
            ws.column_dimensions[get_column_letter(insert_at + offset)].width = widths[offset]
    ws.merge_cells(
        start_row=header_row,
        start_column=insert_at,
        end_row=header_row,
        end_column=insert_at + block - 1,
    )
    return insert_at


def _map_col(col: int, insertions: list[tuple[int, int]]) -> int:
    for insert_at, width in insertions:
        if col >= insert_at:
            col += width
    return col


def _assign(cell_owner, sheet: str, row: int, col: int, value) -> None:
    cell = cell_owner.cell(row, col)
    if isinstance(cell, MergedCell):
        raise AppError(f"「{sheet}」第{row}行第{col}列是合并单元格，无法写入")
    cell.value = value


def apply_plan(
    path: Path,
    writes: list[Write],
    owned: list[dict],
    layout: SheetLayout,
) -> tuple[list[dict], str | None]:
    """把汇总结果写进单元格。新值来自已确认单据的合计，不读取旧值再累加。

    目标表还没有该日期时，先按日期顺序补一组三列，再写数量和金额。单价列留空。
    owned 是上次本程序写过的格子。这次合计里没有的，清空。其他格子不动。
    """
    allowed = set(layout.writable_sheets) - set(layout.never_write_sheets)
    for write in writes:
        if write.sheet in layout.never_write_sheets or write.sheet not in allowed:
            raise AppError(f"拒绝写入「{write.sheet}」")
    pending_owned: list[tuple[str, int, int]] = []
    for cell in owned:
        key = (str(cell.get("sheet") or ""), int(cell.get("row") or 0), int(cell.get("col") or 0))
        if key[0] in layout.never_write_sheets or key[0] not in allowed:
            continue
        pending_owned.append(key)
    if not writes and not pending_owned:
        return [], None

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    backup = Path(str(path) + f".bak_{stamp}")
    shutil.copy2(path, backup)
    workbook = None
    try:
        workbook = load_workbook(path)
        before_forbidden = _snapshot_forbidden(workbook, layout)
        block = layout.amount_column_offset + 1
        needed: dict[str, dict[int, str]] = {}
        for write in writes:
            if write.qty_col is not None and write.amt_col is not None:
                continue
            if write.sheet in layout.never_write_sheets or write.sheet not in allowed:
                raise AppError(f"拒绝写入「{write.sheet}」")
            if not write.serial:
                raise AppError(f"「{write.sheet}」日期 {write.date} 无法对应到列")
            needed.setdefault(write.sheet, {}).setdefault(write.serial, write.date)
        insertions: dict[str, list[tuple[int, int]]] = {}
        for sheet, markers in needed.items():
            if sheet not in workbook.sheetnames:
                raise AppError(f"物料表缺少「{sheet}」")
            insertions[sheet] = []
            for serial in sorted(markers):
                inserted = _insert_date_group(workbook[sheet], sheet, serial, markers[serial], layout)
                if inserted is not None:
                    insertions[sheet].append((inserted, block))
        resolved: dict[str, dict] = {}
        for sheet in needed:
            _header, dates, _broken = inspect_dates(workbook[sheet], sheet, layout)
            resolved[sheet] = dates

        desired: dict[tuple[str, int, int], int | float] = {}
        for write in writes:
            if write.qty_col is None or write.amt_col is None:
                found = resolved.get(write.sheet, {}).get(write.serial or -1)
                if found is None:
                    raise AppError(f"「{write.sheet}」未能加上日期 {write.date} 的数量/金额列")
                qty_col, amt_col = found.qty_col, found.amt_col
            else:
                shifted = insertions.get(write.sheet, [])
                qty_col = _map_col(write.qty_col, shifted)
                amt_col = _map_col(write.amt_col, shifted)
            if write.qty is not None:
                desired[(write.sheet, write.row, qty_col)] = cell_number(write.qty, money=False)
            if write.amount is not None:
                desired[(write.sheet, write.row, amt_col)] = cell_number(write.amount, money=True)
        clears = []
        for sheet, row, col in pending_owned:
            key = (sheet, row, _map_col(col, insertions.get(sheet, [])))
            if key not in desired:
                clears.append(key)

        for sheet, row, col in clears:
            if sheet not in workbook.sheetnames:
                continue
            _assign(workbook[sheet], sheet, row, col, None)
        for (sheet, row, col), value in desired.items():
            if sheet not in workbook.sheetnames:
                raise AppError(f"物料表缺少「{sheet}」")
            _assign(workbook[sheet], sheet, row, col, value)
        workbook.save(path)
    except Exception:
        if workbook is not None:
            workbook.close()
            workbook = None
        shutil.copy2(backup, path)
        raise
    if workbook is not None:
        workbook.close()
    try:
        _verify(path, desired, clears)
        _forbidden_unchanged(path, layout, before_forbidden)
    except Exception:
        shutil.copy2(backup, path)
        raise
    owned_now = [{"sheet": sheet, "row": row, "col": col} for (sheet, row, col) in desired]
    return owned_now, str(backup)


def _verify(path: Path, desired: dict, clears: list[tuple[str, int, int]]) -> None:
    workbook = load_workbook(path, data_only=False)
    try:
        for (sheet, row, col), value in desired.items():
            if not _same(workbook[sheet].cell(row, col).value, value):
                raise AppError(f"回读校验失败：{sheet} 第{row}行第{col}列")
        for sheet, row, col in clears:
            if sheet not in workbook.sheetnames:
                continue
            if workbook[sheet].cell(row, col).value not in (None, ""):
                raise AppError(f"回读校验失败：应清空的单元格仍有值（{sheet} 第{row}行第{col}列）")
    finally:
        workbook.close()
