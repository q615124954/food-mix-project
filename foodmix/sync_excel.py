from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from foodmix.errors import AppError
from foodmix.numbers import cell_number, parse_decimal
from foodmix.plan import Write
from foodmix.settings import SheetLayout


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


def apply_plan(
    path: Path,
    writes: list[Write],
    owned: list[dict],
    layout: SheetLayout,
) -> tuple[list[dict], str | None]:
    """把汇总结果写进单元格。新值来自已确认单据的合计，不读取旧值再累加。

    owned 是上次本程序写过的格子。这次合计里没有的，清空。其他格子不动。
    """
    allowed = set(layout.writable_sheets) - set(layout.never_write_sheets)
    for write in writes:
        if write.sheet in layout.never_write_sheets or write.sheet not in allowed:
            raise AppError(f"拒绝写入「{write.sheet}」")
    desired: dict[tuple[str, int, int], int | float] = {}
    for write in writes:
        if write.qty is not None:
            desired[(write.sheet, write.row, write.qty_col)] = cell_number(write.qty, money=False)
        if write.amount is not None:
            desired[(write.sheet, write.row, write.amt_col)] = cell_number(write.amount, money=True)
    clears = []
    for cell in owned:
        key = (str(cell.get("sheet") or ""), int(cell.get("row") or 0), int(cell.get("col") or 0))
        if key[0] in layout.never_write_sheets or key[0] not in allowed:
            continue
        if key not in desired:
            clears.append(key)
    if not desired and not clears:
        return [], None

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    backup = Path(str(path) + f".bak_{stamp}")
    shutil.copy2(path, backup)
    workbook = load_workbook(path)
    before_forbidden = _snapshot_forbidden(workbook, layout)
    try:
        for sheet, row, col in clears:
            if sheet not in workbook.sheetnames:
                continue
            workbook[sheet].cell(row, col).value = None
        for (sheet, row, col), value in desired.items():
            if sheet not in workbook.sheetnames:
                raise AppError(f"物料表缺少「{sheet}」")
            workbook[sheet].cell(row, col).value = value
        workbook.save(path)
    except Exception:
        workbook.close()
        shutil.copy2(backup, path)
        raise
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
