import json
from datetime import datetime
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from foodmix.errors import AppError
from foodmix.numbers import excel_serial
from foodmix.service import Workspace
from foodmix.settings import state_path
from foodmix.store import read_state, write_state
from tests.test_sync import _invoice, _line, _plant, _use_tmp

AUG1 = "2026-08-01"
AUG20 = "2026-08-20"
JUL15 = "2026-07-15"
SEP2 = "2026-09-02"
SEP3 = "2026-09-03"
SEP5 = "2026-09-05"


def _bill(name, day, lines, status="confirmed"):
    invoice = _invoice(name, lines, status=status)
    invoice["date"] = day
    return invoice


def _block(ws, header_row, col, serial):
    ws.cell(header_row, col, serial)
    ws.merge_cells(start_row=header_row, start_column=col, end_row=header_row, end_column=col + 2)
    for offset, label in enumerate(("数量", "单价", "金额")):
        ws.cell(header_row + 1, col + offset, label)


def _veg(ws, row, code, name):
    ws.cell(row, 2, code)
    ws.cell(row, 4, name)
    ws.cell(row, 5, "斤")


def _fingerprint(path, sheet):
    workbook = load_workbook(path, data_only=False)
    try:
        ws = workbook[sheet]
        values = {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.value not in (None, ""):
                    values[(cell.row, cell.column)] = cell.value
        merges = sorted(
            (item.min_row, item.min_col, item.max_row, item.max_col)
            for item in ws.merged_cells.ranges
        )
        return values, merges
    finally:
        workbook.close()


def _open(path):
    return load_workbook(path, data_only=False)


def _count(ws, row, serial):
    return sum(1 for col in range(1, (ws.max_column or 1) + 1) if ws.cell(row, col).value == serial)


def _august(path, *, daily_dates, meat_dates=None, tail=False):
    """八月矩阵。daily_dates / meat_dates 是 (日期, 起始列)。"""
    workbook = Workbook()
    daily = workbook.active
    daily.title = "日物料采购"
    daily["A1"] = "序号"
    daily["B1"] = "物料代码"
    daily["C1"] = "类别"
    daily["D1"] = "品名"
    daily["E1"] = "单位"
    daily["F1"] = "合计"
    for col, label in ((6, "数量"), (7, "单价"), (8, "金额")):
        daily.cell(2, col, label)
    for day, col in daily_dates:
        _block(daily, 1, col, excel_serial(day))
    _veg(daily, 3, "V-葱", "香葱")
    _veg(daily, 4, "V-菜", "香菜")
    _veg(daily, 5, "V-鸡", "鸡毛菜")
    _veg(daily, 6, "V-小", "小葱")
    if tail:
        daily.cell(1, 15, "备注")
        daily.cell(3, 15, "TAIL")

    meat = workbook.create_sheet("肉类采购")
    meat["A1"] = "肉类采购订货汇总"
    meat.merge_cells("A1:N1")
    meat["A2"] = "序号"
    meat["B2"] = "物料代码"
    meat["D2"] = "品名"
    meat["E2"] = "单位"
    meat["F2"] = "合计"
    for day, col in (meat_dates if meat_dates is not None else [(AUG1, 9)]):
        _block(meat, 2, col, excel_serial(day))
    meat["B4"] = "M-1"
    meat["D4"] = "五花肉"
    meat["E4"] = "斤"
    meat.cell(4, 9, 6)
    meat.cell(4, 10, 9.5)
    meat.cell(4, 11, 60)
    meat.cell(4, 12, "MOVE")

    staff = workbook.create_sheet("员工餐")
    staff["A1"] = "员工餐"
    staff["B3"] = excel_serial(AUG1)
    staff["D3"] = 12345
    staff["H3"] = 888

    company = workbook.create_sheet("公司物料采购")
    company["A1"] = "公司物料采购"
    company["B2"] = excel_serial(SEP2)
    company["C3"] = "COMPANY"

    summary = workbook.create_sheet("综合表")
    summary["A1"] = "综合表"
    summary["B2"] = "SUMMARY"
    workbook.save(path)
    return path


def _sync(monkeypatch, tmp_path, material, invoices):
    _use_tmp(monkeypatch, tmp_path)
    workspace = Workspace()
    workspace.update_settings("", str(material))
    _plant(invoices)
    return workspace, workspace.sync()


def test_september_second_inserts_a_column_and_writes_the_greens(tmp_path, monkeypatch):
    material = _august(
        tmp_path / "八月.xlsx",
        daily_dates=[(AUG1, 9), (AUG20, 12)],
        tail=True,
    )
    daily = _open(material)["日物料采购"]
    # 八月已有列：同步前后必须还在原位，不能串到 9/2。
    daily["F3"] = "=I3"
    daily["H3"] = "KEEPAMT"
    daily.cell(3, 9, 100)
    daily.cell(3, 10, 2)
    daily.cell(3, 11, 999)
    daily.cell(3, 12, 11)
    daily.cell(3, 13, 3)
    daily.cell(3, 14, 22)
    daily.cell(4, 9, 20)
    daily.cell(4, 10, 1)
    daily.cell(4, 11, 40)
    daily.cell(4, 12, 5)
    daily.cell(4, 14, 9)
    daily.cell(5, 9, 7)
    daily.cell(5, 11, 14)
    daily.cell(5, 12, 8)
    daily.cell(5, 14, 16)
    daily.cell(6, 9, 50)
    daily.cell(6, 11, 51)
    _open_book = daily.parent
    _open_book.save(material)
    _open_book.close()

    untouched = {
        name: _fingerprint(material, name)
        for name in ("肉类采购", "员工餐", "公司物料采购", "综合表")
    }
    before = {}
    workbook = _open(material)
    sheet = workbook["日物料采购"]
    for row in range(1, 8):
        for col in range(1, 15):
            before[(row, col)] = sheet.cell(row, col).value
    workbook.close()

    workspace, result = _sync(monkeypatch, tmp_path, material, [
        _bill("a.jpg", SEP2, [
            _line(1, "香葱", "香葱", "2", "斤", "5", "蔬菜"),
            _line(2, "香菜", "香菜", "4", "斤", "8", "蔬菜"),
            _line(3, "鸡毛菜", "鸡毛菜", "6", "斤", "15", "蔬菜"),
        ]),
        _bill("b.jpg", SEP2, [
            _line(1, "香葱", "香葱", "3", "斤", "7", "蔬菜"),
        ]),
        _bill("c.jpg", SEP2, [
            _line(1, "香菜", "香菜", "100", "斤", "100", "蔬菜"),
        ], status="recognized"),
    ])
    written = {row["name"]: row for row in result["written"]}
    assert written["香葱"]["qty"] == "5" and written["香葱"]["amount"] == "12"
    assert written["香菜"]["qty"] == "4" and written["香菜"]["amount"] == "8"
    assert written["鸡毛菜"]["qty"] == "6" and written["鸡毛菜"]["amount"] == "15"
    assert all(row["sheet"] == "日物料采购" for row in result["written"])
    assert result["staff_meal_written"] is False
    assert not any(row["name"] in {"香葱", "香菜", "鸡毛菜"} for row in result["skipped"])

    workbook = _open(material)
    sheet = workbook["日物料采购"]
    serial = excel_serial(SEP2)
    assert _count(sheet, 1, serial) == 1
    assert sheet.cell(1, 15).value == serial
    assert [sheet.cell(2, col).value for col in (15, 16, 17)] == ["数量", "单价", "金额"]
    assert (1, 15, 1, 17) in {
        (item.min_row, item.min_col, item.max_row, item.max_col) for item in sheet.merged_cells.ranges
    }
    assert sheet.cell(3, 15).value == 5 and sheet.cell(3, 16).value is None and sheet.cell(3, 17).value == 12
    assert sheet.cell(4, 15).value == 4 and sheet.cell(4, 16).value is None and sheet.cell(4, 17).value == 8
    assert sheet.cell(5, 15).value == 6 and sheet.cell(5, 16).value is None and sheet.cell(5, 17).value == 15
    assert sheet.cell(6, 15).value is None and sheet.cell(6, 17).value is None
    for key, value in before.items():
        assert sheet.cell(*key).value == value
    assert sheet.cell(1, 9).value == excel_serial(AUG1)
    assert sheet.cell(1, 12).value == excel_serial(AUG20)
    assert sheet.cell(1, 18).value == "备注"
    assert sheet.cell(3, 18).value == "TAIL"
    assert sheet.cell(3, 9).value == 100 and sheet.cell(3, 11).value == 999
    assert sheet["F3"].value == "=I3"
    assert sheet["H3"].value == "KEEPAMT"
    workbook.close()
    for name, fingerprint in untouched.items():
        assert _fingerprint(material, name) == fingerprint

    workbook = _open(material)
    workbook["日物料采购"].cell(3, 15).value = 999
    workbook.save(material)
    workbook.close()
    workspace.sync()
    workbook = _open(material)
    sheet = workbook["日物料采购"]
    assert _count(sheet, 1, serial) == 1
    assert sheet.cell(3, 15).value == 5
    assert sheet.cell(3, 9).value == 100
    assert sheet.max_column == 18
    workbook.close()

    state = read_state(state_path())
    for invoice in state["invoices"]:
        if invoice["id"] == "a.jpg":
            invoice["lines"] = [line for line in invoice["lines"] if line["name"] != "鸡毛菜"]
    write_state(state_path(), state)
    workspace._state = None
    workspace.sync()
    workbook = _open(material)
    sheet = workbook["日物料采购"]
    assert sheet.cell(5, 15).value is None and sheet.cell(5, 17).value is None
    assert sheet.cell(3, 15).value == 5 and sheet.cell(3, 17).value == 12
    assert sheet.cell(4, 15).value == 4
    assert sheet.cell(3, 9).value == 100 and sheet.cell(3, 12).value == 11
    assert _fingerprint(material, "员工餐") == untouched["员工餐"]
    workbook.close()


def test_meat_rows_add_a_column_only_on_the_meat_sheet(tmp_path, monkeypatch):
    material = _august(tmp_path / "八月.xlsx", daily_dates=[(AUG1, 9)])
    before_daily = _fingerprint(material, "日物料采购")
    before_staff = _fingerprint(material, "员工餐")
    before_company = _fingerprint(material, "公司物料采购")
    before_summary = _fingerprint(material, "综合表")
    _sync(monkeypatch, tmp_path, material, [
        _bill("meat.jpg", SEP2, [_line(1, "五花肉", "五花肉", "2", "斤", "30", "肉类")]),
    ])
    assert _fingerprint(material, "日物料采购") == before_daily
    assert _fingerprint(material, "员工餐") == before_staff
    assert _fingerprint(material, "公司物料采购") == before_company
    assert _fingerprint(material, "综合表") == before_summary
    workbook = _open(material)
    meat = workbook["肉类采购"]
    serial = excel_serial(SEP2)
    assert _count(meat, 2, serial) == 1
    assert meat.cell(2, 12).value == serial
    assert [meat.cell(3, col).value for col in (12, 13, 14)] == ["数量", "单价", "金额"]
    assert meat.cell(4, 12).value == 2 and meat.cell(4, 13).value is None and meat.cell(4, 14).value == 30
    assert meat.cell(4, 9).value == 6 and meat.cell(4, 10).value == 9.5 and meat.cell(4, 11).value == 60
    assert meat.cell(4, 15).value == "MOVE"
    merges = {(item.min_row, item.min_col, item.max_row, item.max_col) for item in meat.merged_cells.ranges}
    assert (2, 9, 2, 11) in merges
    assert (2, 12, 2, 14) in merges
    assert (1, 1, 1, 17) in merges
    workbook.close()


def test_new_date_lands_between_existing_dates_without_mixing_them(tmp_path, monkeypatch):
    material = _august(tmp_path / "八月.xlsx", daily_dates=[(AUG1, 9), (SEP5, 12)], tail=False)
    workbook = _open(material)
    sheet = workbook["日物料采购"]
    sheet["A8"] = "=L3"
    sheet["F3"] = "=I3"
    sheet.cell(3, 9, 100)
    sheet.cell(3, 10, 2)
    sheet.cell(3, 11, 999)
    sheet.cell(3, 12, 1)
    sheet.cell(3, 13, 1.25)
    sheet.cell(3, 14, 2)
    workbook.save(material)
    workbook.close()

    workspace, _first = _sync(monkeypatch, tmp_path, material, [
        _bill("later.jpg", SEP5, [_line(1, "香葱", "香葱", "8", "斤", "9", "蔬菜")]),
    ])
    workbook = _open(material)
    sheet = workbook["日物料采购"]
    assert sheet.cell(3, 9).value == 100
    assert sheet.cell(3, 12).value == 8 and sheet.cell(3, 13).value == 1.25 and sheet.cell(3, 14).value == 9
    workbook.close()

    state = read_state(state_path())
    state["invoices"] = [_bill("mid.jpg", SEP2, [_line(1, "香葱", "香葱", "4", "斤", "6", "蔬菜")])]
    write_state(state_path(), state)
    workspace._state = None
    workspace.sync()

    workbook = _open(material)
    sheet = workbook["日物料采购"]
    serial = excel_serial(SEP2)
    assert _count(sheet, 1, serial) == 1
    assert sheet.cell(1, 9).value == excel_serial(AUG1)
    assert sheet.cell(1, 12).value == serial
    assert sheet.cell(1, 15).value == excel_serial(SEP5)
    assert sheet.cell(3, 9).value == 100 and sheet.cell(3, 10).value == 2 and sheet.cell(3, 11).value == 999
    assert sheet.cell(3, 12).value == 4 and sheet.cell(3, 13).value is None and sheet.cell(3, 14).value == 6
    assert sheet.cell(3, 15).value is None and sheet.cell(3, 16).value == 1.25 and sheet.cell(3, 17).value is None
    assert sheet["F1"].value == "合计" and sheet["F3"].value == "=I3"
    assert sheet["A8"].value == "=O3"
    merges = {(item.min_row, item.min_col, item.max_row, item.max_col) for item in sheet.merged_cells.ranges}
    assert (1, 9, 1, 11) in merges
    assert (1, 12, 1, 14) in merges
    assert (1, 15, 1, 17) in merges
    workbook.close()


def test_earlier_date_inserts_before_the_block_and_keeps_the_total(tmp_path, monkeypatch):
    material = _august(tmp_path / "八月.xlsx", daily_dates=[(AUG1, 9)])
    workbook = _open(material)
    sheet = workbook["日物料采购"]
    sheet["F3"] = "=I3"
    sheet["H3"] = "KEEPAMT"
    sheet.cell(3, 9, 100)
    sheet.cell(3, 10, 2)
    sheet.cell(3, 11, 999)
    workbook.save(material)
    workbook.close()
    staff = _fingerprint(material, "员工餐")
    _sync(monkeypatch, tmp_path, material, [
        _bill("july.jpg", JUL15, [_line(1, "香葱", "香葱", "4", "斤", "8", "蔬菜")]),
    ])
    workbook = _open(material)
    sheet = workbook["日物料采购"]
    assert sheet["F1"].value == "合计"
    assert sheet["H3"].value == "KEEPAMT"
    assert sheet["F3"].value == "=L3"
    assert sheet.cell(1, 9).value == excel_serial(JUL15)
    assert sheet.cell(3, 9).value == 4 and sheet.cell(3, 10).value is None and sheet.cell(3, 11).value == 8
    assert sheet.cell(1, 12).value == excel_serial(AUG1)
    assert sheet.cell(3, 12).value == 100 and sheet.cell(3, 13).value == 2 and sheet.cell(3, 14).value == 999
    merges = {(item.min_row, item.min_col, item.max_row, item.max_col) for item in sheet.merged_cells.ranges}
    assert (1, 9, 1, 11) in merges
    assert (1, 12, 1, 14) in merges
    workbook.close()
    assert _fingerprint(material, "员工餐") == staff


def test_same_sync_writes_an_existing_date_on_its_shifted_columns(tmp_path, monkeypatch):
    material = _august(tmp_path / "八月.xlsx", daily_dates=[(AUG1, 9), (SEP5, 12)])
    workbook = _open(material)
    workbook["日物料采购"].cell(3, 9, 100)
    workbook["日物料采购"].cell(3, 10, 2)
    workbook.save(material)
    workbook.close()
    _sync(monkeypatch, tmp_path, material, [
        _bill("mid.jpg", SEP2, [_line(1, "香葱", "香葱", "4", "斤", "6", "蔬菜")]),
        _bill("later.jpg", SEP5, [_line(1, "香葱", "香葱", "8", "斤", "9", "蔬菜")]),
    ])
    workbook = _open(material)
    sheet = workbook["日物料采购"]
    assert sheet.cell(1, 9).value == excel_serial(AUG1)
    assert sheet.cell(3, 9).value == 100 and sheet.cell(3, 10).value == 2
    assert sheet.cell(1, 12).value == excel_serial(SEP2)
    assert sheet.cell(3, 12).value == 4 and sheet.cell(3, 13).value is None and sheet.cell(3, 14).value == 6
    assert sheet.cell(1, 15).value == excel_serial(SEP5)
    assert sheet.cell(3, 15).value == 8 and sheet.cell(3, 16).value is None and sheet.cell(3, 17).value == 9
    workbook.close()


def test_two_missing_dates_insert_in_chronological_order(tmp_path, monkeypatch):
    material = _august(tmp_path / "八月.xlsx", daily_dates=[(AUG1, 9)])
    workbook = _open(material)
    workbook["日物料采购"].cell(3, 9, 100)
    workbook.save(material)
    workbook.close()
    _sync(monkeypatch, tmp_path, material, [
        _bill("later.jpg", SEP3, [_line(1, "香葱", "香葱", "2", "斤", "4", "蔬菜")]),
        _bill("sooner.jpg", SEP2, [_line(1, "香葱", "香葱", "3", "斤", "5", "蔬菜")]),
    ])
    workbook = _open(material)
    sheet = workbook["日物料采购"]
    assert sheet.cell(1, 9).value == excel_serial(AUG1)
    assert sheet.cell(1, 12).value == excel_serial(SEP2)
    assert sheet.cell(1, 15).value == excel_serial(SEP3)
    assert sheet.cell(3, 9).value == 100
    assert sheet.cell(3, 12).value == 3 and sheet.cell(3, 14).value == 5
    assert sheet.cell(3, 15).value == 2 and sheet.cell(3, 17).value == 4
    assert _count(sheet, 1, excel_serial(SEP2)) == 1
    assert _count(sheet, 1, excel_serial(SEP3)) == 1
    workbook.close()


def test_existing_date_column_is_not_duplicated(tmp_path, monkeypatch):
    material = _august(tmp_path / "八月.xlsx", daily_dates=[(AUG1, 9)])
    workbook = _open(material)
    sheet = workbook["日物料采购"]
    sheet.cell(1, 12, datetime(2026, 9, 2))
    sheet.merge_cells(start_row=1, start_column=12, end_row=1, end_column=14)
    for offset, label in enumerate(("数量", "单价", "金额")):
        sheet.cell(2, 12 + offset, label)
    sheet.cell(3, 9, 100)
    sheet.cell(3, 12, 999)
    sheet.cell(3, 14, 1)
    width = sheet.max_column
    workbook.save(material)
    workbook.close()
    _sync(monkeypatch, tmp_path, material, [
        _bill("a.jpg", SEP2, [
            _line(1, "香葱", "香葱", "2", "斤", "5", "蔬菜"),
            _line(2, "香葱", "香葱", "3", "斤", "7", "蔬菜"),
        ]),
    ])
    workbook = _open(material)
    sheet = workbook["日物料采购"]
    assert sheet.max_column == width
    assert sheet.cell(1, 12).value == datetime(2026, 9, 2)
    assert _count(sheet, 1, excel_serial(SEP2)) == 0
    assert sheet.cell(3, 12).value == 5 and sheet.cell(3, 13).value is None and sheet.cell(3, 14).value == 12
    assert sheet.cell(3, 9).value == 100
    headers = [sheet.cell(1, col).value for col in range(1, sheet.max_column + 1)]
    assert headers.count(datetime(2026, 9, 2)) == 1
    workbook.close()


def test_unit_mismatch_does_not_add_a_column(tmp_path, monkeypatch):
    material = _august(tmp_path / "八月.xlsx", daily_dates=[(AUG1, 9)])
    before = material.read_bytes()
    _workspace, result = _sync(monkeypatch, tmp_path, material, [
        _bill("bad.jpg", SEP2, [_line(1, "香葱", "香葱", "1", "公斤", "4", "蔬菜")]),
    ])
    assert material.read_bytes() == before
    assert any("不一致" in row["reason"] for row in result["skipped"])
    assert result["written"] == []


def test_broken_date_header_fails_and_rolls_back(tmp_path, monkeypatch):
    material = _august(tmp_path / "八月.xlsx", daily_dates=[(AUG1, 9)])
    workbook = _open(material)
    sheet = workbook["日物料采购"]
    sheet.cell(1, 12, excel_serial(SEP2))
    sheet.cell(2, 12, "数量")
    sheet.cell(2, 13, "单价")
    sheet.cell(2, 14, "小计")
    sheet.cell(3, 9, 100)
    workbook.save(material)
    workbook.close()
    before = material.read_bytes()
    _use_tmp(monkeypatch, tmp_path)
    workspace = Workspace()
    workspace.update_settings("", str(material))
    _plant([_bill("a.jpg", SEP2, [_line(1, "香葱", "香葱", "2", "斤", "5", "蔬菜")])])
    with pytest.raises(AppError) as caught:
        workspace.sync()
    assert "数量" in str(caught.value) and "金额" in str(caught.value)
    assert material.read_bytes() == before


def test_unreadable_header_fails_without_touching_the_file(tmp_path, monkeypatch):
    material = tmp_path / "坏表.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "日物料采购"
    sheet["A1"] = "没有品名"
    meat = workbook.create_sheet("肉类采购")
    meat["A1"] = "也没有"
    workbook.save(material)
    workbook.close()
    before = material.read_bytes()
    _use_tmp(monkeypatch, tmp_path)
    workspace = Workspace()
    workspace.update_settings("", str(material))
    _plant([_bill("a.jpg", SEP2, [_line(1, "香葱", "香葱", "2", "斤", "5", "蔬菜")])])
    with pytest.raises(AppError) as caught:
        workspace.sync()
    assert "表头" in str(caught.value)
    assert material.read_bytes() == before
    assert list(tmp_path.glob("*.bak_*")) == []


def test_locked_workbook_is_a_clear_error(tmp_path, monkeypatch):
    _use_tmp(monkeypatch, tmp_path)
    from tests.conftest import make_material
    material = make_material(tmp_path / "物料.xlsx")
    before = material.read_bytes()
    workspace = Workspace()
    workspace.update_settings("", str(material))
    _plant([_bill("a.jpg", AUG1, [_line(1, "小葱", "小葱", "2", "斤", "5", "蔬菜")])])

    def locked(*_args, **_kwargs):
        raise PermissionError("locked")

    monkeypatch.setattr("foodmix.sync_excel.load_workbook", locked)
    with pytest.raises(AppError) as caught:
        workspace.sync()
    assert "占用" in str(caught.value)
    assert material.read_bytes() == before


def test_tray_to_board_rule_was_not_added():
    raw = json.loads((Path(__file__).resolve().parents[1] / "config" / "unit_conversions.json").read_text(encoding="utf-8"))
    blob = json.dumps(raw, ensure_ascii=False)
    assert "托" not in blob
    assert [item["id"] for item in raw["rules"]] == ["egg-board-to-piece"]
