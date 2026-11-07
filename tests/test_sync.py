from decimal import Decimal

from openpyxl import load_workbook

from foodmix.numbers import excel_serial
from foodmix.plan import BillLine, Write, build_plan
from foodmix.service import Workspace
from foodmix.settings import load_sheet_layout, state_path
from foodmix.store import empty_state, read_state, write_state
from foodmix.sync_excel import apply_plan
from foodmix.units import load_rules
from tests.conftest import SAMPLE_DATE, make_material


def _use_tmp(monkeypatch, tmp_path):
    monkeypatch.setenv("FOODMIX_STATE_PATH", str(tmp_path / "state.json"))
    monkeypatch.setenv("FOODMIX_SETTINGS_PATH", str(tmp_path / "settings.json"))
    monkeypatch.setenv("FOODMIX_ALIASES_PATH", str(tmp_path / "aliases.json"))
    monkeypatch.delenv("HUNYUAN_API_KEY", raising=False)
    monkeypatch.delenv("MOLLI_ARK_API_KEY", raising=False)


def _line(number, raw, name, qty, unit, amount, category):
    return {
        "line_no": number,
        "raw_name": raw,
        "name": name,
        "qty": qty,
        "unit": unit,
        "amount": amount,
        "category": category,
        "map_status": "exact",
        "map_reason": "",
        "suggestion": "",
        "confidence": None,
        "issues": [],
    }


def _invoice(name, lines, status="confirmed"):
    return {
        "id": name,
        "filename": name,
        "path": name,
        "status": status,
        "date": SAMPLE_DATE,
        "supplier": "美菜",
        "error": "",
        "warning": "",
        "confirmed_at": "",
        "lines": lines,
    }


def _plant(invoices, ledger=None):
    state = empty_state()
    state["invoices"] = invoices
    if ledger:
        state["ledger"] = ledger
    write_state(state_path(), state)


def _read(path):
    workbook = load_workbook(path, data_only=False)
    try:
        daily = workbook["日物料采购"]
        meat = workbook["肉类采购"]
        staff = workbook["员工餐"]
        return {
            "scallion_qty": daily.cell(3, 9).value,
            "scallion_amt": daily.cell(3, 11).value,
            "keep": daily["F3"].value,
            "keep_amt": daily["H3"].value,
            "bag": daily.cell(5, 9).value,
            "egg_qty": daily.cell(4, 9).value,
            "egg_amt": daily.cell(4, 11).value,
            "meat_qty": meat.cell(4, 9).value,
            "meat_amt": meat.cell(4, 11).value,
            "staff_d": staff["D3"].value,
            "staff_h": staff["H3"].value,
        }
    finally:
        workbook.close()


def test_excel_serial_matches_the_accounting_sheet():
    assert excel_serial("2026-08-01") == 46235


def test_sync_rewrites_the_confirmed_sum_and_skips_unknown_names(tmp_path, monkeypatch):
    _use_tmp(monkeypatch, tmp_path)
    material = make_material(tmp_path / "物料.xlsx")
    workspace = Workspace()
    workspace.update_settings("", str(material))
    _plant([
        _invoice("a.jpg", [
            _line(1, "小葱", "小葱", "2", "斤", "5", "蔬菜"),
            _line(2, "五花肉", "五花肉", "1", "斤", "20", "肉类"),
            _line(3, "可乐", "可乐", "1", "瓶", "3", "蔬菜"),
        ]),
        _invoice("b.jpg", [
            _line(1, "小葱", "小葱", "3", "斤", "7", "蔬菜"),
        ]),
        _invoice("c.jpg", [
            _line(1, "小葱", "小葱", "100", "斤", "100", "蔬菜"),
        ], status="recognized"),
    ], ledger={"material_file": str(material), "cells": [{"sheet": "员工餐", "row": 3, "col": 4}]})
    first = workspace.sync()
    values = _read(material)
    assert values["scallion_qty"] == 5
    assert values["scallion_amt"] == 12
    assert values["meat_qty"] == 1
    assert values["meat_amt"] == 20
    assert values["keep"] == "KEEP"
    assert values["keep_amt"] == "KEEPAMT"
    assert values["bag"] == 77
    assert values["staff_d"] == 12345
    assert values["staff_h"] == 888
    assert any(row["raw_name"] == "可乐" and "找不到" in row["reason"] for row in first["skipped"])
    assert first["staff_meal_written"] is False
    assert not any(row["sheet"] == "员工餐" for row in first["written"])

    workbook = load_workbook(material)
    workbook["日物料采购"].cell(3, 9).value = 100
    workbook.save(material)
    workbook.close()
    workspace.sync()
    assert _read(material)["scallion_qty"] == 5

    state = read_state(state_path())
    for invoice in state["invoices"]:
        if invoice["id"] == "b.jpg":
            invoice["status"] = "recognized"
        if invoice["id"] == "a.jpg":
            invoice["lines"] = [line for line in invoice["lines"] if line["category"] != "肉类"]
    write_state(state_path(), state)
    workspace._state = None
    workspace.sync()
    values = _read(material)
    assert values["scallion_qty"] == 2
    assert values["scallion_amt"] == 5
    assert values["meat_qty"] is None
    assert values["meat_amt"] is None
    assert values["staff_d"] == 12345


def test_unit_mismatch_does_not_touch_the_cell_and_eggs_convert(tmp_path, monkeypatch):
    _use_tmp(monkeypatch, tmp_path)
    material = make_material(tmp_path / "物料.xlsx", egg_unit="个")
    before = _read(material)
    workspace = Workspace()
    workspace.update_settings("", str(material))
    _plant([_invoice("bad.jpg", [_line(1, "小葱", "小葱", "1", "把", "4", "蔬菜")])])
    result = workspace.sync()
    assert _read(material)["scallion_qty"] == 100
    assert any("不一致" in row["reason"] for row in result["skipped"])

    _plant([
        _invoice("egg.jpg", [
            _line(1, "鸡蛋", "鸡蛋", "2", "板", "15", "蔬菜"),
            _line(2, "煲汤袋", "煲汤袋", "1", "板", "4", "蔬菜"),
        ])
    ])
    workspace._state = None
    converted = workspace.sync()
    values = _read(material)
    assert values["egg_qty"] == 60
    assert values["egg_amt"] == 15
    assert values["bag"] == 77
    assert any(row["name"] == "煲汤袋" and "不一致" in row["reason"] for row in converted["skipped"])
    assert before["staff_d"] == values["staff_d"]


def test_confirm_does_not_write_the_material_file(tmp_path, monkeypatch):
    _use_tmp(monkeypatch, tmp_path)
    material = make_material(tmp_path / "物料.xlsx")
    before = material.read_bytes()
    workspace = Workspace()
    workspace.update_settings("", str(material))
    _plant([_invoice("a.jpg", [_line(1, "小葱", "小葱", "2", "斤", "5", "蔬菜")], status="recognized")])
    workspace.confirm("a.jpg", SAMPLE_DATE, [{
        "line_no": 1, "name": "小葱", "qty": "2", "unit": "斤", "amount": "5", "category": "蔬菜",
    }])
    assert material.read_bytes() == before
    assert list(tmp_path.glob("*.bak_*")) == []
    assert workspace.get_invoice("a.jpg")["status"] == "confirmed"


def test_apply_refuses_staff_sheet(tmp_path):
    material = make_material(tmp_path / "物料.xlsx")
    write = Write("员工餐", 3, 4, 5, "员工餐", "", "", SAMPLE_DATE, Decimal("1"), Decimal("1"), ["a"])
    try:
        apply_plan(material, [write], [], load_sheet_layout())
        raised = False
    except Exception as exc:
        raised = True
        assert "员工餐" in str(exc)
    assert raised
    assert _read(material)["staff_d"] == 12345


def test_rules_object_is_the_unit_table():
    assert load_rules()
    assert isinstance(load_rules()[0].factor, Decimal)
