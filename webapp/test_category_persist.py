# -*- coding: utf-8 -*-
"""确认类别要落在单据上，重新打开和同步都用这份类别；缺类别时不得默认成蔬菜。"""
import os
import sys

import pytest
from fastapi import HTTPException
from openpyxl import Workbook, load_workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import app


DATE = "2026-08-27"


@pytest.fixture
def isolated(tmp_path, monkeypatch):
    out = tmp_path / "out"
    inn = tmp_path / "in"
    out.mkdir()
    inn.mkdir()
    monkeypatch.setattr(app, "STATE_FILE", str(tmp_path / "state.json"))
    monkeypatch.setattr(app, "_state", {"invoices": []})
    monkeypatch.setattr(app, "_paths", {
        "input_dir": str(inn),
        "output_dir": str(out),
        "material_file": str(tmp_path / "物料明细.xlsx"),
    })
    monkeypatch.setattr(app, "_code_map", {"肉类": {}, "蔬菜": {}})
    monkeypatch.setattr(app, "_name_dict", {"蔬菜": {}, "肉类": {}})
    monkeypatch.setattr(app, "_name_map", {"肉类": {}, "蔬菜": {}})
    return tmp_path


def _pending(filename="meat.jpg"):
    return {
        "filename": filename,
        "status": "pending",
        "ocr_done": True,
        "date": "",
        "items": [{
            "品名": "牛腩", "原品名": "牛腩", "数量": 2, "单价": 30, "金额": 60,
            "类别": "蔬菜",  # OCR 误判；确认时应以用户选择为准
        }],
        "suspects": ["品名存疑"],
    }


def _material(path, serial):
    wb = Workbook()
    meat = wb.active
    meat.title = "肉类采购"
    _sheet(meat, serial, [("M001", "牛腩"), ("M002", "牛骨")])
    veg = wb.create_sheet("日物料采购")
    _sheet(veg, serial, [("V001", "白菜")])
    wb.save(path)


def _sheet(ws, serial, products):
    ws.cell(1, 2, "代码")
    ws.cell(1, 4, "品名")
    ws.cell(1, 5, "单位")
    ws.cell(1, 6, serial)
    for i, (code, name) in enumerate(products):
        row = 3 + i
        ws.cell(row, 2, code)
        ws.cell(row, 4, name)
        ws.cell(row, 5, "斤")


def _land(path, data_rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "已落地"
    ws.append(["日期", "类别", "品名", "数量", "单价", "金额", "来源图片", "合计", "实付合计"])
    for row in data_rows:
        ws.append(list(row))
    wb.save(path)


def _fake_map(calls):
    codes = {
        ("肉类", "牛腩"): "M001",
        ("肉类", "牛骨"): "M002",
        ("蔬菜", "牛腩"): "V001",
        ("蔬菜", "牛骨"): "V001",
        ("蔬菜", "白菜"): "V001",
    }

    def fake(raw_name, category, name_code, model, qty=None, amt=None):
        calls.append((raw_name, category))
        code = codes.get((category, raw_name), "")
        if code not in set(name_code.values()):
            return {"物料代码": "", "目标品名": raw_name, "置信度": 0, "理由": "不在本表", "来源": "error"}
        return {"物料代码": code, "目标品名": raw_name, "置信度": 1.0, "理由": "test", "来源": "mapping"}

    return fake


def _qty(ws, name):
    serial_col = 6
    for r in range(3, ws.max_row + 1):
        if ws.cell(r, 4).value == name:
            return ws.cell(r, serial_col).value
    return None


def test_normalize_item_keeps_chosen_category_and_does_not_invent_vegetable():
    kept = app.normalize_item({"品名": "牛腩", "数量": "2", "金额": "60"}, category="肉类")
    assert kept["类别"] == "肉类"
    assert kept["数量"] == 2

    from_item = app.normalize_item({"品名": "牛腩", "类别": "肉类"})
    assert from_item["类别"] == "肉类"

    missing = app.normalize_item({"品名": "牛腩", "数量": 1, "金额": 1})
    assert "类别" not in missing
    assert missing.get("类别") != "蔬菜"

    # 手写单：整单类别覆盖行上的类别
    overwritten = app.normalize_item({"品名": "牛腩", "类别": "蔬菜"}, category="肉类")
    assert overwritten["类别"] == "肉类"
    # 美菜：行上已有类别不覆盖；空行才用默认
    kept_line = app.normalize_item(
        {"品名": "牛腩", "类别": "肉类"}, category="蔬菜", only_if_blank=True)
    assert kept_line["类别"] == "肉类"
    filled = app.normalize_item({"品名": "白菜", "数量": 1, "金额": 2}, category="蔬菜", only_if_blank=True)
    assert filled["类别"] == "蔬菜"


def test_confirm_persists_category_for_reopen(isolated):
    app._state["invoices"] = [_pending()]
    payload = app.ConfirmPayload(
        date=DATE, category=" 肉类 ", paid_total=60,
        items=[{"品名": "牛腩", "原品名": "牛腩", "数量": 2, "单价": 30, "金额": 60}],
    )
    assert app.confirm("meat.jpg", payload) == {"ok": True}

    inv = app._state["invoices"][0]
    assert inv["status"] == "confirmed"
    assert inv["category"] == "肉类"
    assert inv["items"][0]["类别"] == "肉类"

    app.load_state()
    reopened = app.get_state()["invoices"][0]
    assert reopened["category"] == "肉类"
    assert reopened["items"][0]["类别"] == "肉类"
    # 重新打开时前端优先读保存的类别，不应再落到蔬菜
    shown = (reopened.get("category") or reopened["items"][0].get("类别") or "")
    assert shown == "肉类"

    wb = load_workbook(os.path.join(app._paths["output_dir"], app.LAND_FILENAME))
    ws = wb.active
    assert ws.cell(2, 2).value == "肉类"


def test_confirm_rejects_missing_category_without_defaulting(isolated):
    app._state["invoices"] = [_pending()]
    payload = app.ConfirmPayload(
        date=DATE, category="",
        items=[{"品名": "牛腩", "数量": 2, "单价": 30, "金额": 60}],
    )
    with pytest.raises(HTTPException) as ei:
        app.confirm("meat.jpg", payload)
    assert ei.value.status_code == 400
    assert "不会默认写成蔬菜" in ei.value.detail

    inv = app._state["invoices"][0]
    assert inv["status"] == "pending"
    assert not inv.get("category")
    assert not os.path.exists(os.path.join(app._paths["output_dir"], app.LAND_FILENAME))


def test_old_invoice_without_category_is_not_backfilled(isolated):
    app._state["invoices"] = [{
        "filename": "old.jpg",
        "status": "confirmed",
        "items": [{"品名": "牛腩", "数量": 1, "金额": 10}],
    }]
    inv = app.get_state()["invoices"][0]
    assert not str(inv.get("category") or "").strip()
    assert not str(inv["items"][0].get("类别") or "").strip()
    assert "old.jpg" not in app.confirmed_categories()


def test_sync_uses_saved_category_when_land_file_says_vegetable(isolated, monkeypatch):
    """落地表被误写成蔬菜时，同步仍按确认时保存的肉类入账。"""
    calls = []
    monkeypatch.setattr(app, "map_product_code", _fake_map(calls))
    serial = app._excel_serial(DATE)
    _material(app._paths["material_file"], serial)
    _land(app.land_file(), [
        (DATE, "蔬菜", "牛腩", 2, 30, 60, "meat.jpg", 60, 60),
        (DATE, None, "牛骨", 1, 20, 20, "meat.jpg", None, None),
    ])
    app._state["invoices"] = [{
        "filename": "meat.jpg",
        "status": "confirmed",
        "category": "肉类",
        "items": [
            {"品名": "牛腩", "类别": "肉类"},
            {"品名": "牛骨", "类别": "肉类"},
        ],
    }]

    results, _backup = app.sync_to_material("test-model")
    assert ("牛腩", "肉类") in calls
    assert ("牛骨", "肉类") in calls
    assert all(r.get("来源") != "error" for r in results)

    wb = load_workbook(app._paths["material_file"])
    assert _qty(wb["肉类采购"], "牛腩") == 2
    assert _qty(wb["肉类采购"], "牛骨") == 1
    assert _qty(wb["日物料采购"], "白菜") is None


def test_sync_confirm_roundtrip_keeps_meat_on_meat_sheet(isolated, monkeypatch):
    calls = []
    monkeypatch.setattr(app, "map_product_code", _fake_map(calls))
    serial = app._excel_serial(DATE)
    _material(app._paths["material_file"], serial)
    app._state["invoices"] = [_pending()]
    app._state["invoices"][0]["items"].append({
        "品名": "牛骨", "原品名": "牛骨", "数量": 1, "单价": 20, "金额": 20, "类别": "蔬菜",
    })
    app.confirm("meat.jpg", app.ConfirmPayload(
        date=DATE, category="肉类", paid_total=80,
        items=[
            {"品名": "牛腩", "原品名": "牛腩", "数量": 2, "单价": 30, "金额": 60},
            {"品名": "牛骨", "原品名": "牛骨", "数量": 1, "单价": 20, "金额": 20},
        ],
    ))
    app.load_state()
    results, _backup = app.sync_to_material("test-model")
    assert ("牛腩", "肉类") in calls and ("牛骨", "肉类") in calls
    assert not any("缺少类别" in r.get("理由", "") for r in results)
    wb = load_workbook(app._paths["material_file"])
    assert _qty(wb["肉类采购"], "牛腩") == 2
    assert _qty(wb["肉类采购"], "牛骨") == 1
    assert _qty(wb["日物料采购"], "白菜") is None


def test_sync_old_land_category_still_used_when_state_has_none(isolated, monkeypatch):
    calls = []
    monkeypatch.setattr(app, "map_product_code", _fake_map(calls))
    serial = app._excel_serial(DATE)
    _material(app._paths["material_file"], serial)
    _land(app.land_file(), [
        (DATE, "肉类", "牛腩", 4, 30, 120, "old-meat.jpg", 120, 120),
    ])
    app._state["invoices"] = [{
        "filename": "old-meat.jpg",
        "status": "confirmed",
        "items": [{"品名": "牛腩"}],
    }]
    app.sync_to_material("test-model")
    assert ("牛腩", "肉类") in calls
    wb = load_workbook(app._paths["material_file"])
    assert _qty(wb["肉类采购"], "牛腩") == 4
    assert _qty(wb["日物料采购"], "白菜") is None


def test_sync_old_multirow_keeps_category_from_first_land_row(isolated, monkeypatch):
    """旧落地表类别只写在首行（合并单元格）时，同单后续行仍用该类别，不改成蔬菜。"""
    calls = []
    monkeypatch.setattr(app, "map_product_code", _fake_map(calls))
    serial = app._excel_serial(DATE)
    _material(app._paths["material_file"], serial)
    _land(app.land_file(), [
        (DATE, "肉类", "牛腩", 2, 30, 60, "old-meat.jpg", 80, 80),
        (DATE, None, "牛骨", 1, 20, 20, "old-meat.jpg", None, None),
    ])
    app._state["invoices"] = [{
        "filename": "old-meat.jpg",
        "status": "confirmed",
        "items": [{"品名": "牛腩"}, {"品名": "牛骨"}],
    }]
    results, _backup = app.sync_to_material("test-model")
    assert ("牛腩", "肉类") in calls and ("牛骨", "肉类") in calls
    assert not any("缺少类别" in r.get("理由", "") for r in results)
    wb = load_workbook(app._paths["material_file"])
    assert _qty(wb["肉类采购"], "牛腩") == 2
    assert _qty(wb["肉类采购"], "牛骨") == 1


def test_sync_missing_category_does_not_inherit_vegetable(isolated, monkeypatch):
    calls = []
    monkeypatch.setattr(app, "map_product_code", _fake_map(calls))
    serial = app._excel_serial(DATE)
    _material(app._paths["material_file"], serial)
    _land(app.land_file(), [
        (DATE, "蔬菜", "白菜", 1, 2, 2, "veg.jpg", 2, 2),
        (DATE, None, "牛腩", 9, 30, 270, "meat.jpg", 270, 270),
    ])
    app._state["invoices"] = [
        {"filename": "veg.jpg", "status": "confirmed", "items": [{"品名": "白菜"}]},
        {"filename": "meat.jpg", "status": "confirmed", "items": [{"品名": "牛腩"}]},
    ]
    results, _backup = app.sync_to_material("test-model")
    reasons = [r for r in results if r.get("落地品名") == "牛腩"]
    assert reasons and "缺少类别" in reasons[0]["理由"]
    assert "未默认归入蔬菜" in reasons[0]["理由"]
    assert ("牛腩", "蔬菜") not in calls
    wb = load_workbook(app._paths["material_file"])
    assert _qty(wb["日物料采购"], "白菜") == 1
    assert _qty(wb["肉类采购"], "牛腩") is None


def test_frontend_does_not_default_category_to_vegetable():
    html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static", "index.html")
    html = open(html_path, encoding="utf-8").read()
    assert "|| '蔬菜'" not in html
    assert '|| "蔬菜"' not in html
    assert "请先选择类别（不会默认成蔬菜）" in html
    assert "不会默认成蔬菜" in html
    # 美菜按行选类别，按钮仍是「确认并同步」；手写单仍要求整单类别。
    assert "inv.kind === 'meicai' ? '确认并同步' : '确认落地'" in html
    assert 'data-k="类别"' in html
    assert "默认类别" in html
    assert "有明细未选类别。请选择默认类别，或为每一行选择蔬菜/肉类（不会默认成蔬菜）" in html


def _meicai(filename="mix.png"):
    return {
        "filename": filename,
        "status": "pending",
        "ocr_done": True,
        "kind": "meicai",
        "date": DATE,
        "items": [],
    }


def test_meicai_confirm_lands_each_line_category_and_does_not_restamp_meat(isolated, monkeypatch):
    """默认类别是蔬菜时，肉行仍落地为肉类，空行才用默认。同步按行进对应的表。"""
    calls = []
    monkeypatch.setattr(app, "map_product_code", _fake_map(calls))
    app._state["invoices"] = [_meicai()]
    result = app.confirm("mix.png", app.ConfirmPayload(
        date=DATE, category="蔬菜", paid_total=75,
        items=[
            {"品名": "白菜", "数量": 1, "单价": 10, "金额": 10, "类别": "蔬菜", "渠道": "美菜"},
            {"品名": "牛腩", "数量": 2, "单价": 30, "金额": 60, "类别": "肉类", "渠道": "美菜"},
            {"品名": "土豆", "数量": 1, "单价": 5, "金额": 5, "类别": "", "渠道": "美菜"},
        ],
    ))
    assert result["ok"] is True
    assert "material_sync" in result

    inv = app._state["invoices"][0]
    assert inv["status"] == "confirmed"
    assert inv["category"] == "蔬菜"
    assert [it["类别"] for it in inv["items"]] == ["蔬菜", "肉类", "蔬菜"]
    assert all(it["渠道"] == "美菜" for it in inv["items"])

    wb = load_workbook(os.path.join(app._paths["output_dir"], app.LAND_FILENAME))
    ws = wb.active
    assert [ws.cell(r, 2).value for r in range(2, 5)] == ["蔬菜", "肉类", "蔬菜"]
    assert [ws.cell(r, 3).value for r in range(2, 5)] == ["白菜", "牛腩", "土豆"]
    cat_merges = [mr for mr in ws.merged_cells.ranges if mr.min_col == 2 and mr.max_row > mr.min_row]
    assert cat_merges == []
    assert any(mr.min_col == 1 and mr.max_row > mr.min_row for mr in ws.merged_cells.ranges)
    wb.close()

    serial = app._excel_serial(DATE)
    _material(app._paths["material_file"], serial)
    app.sync_to_material("test-model")
    assert ("牛腩", "肉类") in calls
    assert ("牛腩", "蔬菜") not in calls
    assert ("白菜", "蔬菜") in calls
    booked = load_workbook(app._paths["material_file"])
    assert _qty(booked["肉类采购"], "牛腩") == 2
    assert _qty(booked["日物料采购"], "白菜") == 1
    assert _qty(booked["日物料采购"], "牛腩") is None


def test_meicai_confirm_requires_default_only_when_a_line_lacks_category(isolated):
    app._state["invoices"] = [_meicai()]
    with pytest.raises(HTTPException) as ei:
        app.confirm("mix.png", app.ConfirmPayload(
            date=DATE, category="",
            items=[
                {"品名": "牛腩", "数量": 1, "单价": 30, "金额": 30, "类别": "肉类"},
                {"品名": "白菜", "数量": 1, "单价": 2, "金额": 2},
            ],
        ))
    assert ei.value.status_code == 400
    assert "不会默认写成蔬菜" in ei.value.detail
    assert app._state["invoices"][0]["status"] == "pending"
    assert not os.path.exists(os.path.join(app._paths["output_dir"], app.LAND_FILENAME))

    app.confirm("mix.png", app.ConfirmPayload(
        date=DATE, category="", paid_total=32,
        items=[
            {"品名": "牛腩", "数量": 1, "单价": 30, "金额": 30, "类别": "肉类"},
            {"品名": "白菜", "数量": 1, "单价": 2, "金额": 2, "类别": "蔬菜"},
        ],
    ))
    inv = app._state["invoices"][0]
    assert inv["status"] == "confirmed"
    assert inv["category"] == ""
    assert [it["类别"] for it in inv["items"]] == ["肉类", "蔬菜"]
    wb = load_workbook(os.path.join(app._paths["output_dir"], app.LAND_FILENAME))
    ws = wb.active
    assert [ws.cell(r, 2).value for r in range(2, 4)] == ["肉类", "蔬菜"]
    wb.close()


def test_handwriting_confirm_still_stamps_whole_invoice_category(isolated):
    inv = _pending("hand.jpg")
    inv["kind"] = "handwriting"
    app._state["invoices"] = [inv]
    result = app.confirm("hand.jpg", app.ConfirmPayload(
        date=DATE, category="肉类", paid_total=80,
        items=[
            {"品名": "牛腩", "原品名": "牛腩", "数量": 2, "单价": 30, "金额": 60, "类别": "蔬菜"},
            {"品名": "牛骨", "原品名": "牛骨", "数量": 1, "单价": 20, "金额": 20, "类别": "蔬菜"},
        ],
    ))
    assert result == {"ok": True}
    saved = app._state["invoices"][0]
    assert saved["category"] == "肉类"
    assert [it["类别"] for it in saved["items"]] == ["肉类", "肉类"]
    wb = load_workbook(os.path.join(app._paths["output_dir"], app.LAND_FILENAME))
    ws = wb.active
    assert ws.cell(2, 2).value == "肉类"
    assert ws.cell(3, 2).value in (None, "肉类")
    assert any(mr.min_col == 2 and mr.max_row > mr.min_row for mr in ws.merged_cells.ranges)
    wb.close()


def test_hint_line_category_only_when_catalog_is_unique(monkeypatch):
    monkeypatch.setattr(app, "_legal_names", {app.normalize_pin("白菜")})
    monkeypatch.setattr(app, "_meat_names", {app.normalize_pin("牛腩")})
    monkeypatch.setattr(app, "_name_map", {"蔬菜": {"青菜": "白菜"}, "肉类": {"牛腩块": "牛腩"}})
    monkeypatch.setattr(app, "_code_map", {"蔬菜": {}, "肉类": {"吊龙": "M1"}})
    assert app.hint_line_category("白菜") == "蔬菜"
    assert app.hint_line_category("牛腩") == "肉类"
    assert app.hint_line_category("青菜") == "蔬菜"
    assert app.hint_line_category("吊龙") == "肉类"
    assert app.hint_line_category("鱼腥草") == ""
    assert app.hint_line_category("") == ""
    monkeypatch.setattr(app, "_legal_names", {app.normalize_pin("牛腩")})
    assert app.hint_line_category("牛腩") == ""


def test_pack_meicai_prefills_line_category_without_guessing(monkeypatch):
    monkeypatch.setattr(app, "_legal_names", {app.normalize_pin("白菜")})
    monkeypatch.setattr(app, "_legal_name_list", [])
    monkeypatch.setattr(app, "_meat_names", {app.normalize_pin("牛腩")})
    monkeypatch.setattr(app, "_name_map", {"蔬菜": {}, "肉类": {}})
    monkeypatch.setattr(app, "_code_map", {"蔬菜": {}, "肉类": {}})
    monkeypatch.setattr(app, "_name_dict", {"蔬菜": {}, "肉类": {}})
    packed = app._pack_meicai(
        {"order_total": 11, "order_ids": [], "note": "", "meicai_suspects": []},
        [
            {"品名": "白菜", "数量": 1, "金额": 4, "类别": ""},
            {"品名": "牛腩", "数量": 1, "金额": 6, "类别": "蔬菜"},
            {"品名": "不知道", "数量": 1, "金额": 1},
        ],
        DATE,
    )
    assert [it["类别"] for it in packed["items"]] == ["蔬菜", "蔬菜", ""]


def test_load_legal_names_reads_meat_sheet_for_hints(isolated):
    saved = (
        set(app._legal_names), list(app._legal_name_list), dict(app._legal_units),
        dict(app._legal_codes), dict(app._hist_prices),
        set(app._meat_names), list(app._meat_name_list), app._catalog_error,
    )
    try:
        _material(app._paths["material_file"], app._excel_serial(DATE))
        app.load_legal_names()
        assert app.normalize_pin("白菜") in app._legal_names
        assert app.normalize_pin("牛腩") in app._meat_names
        assert app.normalize_pin("牛腩") not in app._legal_names
        assert app.normalize_pin("白菜") not in app._meat_names
        assert not app._catalog_error
    finally:
        app._legal_names, app._legal_name_list = saved[0], saved[1]
        app._legal_units, app._legal_codes, app._hist_prices = saved[2], saved[3], saved[4]
        app._meat_names, app._meat_name_list = saved[5], saved[6]
        app._catalog_error = saved[7]
