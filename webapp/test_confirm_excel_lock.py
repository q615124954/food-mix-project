# -*- coding: utf-8 -*-
"""落地表被占用写失败时，单据必须保持未确认，不能先把状态写成已确认。"""
import json
import os
import sys

import pytest
from fastapi import HTTPException
from openpyxl.workbook.workbook import Workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import app


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
        "material_file": "",
    })
    monkeypatch.setattr(app, "_code_map", {"肉类": {}, "蔬菜": {}})
    monkeypatch.setattr(app, "_name_dict", {"蔬菜": {}, "肉类": {}})
    monkeypatch.setattr(app, "_name_map", {"肉类": {}, "蔬菜": {}})
    return tmp_path


def _pending(filename="order.jpg"):
    return {
        "filename": filename,
        "status": "pending",
        "ocr_done": True,
        "date": "2026-10-01",
        "category": "蔬菜",
        "total": 10,
        "paid_total": 10,
        "items": [{
            "品名": "白菜", "原品名": "白菜", "数量": 1, "单价": 10, "金额": 10, "类别": "蔬菜",
        }],
        "suspects": ["品名存疑"],
    }


def test_confirm_stays_pending_when_excel_write_fails(isolated, monkeypatch):
    filename = "order.jpg"
    app._state["invoices"] = [_pending(filename)]
    app.save_state()

    def locked_save(self, path):
        # 写盘发生时状态必须仍是未确认，不能先落成已确认再写 Excel
        assert app._state["invoices"][0]["status"] == "pending"
        raise PermissionError("[Errno 13] Permission denied")

    monkeypatch.setattr(Workbook, "save", locked_save)

    payload = app.ConfirmPayload(
        date="2026-10-13",
        category="蔬菜",
        items=[{"品名": "白菜", "原品名": "白菜", "数量": 1, "单价": 10, "金额": 10}],
        paid_total=10,
    )
    with pytest.raises(HTTPException) as ei:
        app.confirm(filename, payload)

    assert ei.value.status_code == 500
    assert "占用" in ei.value.detail
    assert "关闭" in ei.value.detail

    inv = app._state["invoices"][0]
    assert inv["status"] == "pending"

    saved = json.loads(open(app.STATE_FILE, encoding="utf-8").read())
    assert saved["invoices"][0]["status"] == "pending"

    state = app.get_state()
    assert state["summary"]["confirmed"] == 0
    assert state["invoices"][0]["status"] == "pending"
    assert not os.path.exists(app.land_file())
