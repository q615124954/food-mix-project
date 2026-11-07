from pathlib import Path

from fastapi.testclient import TestClient
from PIL import Image

from foodmix.numbers import excel_serial
from foodmix.service import Workspace
from foodmix.settings import state_path
from foodmix.store import read_state, write_state
from tests.conftest import SAMPLE_DATE, make_material
from tests.test_sync import _line, _read, _use_tmp


def test_old_ocr_files_are_gone():
    root = Path(__file__).resolve().parents[1]
    assert not (root / "pipeline_min.py").exists()
    assert not (root / "meicai_vision.py").exists()
    source = "\n".join(path.read_text(encoding="utf-8") for path in (root / "foodmix").glob("*.py"))
    assert "HunyuanOCR" not in source
    assert "/images/ocr" not in source


def test_health_scan_confirm_and_sync(tmp_path, monkeypatch):
    _use_tmp(monkeypatch, tmp_path)
    monkeypatch.delenv("INPUT_DIR", raising=False)
    monkeypatch.delenv("MATERIAL_FILE", raising=False)
    material = make_material(tmp_path / "物料.xlsx")
    before = material.read_bytes()
    image = tmp_path / "美菜.png"
    Image.new("RGB", (12, 12), "white").save(image)
    (tmp_path / "备注.txt").write_text("不是图片", encoding="utf-8")

    from webapp.app import app, workspace
    workspace._state = None
    workspace._catalog = None
    client = TestClient(app)

    health = client.get("/api/health").json()
    assert health["phase"] == 1
    assert health["writes_staff_meal"] is False
    assert "api_key" not in health
    assert health["has_api_key"] is False
    page = client.get("/")
    assert page.status_code == 200
    assert "同步到物料明细" in page.text
    assert "确认（不写表）" in page.text
    assert "一期不写员工餐" in page.text

    client.post("/api/settings", json={"input_dir": str(tmp_path), "material_file": str(material)})
    scanned = client.post("/api/scan").json()
    assert [item["filename"] for item in scanned["invoices"]] == ["美菜.png"]
    assert client.get("/api/invoices/美菜.png/image").status_code == 200

    refused = client.post("/api/invoices/美菜.png/recognize")
    assert refused.status_code == 400
    assert "密钥" in refused.json()["detail"]
    assert material.read_bytes() == before

    state = read_state(state_path())
    state["invoices"][0]["status"] = "recognized"
    state["invoices"][0]["date"] = SAMPLE_DATE
    state["invoices"][0]["lines"] = [_line(1, "无敌美味小小葱", "无敌美味小小葱", "2", "斤", "5", "蔬菜")]
    write_state(state_path(), state)
    workspace._state = None

    confirmed = client.post("/api/invoices/美菜.png/confirm", json={
        "date": SAMPLE_DATE,
        "lines": [{"line_no": 1, "name": "小葱", "qty": "2", "unit": "斤", "amount": "5", "category": "蔬菜"}],
    })
    assert confirmed.status_code == 200
    assert confirmed.json()["lines"][0]["name"] == "小葱"
    assert confirmed.json()["lines"][0]["map_status"] == "alias"
    assert material.read_bytes() == before
    assert list(tmp_path.glob("*.bak_*")) == []

    synced = client.post("/api/sync")
    assert synced.status_code == 200
    body = synced.json()
    assert body["staff_meal_written"] is False
    assert body["written"][0]["sheet"] == "日物料采购"
    assert body["written"][0]["qty"] == "2"
    workbook_values = _read(material)
    assert workbook_values["scallion_qty"] == 2
    assert workbook_values["staff_d"] == 12345

    workbook = __import__("openpyxl").load_workbook(material)
    workbook["日物料采购"].cell(3, 9).value = 50
    workbook.save(material)
    workbook.close()
    client.post("/api/sync")
    assert _read(material)["scallion_qty"] == 2
    assert excel_serial(SAMPLE_DATE) == 46235


def test_recognize_keeps_each_line_category(tmp_path, monkeypatch):
    _use_tmp(monkeypatch, tmp_path)
    material = make_material(tmp_path / "物料.xlsx")
    folder = tmp_path / "shots"
    folder.mkdir()
    image = folder / "单.jpg"
    image.write_bytes(b"fake")
    workspace = Workspace()
    workspace.update_settings(str(folder), str(material))
    workspace.scan()

    class FakeVision:
        def extract(self, image_bytes, mime, prompt):
            assert "电子" in prompt or "JSON" in prompt
            return """
            {"date":"2026-08-01","supplier":"美菜","category":"蔬菜","lines":[
              {"name":"小葱","qty":2,"unit":"斤","amount":5,"category":"蔬菜"},
              {"name":"五花肉","qty":1,"unit":"斤","amount":20,"category":"肉类"},
              {"name":"可乐","qty":1,"unit":"瓶","amount":3,"category":"蔬菜"}
            ]}
            """

    class FakeMapper:
        def suggest_batch(self, queries, prompt, config):
            assert queries and queries[0]["raw_name"] == "可乐"
            return {"3": {"key": "3", "name": "雪碧", "confidence": 0.99}}

    invoice = workspace.recognize("单.jpg", vision=FakeVision(), suggestor=FakeMapper())
    categories = [line["category"] for line in invoice["lines"]]
    assert categories == ["蔬菜", "肉类", "蔬菜"]
    by_name = {line["raw_name"]: line for line in invoice["lines"]}
    assert by_name["小葱"]["preview"]["sheet"] == "日物料采购"
    assert by_name["五花肉"]["preview"]["sheet"] == "肉类采购"
    assert by_name["可乐"]["preview"]["will_write"] is False
    assert "拒绝" in by_name["可乐"]["map_reason"] or "找不到" in by_name["可乐"]["map_reason"]
