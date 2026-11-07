# -*- coding: utf-8 -*-
"""美菜视觉 JSON → 复核页形状。不调用模力方舟。

Shot3 煲汤袋（常见 3/14.2，金标 1/4.8）不在这里修，也不当作通过条件。
"""
import copy
import json
import os
import sys
import tempfile
import unittest
from unittest.mock import patch

from PIL import Image

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(ROOT, "webapp"))
sys.path.insert(0, ROOT)

import meicai_vision
import pipeline_min as core
import app as webapp


def _review_payload():
    return {
        "date": "2026-09-03",
        "paid": 91.81,
        "order_ids": ["2997320402"],
        "items": [
            {"品名": "小葱", "数量": 3, "金额": 26.91, "单位": "斤", "渠道": "美菜", "订单号": "2997320402"},
            {"品名": "小葱", "数量": 5, "金额": 44.40, "单位": "斤", "渠道": "美菜"},
            {"品名": "香菜", "数量": 1, "金额": 8.5, "单位": "斤", "渠道": "美菜"},
            {"品名": "鱼腥草", "数量": 2, "金额": 12, "单位": "斤", "渠道": "美菜"},
        ],
    }


class ModelChoiceTests(unittest.TestCase):
    def test_default_is_qwen35_flash(self):
        with patch.dict(os.environ, {"MEICAI_VISION_MODEL": ""}, clear=False):
            self.assertEqual(meicai_vision.selected_model(), "Qwen3.5-Flash")
        self.assertEqual(meicai_vision.FLASH_MODEL, "Qwen3.5-Flash")

    def test_vl_plus_is_opt_in(self):
        with patch.dict(os.environ, {"MEICAI_VISION_MODEL": "Qwen3-VL-Plus"}, clear=False):
            self.assertEqual(meicai_vision.selected_model(), "Qwen3-VL-Plus")
        with patch.dict(os.environ, {"MEICAI_VISION_MODEL": "vl-plus"}, clear=False):
            self.assertEqual(meicai_vision.selected_model(), "Qwen3-VL-Plus")

    def test_unknown_name_stays_on_flash(self):
        for raw in ("Qwen3.6-Plus", "Qwen3-VL-Flash", "flash", "Qwen3.5-Flash"):
            with patch.dict(os.environ, {"MEICAI_VISION_MODEL": raw}, clear=False):
                self.assertEqual(meicai_vision.selected_model(), "Qwen3.5-Flash", raw)


class ReviewShapeTests(unittest.TestCase):
    def test_json_becomes_review_rows(self):
        text = "```json\n" + json.dumps(_review_payload(), ensure_ascii=False) + "\n```"
        doc = meicai_vision.review_doc_from_model_text(text)
        by_name = {it["品名"]: it for it in doc["items"]}
        self.assertEqual(set(by_name), {"小葱", "香菜", "鱼腥草"})
        scallion = by_name["小葱"]
        self.assertAlmostEqual(scallion["数量"], 8.0, places=2)
        self.assertAlmostEqual(scallion["金额"], 71.31, places=2)
        self.assertEqual(scallion["单位"], "斤")
        self.assertEqual(scallion["渠道"], "美菜")
        self.assertEqual(scallion["订单号"], "2997320402")
        self.assertFalse(scallion["实收缺失"])
        self.assertEqual(scallion["勾稽"], "OK")
        self.assertAlmostEqual(scallion["单价"], 8.91, places=2)
        self.assertEqual(scallion["日期"], "2026-09-03")
        self.assertEqual(by_name["香菜"]["单位"], "斤")
        self.assertAlmostEqual(by_name["鱼腥草"]["金额"], 12.0, places=2)
        self.assertAlmostEqual(doc["paid"], 91.81, places=2)
        self.assertEqual(doc["order_ids"], ["2997320402"])
        self.assertEqual(doc["kind"], "美菜")

    def test_missing_receive_stays_null(self):
        # 不把缺实收的行填成订货数。煲汤袋数量对错不是这条测试的目标。
        text = json.dumps({
            "date": "2026-09-04",
            "paid": 4.8,
            "order_ids": ["1"],
            "items": [{
                "品名": "煲汤袋", "数量": None, "金额": None, "单位": "份",
                "渠道": "美菜", "实收缺失": True,
            }],
        })
        doc = meicai_vision.review_doc_from_model_text(text)
        row = doc["items"][0]
        self.assertEqual(row["品名"], "煲汤袋")
        self.assertIsNone(row["数量"])
        self.assertIsNone(row["金额"])
        self.assertTrue(row["实收缺失"])
        self.assertIn("实收", " ".join(doc["suspects"]))

    def test_boxes_are_rejected(self):
        text = json.dumps([{"bbox": [0, 0, 1, 1], "text": "小葱"}])
        with self.assertRaises(ValueError):
            meicai_vision.review_doc_from_model_text(text)


class RecognizeOfflineTests(unittest.TestCase):
    def test_no_key_does_not_post(self):
        def boom(*_args, **_kwargs):
            raise AssertionError("should not call chat completions")

        env = {"HUNYUAN_API_KEY": "", "MOLLI_ARK_API_KEY": "", "MEICAI_VISION_MODEL": ""}
        with patch.dict(os.environ, env, clear=False):
            with self.assertRaises(RuntimeError) as ctx:
                meicai_vision.recognize("missing.jpg", caller=boom)
        self.assertIn("HUNYUAN_API_KEY", str(ctx.exception))

    def test_caller_gets_flash_by_default_and_plus_when_configured(self):
        seen = {}

        def fake(key, model, path, timeout):
            seen["key"] = key
            seen["model"] = model
            seen["path"] = path
            return {"text": json.dumps(_review_payload(), ensure_ascii=False), "error": ""}

        env = {"HUNYUAN_API_KEY": "test-key", "MOLLI_ARK_API_KEY": "", "MEICAI_VISION_MODEL": ""}
        with patch.dict(os.environ, env, clear=False):
            doc = meicai_vision.recognize("shot.jpg", caller=fake)
        self.assertEqual(seen["key"], "test-key")
        self.assertEqual(seen["model"], "Qwen3.5-Flash")
        self.assertEqual(seen["path"], "shot.jpg")
        self.assertEqual(doc["model"], "Qwen3.5-Flash")
        self.assertEqual(doc["items"][0]["品名"], "小葱")

        env["MEICAI_VISION_MODEL"] = "Qwen3-VL-Plus"
        with patch.dict(os.environ, env, clear=False):
            doc = meicai_vision.recognize("shot.jpg", caller=fake)
        self.assertEqual(seen["model"], "Qwen3-VL-Plus")
        self.assertEqual(doc["model"], "Qwen3-VL-Plus")

    def test_http_error_is_raised(self):
        def fake(_key, _model, _path, _timeout):
            return {"text": "", "error": "HTTP 401: unauthorized"}

        env = {"HUNYUAN_API_KEY": "test-key", "MEICAI_VISION_MODEL": ""}
        with patch.dict(os.environ, env, clear=False):
            with self.assertRaises(RuntimeError) as ctx:
                meicai_vision.recognize("shot.jpg", caller=fake)
        self.assertIn("401", str(ctx.exception))


class RouteTests(unittest.TestCase):
    def setUp(self):
        self._paths = dict(webapp._paths)
        self._state = copy.deepcopy(webapp._state)

    def tearDown(self):
        webapp._paths.clear()
        webapp._paths.update(self._paths)
        webapp._state.clear()
        webapp._state.update(self._state)

    def test_long_header_meicai_skips_tiled_hunyuan(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        Image.new("RGB", (1080, 3200), "white").save(os.path.join(tmp.name, "long.png"))
        webapp._paths["input_dir"] = tmp.name
        header = "\n".join([
            "美菜",
            "送达时间 2026-09-03 06:30:00",
            "订单号：2997320402 复制",
            "多退少补",
        ])
        self.assertTrue(core.is_meicai_ocr(header))

        def strip(_path, timeout=600):
            return header

        def tiled(*_args, **_kwargs):
            raise AssertionError("顶条已是美菜，不应再分条 OCR")

        def full(*_args, **_kwargs):
            raise AssertionError("顶条已是美菜，不应再整图 OCR")

        def vision(path):
            self.assertTrue(path.endswith("long.png"))
            return {
                "date": "2026-09-03",
                "paid": 12.0,
                "order_ids": ["2997320402"],
                "note": "订单号 2997320402",
                "items": [{
                    "品名": "鱼腥草", "数量": 2, "金额": 12, "单位": "斤",
                    "渠道": "美菜", "实收缺失": False, "勾稽": "OK", "单价": 6,
                }],
                "suspects": [],
                "model": "Qwen3.5-Flash",
            }

        old_strip = core.call_ocr_top_strip
        old_tiled = core.call_ocr_tiled
        old_full = core.call_ocr
        old_vision = webapp.recognize_meicai_image
        core.call_ocr_top_strip = strip
        core.call_ocr_tiled = tiled
        core.call_ocr = full
        webapp.recognize_meicai_image = vision
        try:
            res = webapp.ocr_one("long.png")
        finally:
            core.call_ocr_top_strip = old_strip
            core.call_ocr_tiled = old_tiled
            core.call_ocr = old_full
            webapp.recognize_meicai_image = old_vision
        self.assertEqual(res["kind"], "meicai")
        self.assertEqual(res["vision_model"], "Qwen3.5-Flash")
        self.assertEqual(res["items"][0]["品名"], "鱼腥草")
        self.assertEqual(res["items"][0]["单位"], "斤")
        self.assertAlmostEqual(res["paid_total"], 12.0, places=2)
        self.assertIn("2997320402", res["order_ids"])

    def test_handwriting_does_not_call_vision(self):
        html = "\n".join([
            "<table><tr><td>编号</td><td>商品名称</td><td>单位</td><td>数量</td><td>单价</td><td>金额</td><td>备注</td></tr>",
            "<tr><td>1</td><td>白菜</td><td>斤</td><td>2</td><td>3</td><td>6</td><td></td></tr>",
            "</table>",
        ])
        self.assertFalse(core.is_meicai_ocr(html))

        def fake(_path, timeout=600, enhance=False):
            return html

        def boom(_path):
            raise AssertionError("手写单不应调用美菜视觉识别")

        old_full = core.call_ocr
        old_vision = webapp.recognize_meicai_image
        core.call_ocr = fake
        webapp.recognize_meicai_image = boom
        try:
            res = webapp.ocr_one("hand.jpg")
        finally:
            core.call_ocr = old_full
            webapp.recognize_meicai_image = old_vision
        self.assertEqual(res["kind"], "handwriting")
        self.assertEqual(res["items"][0]["品名"], "白菜")
        self.assertNotIn("vision_model", res)


class ConfirmSyncTests(unittest.TestCase):
    def setUp(self):
        self._paths = dict(webapp._paths)
        self._state = copy.deepcopy(webapp._state)
        self._state_file = webapp.STATE_FILE
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        out = os.path.join(self.tmp.name, "out")
        os.makedirs(out)
        webapp.STATE_FILE = os.path.join(self.tmp.name, "state.json")
        webapp._paths.clear()
        webapp._paths.update({
            "input_dir": self.tmp.name,
            "output_dir": out,
            "material_file": "",
        })
        webapp._state.clear()
        webapp._state["invoices"] = []
        self._sync = webapp.sync_to_material
        self.sync_calls = []

        def fake_sync(model):
            self.sync_calls.append(model)
            land = webapp.land_file()
            if os.path.exists(land):
                os.remove(land)
            return ([{"日期": "2026-09-03", "落地品名": "小葱"}], land + ".bak")

        webapp.sync_to_material = fake_sync

    def tearDown(self):
        webapp.sync_to_material = self._sync
        webapp.STATE_FILE = self._state_file
        webapp._paths.clear()
        webapp._paths.update(self._paths)
        webapp._state.clear()
        webapp._state.update(self._state)

    def _invoice(self, filename, kind):
        return {
            "filename": filename,
            "status": "pending",
            "ocr_done": True,
            "kind": kind,
            "date": "2026-09-03",
            "items": [{
                "品名": "小葱", "原品名": "小葱", "数量": 1, "单价": 8, "金额": 8,
                "单位": "斤", "渠道": "美菜" if kind == "meicai" else "",
            }],
            "order_total": 8,
            "note": "",
        }

    def _payload(self):
        return webapp.ConfirmPayload(
            date="2026-09-03",
            category="蔬菜",
            items=[{"品名": "小葱", "原品名": "小葱", "数量": 1, "单价": 8, "金额": 8, "单位": "斤"}],
            paid_total=8,
        )

    def test_handwriting_confirm_does_not_sync(self):
        material = os.path.join(self.tmp.name, "material.xlsx")
        with open(material, "wb") as handle:
            handle.write(b"placeholder")
        webapp._paths["material_file"] = material
        webapp._state["invoices"] = [self._invoice("hand.jpg", "handwriting")]
        result = webapp.confirm("hand.jpg", self._payload())
        self.assertEqual(result, {"ok": True})
        self.assertEqual(self.sync_calls, [])
        self.assertTrue(os.path.exists(webapp.land_file()))
        self.assertEqual(webapp._state["invoices"][0]["status"], "confirmed")

    def test_meicai_confirm_does_not_sync_or_touch_material(self):
        material = os.path.join(self.tmp.name, "material.xlsx")
        with open(material, "wb") as handle:
            handle.write(b"placeholder")
        webapp._paths["material_file"] = material
        webapp._state["invoices"] = [self._invoice("meicai.png", "meicai")]
        result = webapp.confirm("meicai.png", self._payload())
        self.assertEqual(result, {"ok": True})
        self.assertEqual(self.sync_calls, [])
        self.assertNotIn("material_sync", result)
        self.assertTrue(os.path.exists(webapp.land_file()))
        self.assertEqual(webapp._state["invoices"][0]["status"], "confirmed")
        with open(material, "rb") as handle:
            self.assertEqual(handle.read(), b"placeholder")

    def test_meicai_without_material_keeps_land_file(self):
        webapp._state["invoices"] = [self._invoice("meicai.png", "meicai")]
        result = webapp.confirm("meicai.png", self._payload())
        self.assertEqual(result, {"ok": True})
        self.assertNotIn("material_sync", result)
        self.assertEqual(self.sync_calls, [])
        self.assertTrue(os.path.exists(webapp.land_file()))
        self.assertEqual(webapp._state["invoices"][0]["status"], "confirmed")


if __name__ == "__main__":
    unittest.main()
