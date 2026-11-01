# -*- coding: utf-8 -*-
"""视觉对照脚本的离线形状。不调用模力方舟。"""
import json
import os
import subprocess
import sys
import unittest

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(ROOT, "scripts"))
sys.path.insert(0, ROOT)

import pipeline_min as core
import vision_meicai_compare as harness

SHOT1 = os.path.join(ROOT, "samples", "gold", "meicai_shot1.json")
SHOT3 = os.path.join(ROOT, "samples", "vision_exp", "fixture_shot3.json")
SHOT1_JPG = os.path.join(ROOT, "samples", "vision_exp", "shot1.jpg")


def _tiny_meicai_text():
    return "\n".join([
        "美菜",
        "送达时间 2026-09-02 10:00:00",
        "订单号：2996559029 复制",
        "小葱 斤",
        "订 ￥10.00/2.00斤",
        "收 ￥26.90/5.00斤",
        "商品多退少补合计￥26.90",
    ])


class ClassifyTests(unittest.TestCase):
    def test_prompt_asks_for_review_item_shape(self):
        for word in ("品名", "数量", "金额", "单位", "渠道", "订单号", "实收缺失", "票面单价"):
            self.assertIn(word, harness.VISION_PROMPT)

    def test_boxes_only_list_is_bad_interface(self):
        text = json.dumps([
            {"bbox": [0, 0, 10, 10], "text": "小葱"},
            {"polygon": [[0, 0], [1, 0], [1, 1]], "text": "收"},
        ], ensure_ascii=False)
        kind, _data = harness.classify_model_text(text)
        self.assertEqual(kind, "boxes")

    def test_boxes_wrapped_in_object(self):
        text = json.dumps({"boxes": [{"bounding_box": [1, 2, 3, 4], "text": "鸡毛菜"}]})
        self.assertEqual(harness.classify_model_text(text)[0], "boxes")

    def test_coordinate_dump_without_items(self):
        text = "(12, 30),(40, 50)已完成\n(1, 2),(3, 4)\n[0, 1, 2, 3]"
        self.assertEqual(harness.classify_model_text(text)[0], "boxes")

    def test_items_json_is_not_boxes(self):
        text = "```json\n" + json.dumps({
            "date": "2026-09-02",
            "paid": 26.9,
            "order_ids": ["2996559029"],
            "items": [{"品名": "小葱", "数量": 5, "金额": 26.9, "单位": "斤", "渠道": "美菜"}],
        }, ensure_ascii=False) + "\n```"
        kind, data = harness.classify_model_text(text)
        self.assertEqual(kind, "items")
        doc = harness.finalize_vision_doc(data)
        self.assertEqual(doc["items"][0]["品名"], "小葱")
        self.assertEqual(doc["items"][0]["渠道"], "美菜")
        self.assertEqual(doc["items"][0]["勾稽"], "OK")
        self.assertAlmostEqual(doc["paid"], 26.9)

    def test_items_win_when_boxes_also_present(self):
        text = json.dumps({
            "items": [{"品名": "香菜", "数量": 2, "金额": 13.7, "单位": "斤"}],
            "boxes": [{"bbox": [0, 0, 1, 1]}],
        })
        self.assertEqual(harness.classify_model_text(text)[0], "items")

    def test_empty_items_array_is_structured(self):
        self.assertEqual(harness.classify_model_text('{"items": [], "paid": null}')[0], "items")


class CompareTests(unittest.TestCase):
    def test_shot1_gold_matches_its_own_lines(self):
        gold = harness.load_gold(SHOT1)
        items = []
        for line in gold["lines"]:
            items.append({
                "品名": line["品名"], "数量": line["数量"], "金额": line["金额"], "单位": line["单位"],
            })
        scored = harness.compare_items(items, gold["paid"], gold)
        self.assertEqual(scored["line_hits"], scored["gold_lines"])
        self.assertEqual(scored["gold_lines"], 6)
        self.assertEqual(scored["fabricated_lines"], 0)
        self.assertTrue(scored["paid_match"])
        self.assertEqual(scored["missed_lines"], 0)

    def test_fabricated_line_and_paid_mismatch(self):
        gold = harness.load_gold(SHOT1)
        items = [{
            "品名": line["品名"], "数量": line["数量"], "金额": line["金额"], "单位": line["单位"],
        } for line in gold["lines"]]
        items.append({"品名": "不存在的货", "数量": 1, "金额": 9.9, "单位": "斤"})
        scored = harness.compare_items(items, 1, gold)
        self.assertEqual(scored["line_hits"], 6)
        self.assertEqual(scored["fabricated_lines"], 1)
        self.assertEqual(scored["fabricated"][0]["品名"], "不存在的货")
        self.assertFalse(scored["paid_match"])

    def test_shot3_fixture_null_qty_can_hit(self):
        gold = harness.load_gold(SHOT3)
        self.assertGreater(gold["lines"] and len(gold["lines"]), 0)
        self.assertTrue(all(line["数量"] is None and line["金额"] is None for line in gold["lines"]))
        scored = harness.compare_items(gold["lines"], gold["paid"], gold)
        self.assertEqual(scored["line_hits"], len(gold["lines"]))
        self.assertTrue(scored["paid_match"])

    def test_filled_qty_does_not_match_null_gold(self):
        gold = harness.load_gold(SHOT3)
        pred = [dict(line) for line in gold["lines"]]
        pred[0] = dict(pred[0])
        pred[0]["数量"] = 1
        pred[0]["金额"] = 2
        scored = harness.compare_items(pred, gold["paid"], gold)
        self.assertEqual(scored["line_hits"], len(gold["lines"]) - 1)
        self.assertEqual(scored["fabricated_lines"], 1)

    def test_baseline_text_uses_parse_meicai(self):
        doc = harness.baseline_doc_from_text(_tiny_meicai_text())
        self.assertEqual(doc["items"][0]["品名"], "小葱")
        self.assertAlmostEqual(doc["items"][0]["数量"], 5.0)
        self.assertAlmostEqual(doc["items"][0]["金额"], 26.9)
        self.assertEqual(doc["items"][0]["渠道"], "美菜")
        self.assertAlmostEqual(doc["paid"], 26.9)
        self.assertIn("单价", doc["items"][0])


class ReportShapeTests(unittest.TestCase):
    def test_missing_key_keeps_metrics_empty_and_does_not_call(self):
        calls = {"n": 0}

        def boom(*_args, **_kwargs):
            calls["n"] += 1
            raise AssertionError("should not call")

        report = harness.run_report(
            [(os.path.join(ROOT, "samples/vision_exp/shot1.jpg"), SHOT1)],
            list(harness.DEFAULT_MODELS),
            api_key="",
            baseline_caller=boom,
            vision_caller=boom,
        )
        self.assertEqual(calls["n"], 0)
        self.assertFalse(report["api_key_present"])
        self.assertFalse(report["contains_live_measurements"])
        self.assertFalse(report["vision_is_production_default"])
        self.assertEqual(report["production_default_model"], "HunyuanOCR")
        self.assertEqual(report["fallback_model"], "Qwen3.5-Flash")
        self.assertIn("未读取到", report["note"])
        runs = report["samples"][0]["runs"]
        self.assertEqual([run["model"] for run in runs], [
            "HunyuanOCR", "Qwen3.5-OCR", "Qwen3-VL-Plus", "Qwen3.5-Flash",
        ])
        self.assertEqual(runs[0]["endpoint"], "https://ai.gitee.com/v1/images/ocr")
        self.assertEqual(runs[1]["endpoint"], "https://ai.gitee.com/v1/chat/completions")
        for run in runs:
            self.assertEqual(run["status"], "skipped_no_api_key")
            self.assertIsNone(run["line_hits"])
            self.assertIsNone(run["paid_match"])
            self.assertIsNone(run["fabricated_lines"])
            self.assertIsNone(run["latency_ms"])
            self.assertIsNone(run["items"])
            self.assertEqual(run["gold_lines"], 6)
        markdown = harness.render_markdown(report)
        self.assertIn("不是 0", markdown)
        self.assertNotIn("| 0 |", markdown)

    def test_dry_run_with_key_still_skips_network(self):
        calls = {"n": 0}

        def boom(*_args, **_kwargs):
            calls["n"] += 1
            raise AssertionError("dry-run must not call")

        report = harness.run_report(
            [("missing.jpg", SHOT1)],
            ["Qwen3.5-OCR"],
            api_key="present-but-not-sent",
            key_source="HUNYUAN_API_KEY",
            dry_run=True,
            baseline_caller=boom,
            vision_caller=boom,
        )
        self.assertEqual(calls["n"], 0)
        self.assertTrue(report["api_key_present"])
        self.assertEqual(report["key_source"], "HUNYUAN_API_KEY")
        self.assertNotIn("present-but-not-sent", json.dumps(report))
        self.assertTrue(all(run["status"] == "dry_run" for run in report["samples"][0]["runs"]))

    def test_boxes_are_labeled_and_fallback_is_scored(self):
        def vision(_key, model, _image, _timeout):
            if model == "Qwen3.5-OCR":
                return {
                    "text": json.dumps([{"bbox": [0, 0, 8, 8], "text": "only-box"}]),
                    "latency_ms": 15,
                    "http_status": 200,
                    "error": "",
                }
            if model == "Qwen3-VL-Plus":
                return {
                    "text": json.dumps({
                        "paid": 26.9,
                        "items": [{"品名": "小葱", "数量": 5, "金额": 26.9, "单位": "斤"}],
                    }, ensure_ascii=False),
                    "latency_ms": 20,
                    "http_status": 200,
                    "error": "",
                }
            if model == "Qwen3.5-Flash":
                return {
                    "text": json.dumps({
                        "paid": 26.9,
                        "items": [{"品名": "小葱", "数量": 5, "金额": 26.9, "单位": "斤"}],
                    }, ensure_ascii=False),
                    "latency_ms": 30,
                    "http_status": 200,
                    "error": "",
                }
            raise AssertionError(model)

        def baseline(_image, _timeout, _key):
            return {"text": _tiny_meicai_text(), "route": "hunyuan_tiled", "latency_ms": 5, "error": ""}

        gold_path = self._write_gold()
        report = harness.run_report(
            [(SHOT1_JPG, gold_path)],
            ["Qwen3.5-OCR", "Qwen3-VL-Plus"],
            api_key="k",
            baseline_caller=baseline,
            vision_caller=vision,
        )
        runs = {run["model"]: run for run in report["samples"][0]["runs"]}
        self.assertEqual(runs["Qwen3.5-OCR"]["status"], "接口形态不合格")
        self.assertEqual(runs["Qwen3.5-OCR"]["label"], "接口形态不合格")
        self.assertIsNone(runs["Qwen3.5-OCR"]["line_hits"])
        self.assertEqual(runs["Qwen3.5-OCR"]["latency_ms"], 15)
        self.assertEqual(runs["Qwen3-VL-Plus"]["status"], "ok")
        self.assertEqual(runs["Qwen3-VL-Plus"]["line_hits"], 1)
        self.assertEqual(runs["Qwen3-VL-Plus"]["fabricated_lines"], 0)
        self.assertTrue(runs["Qwen3-VL-Plus"]["paid_match"])
        self.assertEqual(runs["Qwen3.5-Flash"]["role"], "fallback")
        self.assertEqual(runs["Qwen3.5-Flash"]["status"], "ok")
        self.assertEqual(runs["Qwen3.5-Flash"]["line_hits"], 1)
        self.assertEqual(runs["HunyuanOCR"]["route"], "hunyuan_tiled")
        self.assertEqual(runs["HunyuanOCR"]["line_hits"], 1)
        self.assertTrue(report["contains_live_measurements"])

    def test_fallback_not_called_when_ocr_returns_items(self):
        seen = []

        def vision(_key, model, _image, _timeout):
            seen.append(model)
            return {
                "text": json.dumps({"paid": 26.9, "items": [
                    {"品名": "小葱", "数量": 5, "金额": 26.9, "单位": "斤"},
                ]}),
                "latency_ms": 1,
                "http_status": 200,
                "error": "",
            }

        report = harness.run_report(
            [(SHOT1_JPG, self._write_gold())],
            ["Qwen3.5-OCR"],
            api_key="k",
            include_baseline=False,
            baseline_caller=lambda *_a, **_k: None,
            vision_caller=vision,
        )
        self.assertEqual(seen, ["Qwen3.5-OCR"])
        fallback = report["samples"][0]["runs"][-1]
        self.assertEqual(fallback["status"], "not_run")
        self.assertIsNone(fallback["line_hits"])

    def test_vl_boxes_are_labeled_without_fallback(self):
        def vision(_key, model, _image, _timeout):
            return {
                "text": json.dumps({"boxes": [{"bbox": [1, 1, 2, 2]}]}),
                "latency_ms": 4,
                "http_status": 200,
                "error": "",
            }

        report = harness.run_report(
            [(SHOT1_JPG, self._write_gold())],
            ["Qwen3-VL-Plus"],
            api_key="k",
            include_baseline=False,
            vision_caller=vision,
        )
        vl, fallback = report["samples"][0]["runs"]
        self.assertEqual(vl["status"], "接口形态不合格")
        self.assertEqual(fallback["status"], "not_run")

    def _write_gold(self):
        import tempfile
        handle = tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8")
        json.dump({
            "paid": 26.9,
            "日物料采购_复核后期望行": [
                {"name": "小葱", "unit": "斤", "qty": 5, "amt": 26.9},
            ],
        }, handle, ensure_ascii=False)
        handle.close()
        self.addCleanup(lambda: os.remove(handle.name))
        return handle.name


class BaselineRouteTests(unittest.TestCase):
    def test_short_image_does_not_tile(self):
        calls = []

        def tiled(path, timeout=600, overlap=None):
            calls.append("tiled")
            return "tiled"

        def full(path, timeout=600, enhance=False):
            calls.append("full")
            return "full"

        self._patch(tiled, full, lambda _path: False)
        text, route = harness.hunyuan_meicai_text("x.jpg", 3)
        self.assertEqual((text, route, calls), ("full", "hunyuan_full_frame", ["full"]))

    def test_long_screenshot_keeps_tiled_when_full_has_no_extra_recv(self):
        calls = []
        tiled_text = _tiny_meicai_text()

        def tiled(path, timeout=600, overlap=None):
            calls.append("tiled")
            return tiled_text

        def full(path, timeout=600, enhance=False):
            calls.append("full")
            return "美菜\n订单号：2996559029\n小葱 斤\n"

        self._patch(tiled, full, lambda _path: True)
        text, route = harness.hunyuan_meicai_text("x.jpg", 3)
        self.assertEqual(route, "hunyuan_tiled")
        self.assertEqual(text, tiled_text)
        self.assertEqual(calls, ["tiled", "full"])

    def test_tiled_failure_falls_back_to_full_frame(self):
        def tiled(path, timeout=600, overlap=None):
            raise RuntimeError("slice failed")

        def full(path, timeout=600, enhance=False):
            return "full-text"

        self._patch(tiled, full, lambda _path: True)
        text, route = harness.hunyuan_meicai_text("x.jpg", 3)
        self.assertEqual((text, route), ("full-text", "hunyuan_full_frame"))

    def _patch(self, tiled, full, looks):
        orig = (core.call_ocr_tiled, core.call_ocr, core.looks_like_long_screenshot)
        core.call_ocr_tiled = tiled
        core.call_ocr = full
        core.looks_like_long_screenshot = looks
        self.addCleanup(self._restore, orig)

    def _restore(self, orig):
        core.call_ocr_tiled, core.call_ocr, core.looks_like_long_screenshot = orig


class ProductionGuardTests(unittest.TestCase):
    def test_production_files_do_not_call_this_harness(self):
        self.assertEqual(core.MODEL, "HunyuanOCR")
        self.assertEqual(core.API_URL, "https://ai.gitee.com/v1/images/ocr")
        for rel in ("pipeline_min.py", "webapp/app.py"):
            with open(os.path.join(ROOT, rel), encoding="utf-8") as handle:
                text = handle.read()
            self.assertNotIn("vision_meicai_compare", text)
            self.assertNotIn("Qwen3.5-OCR", text)
            self.assertNotIn("Qwen3-VL-Plus", text)

    def test_cli_dry_run_exits_zero(self):
        env = os.environ.copy()
        env.pop("HUNYUAN_API_KEY", None)
        env.pop("MOLLI_ARK_API_KEY", None)
        proc = subprocess.run(
            [sys.executable, os.path.join(ROOT, "scripts", "vision_meicai_compare.py"), "--dry-run"],
            cwd=ROOT,
            env=env,
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertIn("skipped_no_api_key", proc.stdout)
        self.assertIn("Qwen3.5-Flash", proc.stdout)
        self.assertIn("shot1.jpg", proc.stdout)
        self.assertNotIn("| 0 |", proc.stdout)


if __name__ == "__main__":
    unittest.main()
