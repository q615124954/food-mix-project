# -*- coding: utf-8 -*-
"""美菜订单截图：实收解析、按日合并、长图分条。不调在线 OCR。

金标 samples/gold/meicai_2026-09-09.json 是复核落地后的口径，不是 OCR 像素猜测。
OCR 文本里没有「收」时，数量和金额留空并标存疑。
"""
import copy
import json
import os
import sys
import tempfile
import unittest

from PIL import Image

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(ROOT, "webapp"))
sys.path.insert(0, ROOT)

import pipeline_min as core
import app as webapp
from fastapi import HTTPException
from openpyxl import load_workbook

GOLD_JSON = os.path.join(ROOT, "samples", "gold", "meicai_2026-09-09.json")
OCR_TXT = os.path.join(ROOT, "samples", "gold", "ocr", "meicai_2026-09-09.ocr.txt")
HAND_TXT = os.path.join(ROOT, "samples", "gold", "ocr", "receipt_2026-08-11.ocr.txt")


def _ocr_text():
    with open(OCR_TXT, encoding="utf-8") as f:
        return f.read()


def _gold():
    with open(GOLD_JSON, encoding="utf-8") as f:
        return json.load(f)


class MeicaiParse(unittest.TestCase):
    def test_gold_rows_qty_amt_date_paid(self):
        parsed = core.parse_meicai_text(_ocr_text())
        gold = _gold()
        self.assertTrue(core.is_meicai_ocr(_ocr_text()))
        self.assertEqual(parsed["date"], gold["delivery_date"])
        self.assertEqual(parsed["paid"], gold["paid"])
        self.assertEqual(parsed["order_total"], gold["paid"])
        self.assertEqual(parsed["order_ids"], gold["order_ids"])
        self.assertIn("3002160333", parsed["note"])
        self.assertIn("3002160335", parsed["note"])
        rows = gold["日物料采购_复核后期望行"]
        self.assertEqual(len(parsed["items"]), len(rows))
        items, _stat = core.derive_price(parsed["items"])
        total = 0.0
        for got, want in zip(items, rows):
            self.assertEqual(got["品名"], want["name"])
            self.assertEqual(got["单位"], want["unit"])
            self.assertAlmostEqual(got["数量"], want["qty"], places=2)
            self.assertAlmostEqual(got["金额"], want["amt"], places=2)
            self.assertAlmostEqual(got["单价"], round(want["price"], 2), places=2)
            self.assertFalse(got.get("实收缺失"))
            self.assertFalse(got.get("员工区"))
            self.assertEqual(got.get("渠道"), "美菜")
            total += got["金额"]
        self.assertAlmostEqual(round(total, 2), gold["paid"], places=2)
        self.assertAlmostEqual(round(total, 2) - gold["paid"], gold["员工餐"]["amt"], places=2)
        names = [it["品名"] for it in items]
        self.assertNotIn("直降", names)
        self.assertFalse(any("欠款" in n for n in names))
        self.assertFalse(any("补差" in n or n.startswith("+") for n in names))
        # 实收，不是下单 3.00 斤 / 19.08 元
        self.assertAlmostEqual(items[0]["数量"], 3.08, places=2)
        self.assertAlmostEqual(items[0]["金额"], 19.59, places=2)
        self.assertNotAlmostEqual(items[0]["金额"], 19.08, places=2)

    def test_handwriting_html_is_not_meicai(self):
        with open(HAND_TXT, encoding="utf-8") as f:
            text = f.read()
        self.assertFalse(core.is_meicai_ocr(text))
        items = core.parse_line_items(core.parse_table(text))
        self.assertGreater(len(items), 0)
        self.assertEqual(core.parse_meicai_text(text)["items"], [])

    def test_missing_receive_line_does_not_copy_order_qty(self):
        text = "\n".join([
            "美菜",
            "送达时间 2026-09-10 06:00:00",
            "￥10.00",
            "下单实付￥9.00+需补交金额￥1.00",
            "订单号：3002160999 复制",
            "小葱 优质 水洗 斤 ￥10.00/2.00斤",
            "订 ￥10.00/2.00斤",
            "付款信息",
            "商品多退少补合计￥10.00",
            "历史欠款 ￥3.00",
            "直降 -￥1.00",
        ])
        parsed = core.parse_meicai_text(text)
        self.assertEqual(len(parsed["items"]), 1)
        it = parsed["items"][0]
        self.assertEqual(it["品名"], "小葱")
        self.assertIsNone(it["数量"])
        self.assertIsNone(it["金额"])
        self.assertTrue(it["实收缺失"])
        self.assertEqual(parsed["paid"], 10.0)
        self.assertTrue(any("实收" in s for s in parsed["suspects"]))
        self.assertNotEqual(parsed["paid"], 9.0)
        self.assertNotEqual(parsed["paid"], 12.0)

    def test_day_merge_concat_sum_and_conflict(self):
        a = core.parse_meicai_text(_ocr_text())
        b = {
            "date": "2026-09-09",
            "paid": 20.0,
            "order_ids": ["3002160336"],
            "items": [{
                "品名": "小葱", "数量": 1.0, "金额": 7.0, "单位": "斤",
                "渠道": "美菜", "订单号": "3002160336", "实收缺失": False,
            }],
            "suspects": [],
        }
        merged = core.merge_meicai_docs([a, b])
        self.assertEqual(len(merged), 1)
        doc = merged[0]
        self.assertEqual(len(doc["items"]), 8)
        self.assertEqual(doc["items"][-1]["品名"], "小葱")
        self.assertAlmostEqual(doc["paid"], 183.25, places=2)
        self.assertEqual(doc["order_ids"], a["order_ids"] + ["3002160336"])
        # 同名不合并
        self.assertEqual(sum(1 for it in doc["items"] if it["品名"] == "小葱"), 2)

        again = core.merge_meicai_docs([a, a])
        self.assertEqual(len(again[0]["items"]), len(a["items"]))
        self.assertAlmostEqual(again[0]["paid"], a["paid"], places=2)

        conflict = {
            "date": "2026-09-09",
            "paid": 5.0,
            "order_ids": ["3002160333"],
            "items": [{
                "品名": "多肉排骨粒", "数量": 1.0, "金额": 5.0, "单位": "斤",
                "渠道": "美菜", "订单号": "3002160333",
            }],
            "suspects": [],
        }
        bad = core.merge_meicai_docs([a, conflict])
        self.assertEqual(len(bad[0]["items"]), len(a["items"]) + 1)
        self.assertTrue(any("3002160333" in s and "不一致" in s for s in bad[0]["suspects"]))
        self.assertEqual(bad[0]["order_ids"].count("3002160333"), 1)

        other_day = {
            "date": "2026-09-10",
            "paid": 1.0,
            "order_ids": ["9"],
            "items": [{"品名": "盐", "数量": 1, "金额": 1, "订单号": "9"}],
            "suspects": [],
        }
        split = core.merge_meicai_docs([a, other_day])
        self.assertEqual([d["date"] for d in split], ["2026-09-09", "2026-09-10"])


class MeicaiRoute(unittest.TestCase):
    def setUp(self):
        self._call = webapp.core.call_ocr
        self._tiled = webapp.core.call_ocr_tiled
        self._state = copy.deepcopy(webapp._state)
        self._paths = dict(webapp._paths)

    def tearDown(self):
        webapp.core.call_ocr = self._call
        webapp.core.call_ocr_tiled = self._tiled
        webapp._state.clear()
        webapp._state.update(self._state)
        webapp._paths.clear()
        webapp._paths.update(self._paths)

    def test_ocr_one_routes_meicai_and_skips_enhance(self):
        calls = []

        def fake(path, timeout=600, enhance=False):
            calls.append(enhance)
            return _ocr_text()

        webapp.core.call_ocr = fake
        res = webapp.ocr_one("meicai_2026-09-09.png")
        self.assertEqual(calls, [False])
        self.assertEqual(res["kind"], "meicai")
        self.assertEqual(res["date"], "2026-09-09")
        self.assertEqual(res["order_total"], 163.25)
        self.assertEqual(res["paid_total"], 163.25)
        self.assertEqual(len(res["items"]), 7)
        gold = _gold()["日物料采购_复核后期望行"]
        for got, want in zip(res["items"], gold):
            self.assertEqual(got.get("原品名") or got["品名"], want["name"])
            self.assertAlmostEqual(got["数量"], want["qty"], places=2)
            self.assertAlmostEqual(got["金额"], want["amt"], places=2)
            self.assertNotEqual(got.get("归属"), "员工餐")
        self.assertEqual(res["meicai_part"]["order_ids"], _gold()["order_ids"])

    def test_day_merge_on_invoices_and_land_trace(self):
        text = _ocr_text()
        first = core.parse_meicai_text(text)
        second = {
            "date": "2026-09-09",
            "paid": 20.0,
            "order_ids": ["3002160999"],
            "note": "订单号 3002160999",
            "items": [{
                "品名": "生菜", "数量": 2.0, "金额": 20.0, "单位": "斤",
                "渠道": "美菜", "订单号": "3002160999", "实收缺失": False, "日期": "2026-09-09",
            }],
            "suspects": [],
        }
        parsed = [first, second]
        invoices = []
        for i, doc in enumerate(parsed):
            items, _ = core.derive_price(doc["items"])
            invoices.append({
                "filename": f"shot-{i}.png",
                "status": "pending",
                "ocr_done": True,
                "kind": "meicai",
                "date": doc["date"],
                "items": items,
                "order_total": doc["paid"],
                "paid_total": doc["paid"],
                "order_ids": doc["order_ids"],
                "note": doc["note"],
                "suspects": [],
                "meicai_part": {
                    "date": doc["date"],
                    "paid": doc["paid"],
                    "order_ids": doc["order_ids"],
                    "items": copy.deepcopy(items),
                    "note": doc["note"],
                    "suspects": doc["suspects"],
                },
            })
        webapp.apply_meicai_day_merge(invoices)
        self.assertEqual(invoices[0]["status"], "pending")
        self.assertEqual(len(invoices[0]["items"]), 8)
        self.assertEqual(invoices[0]["items"][-1]["品名"], "生菜")
        self.assertEqual(invoices[1]["status"], "merged")
        self.assertEqual(invoices[1]["items"], [])
        self.assertIn("3002160333", invoices[0]["note"])
        self.assertIn("3002160999", invoices[0]["order_ids"])
        self.assertEqual(invoices[0]["order_ids"].count("3002160333"), 1)
        self.assertAlmostEqual(invoices[0]["paid_total"], 183.25, places=2)

        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        webapp._paths["output_dir"] = tmp.name
        webapp.append_land(
            "shot-0.png", "2026-09-09", "蔬菜",
            invoices[0]["total"], invoices[0]["paid_total"], invoices[0]["items"],
            note=invoices[0]["note"],
        )
        wb = load_workbook(os.path.join(tmp.name, "采购入库_落地.xlsx"))
        ws = wb.active
        self.assertEqual(ws.cell(1, 9).value, "实付合计")
        self.assertEqual(ws.cell(1, 13).value, "溯源")
        self.assertEqual(ws.cell(2, 3).value, "多肉排骨粒")
        self.assertAlmostEqual(ws.cell(2, 4).value, 3.08, places=2)
        self.assertAlmostEqual(ws.cell(2, 6).value, 19.59, places=2)
        self.assertIn("3002160333", str(ws.cell(2, 13).value))
        self.assertEqual(ws.max_row, 9)  # header + 7 goods + 并入的生菜
        wb.close()

    def test_confirm_rejects_merged_screenshot(self):
        webapp._state["invoices"] = [{
            "filename": "shot-1.png",
            "status": "merged",
            "ocr_done": True,
            "kind": "meicai",
            "date": "2026-09-09",
            "items": [],
            "note": "已并入 shot-0.png",
        }]
        with self.assertRaises(HTTPException) as ctx:
            webapp.confirm("shot-1.png", webapp.ConfirmPayload(
                date="2026-09-09", category="蔬菜",
                items=[{"品名": "小葱", "数量": 1, "金额": 1}],
                paid_total=1,
            ))
        self.assertEqual(ctx.exception.status_code, 400)
        self.assertIn("并入", ctx.exception.detail)


class LongScreenshot(unittest.TestCase):
    def test_slices_cover_ends_without_exceeding_edge(self):
        img = Image.new("RGB", (1080, 4926), "white")
        img.putpixel((0, 0), (255, 0, 0))
        img.putpixel((10, 4925), (0, 255, 0))
        slices = list(core.iter_vertical_slices(img, max_edge=2048, overlap=0.15))
        self.assertGreaterEqual(len(slices), 3)
        self.assertTrue(all(max(s.size) <= 2048 for s in slices))
        self.assertTrue(any(s.getpixel((0, 0)) == (255, 0, 0) for s in slices))
        self.assertTrue(any(s.getpixel((10, s.size[1] - 1)) == (0, 255, 0) for s in slices))

    def test_join_drops_exact_overlap(self):
        joined = core.join_ocr_slices([
            "订单号：1\n收 ￥1.00/1.00斤\n鸡毛菜",
            "鸡毛菜\n收 ￥2.00/2.00斤",
        ])
        self.assertEqual(joined, "订单号：1\n收 ￥1.00/1.00斤\n鸡毛菜\n收 ￥2.00/2.00斤")

    def test_geometry_gate(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        tall = os.path.join(tmp.name, "tall.png")
        short = os.path.join(tmp.name, "short.png")
        Image.new("RGB", (1080, 4926), "white").save(tall)
        Image.new("RGB", (1600, 2000), "white").save(short)
        self.assertTrue(core.looks_like_long_screenshot(tall))
        self.assertFalse(core.looks_like_long_screenshot(short))
        self.assertFalse(core.looks_like_long_screenshot(os.path.join(tmp.name, "missing.png")))

    def test_ocr_one_uses_tiled_text_for_long_image(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        path_dir = tmp.name
        Image.new("RGB", (1080, 3000), "white").save(os.path.join(path_dir, "long.png"))
        old_paths = dict(webapp._paths)
        old_call = webapp.core.call_ocr
        old_tiled = webapp.core.call_ocr_tiled
        webapp._paths["input_dir"] = path_dir
        seen = {"tiled": 0, "full": 0}

        def tiled(path, timeout=600, overlap=0.15):
            seen["tiled"] += 1
            return _ocr_text()

        def full(path, timeout=600, enhance=False):
            seen["full"] += 1
            return "<table><tr><td>1</td><td>白菜</td><td>1</td><td>2</td><td>2</td></tr></table>"

        webapp.core.call_ocr_tiled = tiled
        webapp.core.call_ocr = full
        try:
            res = webapp.ocr_one("long.png")
        finally:
            webapp.core.call_ocr = old_call
            webapp.core.call_ocr_tiled = old_tiled
            webapp._paths.clear()
            webapp._paths.update(old_paths)
        self.assertEqual(seen["tiled"], 1)
        self.assertEqual(seen["full"], 0)
        self.assertEqual(res["kind"], "meicai")
        self.assertEqual(len(res["items"]), 7)


if __name__ == "__main__":
    unittest.main()
