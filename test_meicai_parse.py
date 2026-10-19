# -*- coding: utf-8 -*-
"""美菜订单截图：实收解析、按日合并、长图分条。不调在线 OCR。

金标 samples/gold/meicai_2026-09-09.json 是复核落地后的口径，不是 OCR 像素猜测。
OCR 文本里没有「收」时，数量和金额留空并标存疑。
Shot1 / Shot2 金标同样是复核后的落地值。仓库里没有这两张的 OCR 文本时，
不对它们做「解析结果等于金标行」的断言；一旦把文本放到
samples/gold/ocr/meicai_shot1.ocr.txt 或 meicai_shot2.ocr.txt，回归会自动跑。
Shot2 的弱 OCR（没有「收」行）必须标存疑，不能把金标里的实收数编出来。
"""
import copy
import json
import os
import re
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
SHOT_GOLD = {
    "shot1": os.path.join(ROOT, "samples", "gold", "meicai_shot1.json"),
    "shot2": os.path.join(ROOT, "samples", "gold", "meicai_shot2.json"),
}
SHOT_OCR = {
    "shot1": os.path.join(ROOT, "samples", "gold", "ocr", "meicai_shot1.ocr.txt"),
    "shot2": os.path.join(ROOT, "samples", "gold", "ocr", "meicai_shot2.ocr.txt"),
}
SHOT3_OCR = os.path.join(ROOT, "samples", "gold", "ocr", "meicai_shot3.ocr.txt")
SHOT3_TILED_OCR = os.path.join(ROOT, "samples", "gold", "ocr", "meicai_shot3.tiled.ocr.txt")


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
        # 长图会再打一次整图，但整图不是美菜实收，分数更低，仍用分条原文。
        self.assertEqual(seen["full"], 1)
        self.assertEqual(res["kind"], "meicai")
        self.assertEqual(len(res["items"]), 7)
        self.assertAlmostEqual(res["items"][0]["数量"], 3.08, places=2)

    def test_ocr_one_prefers_full_when_it_has_more_recv(self):
        """分条几乎没有可用「收」、整图有实收时，用整图原文。不把订货数写成实收。"""
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        path_dir = tmp.name
        Image.new("RGB", (1080, 3000), "white").save(os.path.join(path_dir, "long.png"))
        old_paths = dict(webapp._paths)
        old_call = webapp.core.call_ocr
        old_tiled = webapp.core.call_ocr_tiled
        webapp._paths["input_dir"] = path_dir

        def tiled(path, timeout=600, overlap=0.15):
            return "\n".join([
                "美菜",
                "送达时间 2026-09-09 06:35:03",
                "订单号：3002160333 复制",
                "(56,141),(117,169)已完成",
                "美菜配送-已送达",
                "鸡毛菜 斤",
                "￥14.55/3.00斤",
                "订",
                "￥14.55/3.00斤",
                "付款信息",
                "商品多退少补合计￥163.25",
            ])

        def full(path, timeout=600, enhance=False):
            return _ocr_text()

        webapp.core.call_ocr_tiled = tiled
        webapp.core.call_ocr = full
        try:
            res = webapp.ocr_one("long.png")
        finally:
            webapp.core.call_ocr = old_call
            webapp.core.call_ocr_tiled = old_tiled
            webapp._paths.clear()
            webapp._paths.update(old_paths)
        self.assertEqual(res["kind"], "meicai")
        self.assertEqual(len(res["items"]), 7)
        self.assertAlmostEqual(res["items"][0]["数量"], 3.08, places=2)
        self.assertAlmostEqual(res["items"][0]["金额"], 19.59, places=2)

    def test_long_handwriting_stays_on_tiled_text(self):
        """两份都不像美菜时不换文本。手写单仍走原来的表解析。"""
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        path_dir = tmp.name
        Image.new("RGB", (1080, 3000), "white").save(os.path.join(path_dir, "long.png"))
        old_paths = dict(webapp._paths)
        old_call = webapp.core.call_ocr
        old_tiled = webapp.core.call_ocr_tiled
        webapp._paths["input_dir"] = path_dir
        with open(HAND_TXT, encoding="utf-8") as f:
            hand = f.read()

        def tiled(path, timeout=600, overlap=0.15):
            return hand

        def full(path, timeout=600, enhance=False):
            return "<table><tr><td>1</td><td>不该用这份</td><td>9</td><td>9</td><td>9</td></tr></table>"

        webapp.core.call_ocr_tiled = tiled
        webapp.core.call_ocr = full
        try:
            res = webapp.ocr_one("long.png")
        finally:
            webapp.core.call_ocr = old_call
            webapp.core.call_ocr_tiled = old_tiled
            webapp._paths.clear()
            webapp._paths.update(old_paths)
        self.assertEqual(res["kind"], "handwriting")
        names = [it["品名"] for it in res["items"]]
        self.assertIn("薄荷叶", names)
        self.assertNotIn("不该用这份", names)


def _read_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _shot_gold(name):
    return _read_json(SHOT_GOLD[name])


def _doc_from_gold(gold):
    items = []
    for row in gold["日物料采购_复核后期望行"]:
        items.append({
            "品名": row["name"],
            "数量": row["qty"],
            "金额": row["amt"],
            "单位": row["unit"],
            "渠道": "美菜",
            "订单号": "",
            "实收缺失": False,
        })
    return {
        "date": gold["delivery_date"],
        "paid": gold["paid"],
        "order_ids": list(gold["order_ids"]),
        "items": items,
        "suspects": [],
    }


def _assert_filled_matches_gold(test, parsed, gold):
    """「收」行齐全时，解析结果应落在复核金标上（单价按分四舍五入）。"""
    test.assertEqual(parsed["date"], gold["delivery_date"])
    test.assertEqual(parsed["paid"], gold["paid"])
    test.assertEqual(parsed["order_ids"], gold["order_ids"])
    rows = gold["日物料采购_复核后期望行"]
    test.assertEqual(len(parsed["items"]), len(rows))
    items, _stat = core.derive_price(parsed["items"])
    total = 0.0
    for got, want in zip(items, rows):
        test.assertEqual(got["品名"], want["name"])
        test.assertEqual(got["单位"], want["unit"])
        test.assertAlmostEqual(got["数量"], want["qty"], places=2)
        test.assertAlmostEqual(got["金额"], want["amt"], places=2)
        test.assertFalse(got.get("实收缺失"))
        total += got["金额"]
    test.assertAlmostEqual(round(total, 2), gold["paid"], places=2)
    test.assertEqual(gold["员工餐"]["amt"], 0)
    test.assertAlmostEqual(round(total, 2) - gold["paid"], 0, places=2)


def _assert_missing_receive_not_invented(test, parsed, gold, text):
    """没有「收」的行不许填金标实收，也不许填标题上的订货数。"""
    test.assertEqual(parsed["date"], gold["delivery_date"])
    if re.search(r"商品多退少补合计\s*[￥¥]\s*" + re.escape(f"{gold['paid']:.2f}"), text):
        test.assertEqual(parsed["paid"], gold["paid"])
    recv_n = len(re.findall(r"(?:^|\n)\s*收\s*[￥¥]", text))
    filled = [it for it in parsed["items"] if it.get("数量") is not None or it.get("金额") is not None]
    test.assertLessEqual(len(filled), recv_n)
    for it in parsed["items"]:
        if it.get("实收缺失") or recv_n == 0:
            test.assertIsNone(it.get("数量"))
            test.assertIsNone(it.get("金额"))
            test.assertTrue(it.get("实收缺失"))
    for oid in parsed["order_ids"]:
        test.assertIn(oid, text)
    if recv_n == 0:
        test.assertTrue(any("实收" in s for s in parsed["suspects"]))
        gold_amts = {round(float(row["amt"]), 2) for row in gold["日物料采购_复核后期望行"]}
        for it in parsed["items"]:
            test.assertNotIn(it.get("金额"), gold_amts)


class ShotGoldTargets(unittest.TestCase):
    """Shot1 / Shot2 是复核落地金标。9/9 OCR 文本仍是解析必须通过的那条。"""

    def test_fixtures_match_approved_landing_totals(self):
        shot1 = _shot_gold("shot1")
        shot2 = _shot_gold("shot2")
        for gold, day, paid, n_orders in (
            (shot1, "2026-09-02", 258.23, 4),
            (shot2, "2026-09-03", 258.99, 3),
        ):
            rows = gold["日物料采购_复核后期望行"]
            self.assertEqual(gold["delivery_date"], day)
            self.assertEqual(gold["paid"], paid)
            self.assertEqual(gold["员工餐"]["amt"], 0)
            self.assertTrue(gold["paid_matches_line_sum"])
            self.assertAlmostEqual(sum(r["amt"] for r in rows), paid, places=2)
            self.assertEqual(len(gold["order_ids"]), n_orders)
            self.assertEqual(len(gold["order_ids"]), len(set(gold["order_ids"])))
        # Shot2 同一张单里两行小葱，落地时不按品名并成一行
        scallions = [r for r in shot2["日物料采购_复核后期望行"] if r["name"] == "小葱"]
        self.assertEqual(len(scallions), 2)
        self.assertNotEqual(scallions[0]["qty"], scallions[1]["qty"])

    def test_different_delivery_dates_stay_separate_docs(self):
        merged = core.merge_meicai_docs([
            _doc_from_gold(_shot_gold("shot1")),
            _doc_from_gold(_shot_gold("shot2")),
        ])
        self.assertEqual([d["date"] for d in merged], ["2026-09-02", "2026-09-03"])
        self.assertAlmostEqual(merged[0]["paid"], 258.23, places=2)
        self.assertAlmostEqual(merged[1]["paid"], 258.99, places=2)
        self.assertEqual(len(merged[1]["items"]), 10)

    def test_shot2_weak_ocr_marks_receive_missing(self):
        """按 Shot2 金标备注构造的弱 OCR：没有「收」分行，数字是订货量，还漏了一个订单号。"""
        text = "\n".join([
            "美菜",
            "送达时间 2026-09-03 06:30:00",
            "￥258.99",
            "下单实付￥250.00",
            "订单号：2997320400 复制",
            "鸡毛菜 斤 ￥36.00/6.00斤",
            "订 ￥36.00/6.00斤",
            "黄瓜 二级 斤 ￥11.00/5.00斤",
            "订 ￥11.00/5.00斤",
            "多肉排骨粒 斤 ￥19.08/3.00斤",
            "订 ￥19.08/3.00斤",
            "鲜猪后腿瘦肉 纯瘦 斤 ￥48.00/5.00斤",
            "订 ￥48.00/5.00斤",
            "绿豆芽 斤 ￥3.00/3.00斤",
            "订 ￥3.00/3.00斤",
            "订单号：2997320402 复制",
            "小葱 优质 水洗 斤 ￥26.00/3.00斤",
            "订 ￥26.00/3.00斤",
            "小葱 优质 水洗 斤 ￥44.00/5.00斤",
            "订 ￥44.00/5.00斤",
            "香菜 斤 ￥19.00/3.00斤",
            "订 ￥19.00/3.00斤",
            "鲜鸡蛋 白壳 中码 托(30 个) ￥40.00/2.00托",
            "订 ￥40.00/2.00托",
            "鱼腥草 斤 ￥6.00/1.00斤",
            "订 ￥6.00/1.00斤",
            "付款信息",
            "直降 -￥5.00",
            "历史欠款 ￥2.42",
            "商品多退少补合计￥258.99",
        ])
        gold = _shot_gold("shot2")
        parsed = core.parse_meicai_text(text)
        self.assertTrue(core.is_meicai_ocr(text))
        _assert_missing_receive_not_invented(self, parsed, gold, text)
        self.assertEqual(parsed["order_ids"], ["2997320400", "2997320402"])
        self.assertNotIn("2997320401", parsed["order_ids"])
        names = [it["品名"] for it in parsed["items"]]
        self.assertEqual(names.count("小葱"), 2)
        self.assertIn("黄瓜 二级", names)
        self.assertIn("鲜鸡蛋 白壳", names)
        self.assertNotIn("直降", names)
        self.assertFalse(any("欠款" in n for n in names))
        # 订货数 5.00 / 11.00 不能当成黄瓜的实收（金标是 5.08 斤、11.84 元）
        cucumber = next(it for it in parsed["items"] if it["品名"] == "黄瓜 二级")
        self.assertIsNone(cucumber["数量"])
        self.assertIsNone(cucumber["金额"])

    def test_optional_shot_ocr_transcripts(self):
        """有 OCR 文本才和金标行对账。Shot2 若「收」行不够，只检查不编造。"""
        ran = False
        for name in ("shot1", "shot2"):
            path = SHOT_OCR[name]
            if not os.path.isfile(path):
                continue
            ran = True
            with open(path, encoding="utf-8") as f:
                text = f.read()
            gold = _shot_gold(name)
            parsed = core.parse_meicai_text(text)
            self.assertTrue(core.is_meicai_ocr(text), name)
            recv_n = len(re.findall(r"(?:^|\n)\s*收\s*[￥¥]", text))
            if recv_n >= len(gold["日物料采购_复核后期望行"]):
                _assert_filled_matches_gold(self, parsed, gold)
            else:
                _assert_missing_receive_not_invented(self, parsed, gold, text)
        if not ran:
            self.skipTest("shot1/shot2 的 OCR 文本还没放进 samples/gold/ocr/")


def _read_text(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


class MeicaiBareRecvAndNoise(unittest.TestCase):
    """长图分条 OCR：收和金额拆行、坐标贴在文案上、标价单独一行。"""

    def test_bare_recv_merges_next_line_qty(self):
        text = "\n".join([
            "美菜",
            "送达时间 2026-09-04 07:12:48",
            "订单号：2998060810 复制",
            "促 大娃娃菜 优质 斤",
            "￥6.90/3.00斤",
            "订",
            "￥6.90/3.00斤",
            "收",
            "￥6.99/3.04斤",
            "付款信息",
            "商品多退少补合计￥6.99",
        ])
        parsed = core.parse_meicai_text(text)
        self.assertEqual(len(parsed["items"]), 1)
        it = parsed["items"][0]
        self.assertEqual(it["品名"], "大娃娃菜")
        self.assertAlmostEqual(it["数量"], 3.04, places=2)
        self.assertAlmostEqual(it["金额"], 6.99, places=2)
        self.assertEqual(it["单位"], "斤")
        self.assertFalse(it["实收缺失"])
        self.assertNotAlmostEqual(it["数量"], 3.00, places=2)
        self.assertNotAlmostEqual(it["金额"], 6.90, places=2)

    def test_bare_recv_allows_spaces_and_fullwidth_yen(self):
        text = "\n".join([
            "订单号：2998060810 复制",
            "送达时间 2026-09-04 07:12:48",
            "美菜",
            "鸡毛菜 普通 斤",
            "  收  ",
            "¥6.54/3.00斤",
        ])
        it = core.parse_meicai_text(text)["items"][0]
        self.assertEqual(it["品名"], "鸡毛菜")
        self.assertAlmostEqual(it["数量"], 3.0, places=2)
        self.assertAlmostEqual(it["金额"], 6.54, places=2)
        self.assertFalse(it["实收缺失"])

    def test_noise_bbox_and_chrome_are_not_names(self):
        text = "\n".join([
            "美菜",
            "订单详情",
            "(56,141),(117,169)已完成",
            "(49,217),(537,236)",
            "美菜配送-已送达",
            "您的订单被拆分为1个包裹配送，点击查看更多物流信息",
            "送达时间 2026-09-04 07:12:48",
            "签收方式 信任签收 | 店门口",
            "司机姓名：苏师傅 联系司机",
            "退款详情",
            "订单号：2998060810 复制",
            "大娃娃菜 优质 斤",
            "收",
            "￥1.50/1.00斤",
            "付款信息",
            "商品多退少补合计￥1.50",
        ])
        parsed = core.parse_meicai_text(text)
        names = [it["品名"] for it in parsed["items"]]
        self.assertEqual(names, ["大娃娃菜"])
        self.assertAlmostEqual(parsed["items"][0]["数量"], 1.0, places=2)
        blob = "\n".join(names)
        for word in ("已完成", "美菜配送", "退款详情", "送达时间", "信任签收", "查看更多物流", "司机姓名", "订单详情"):
            self.assertNotIn(word, blob)
        self.assertIsNone(re.search(r"\(\s*\d+\s*,\s*\d+\s*\)", blob))
        self.assertEqual(parsed["date"], "2026-09-04")

    def test_missing_recv_keeps_name_when_price_is_on_next_line(self):
        text = "\n".join([
            "美菜",
            "送达时间 2026-09-04 07:12:48",
            "￥6.90",
            "订单号：2998060810 复制",
            "大娃娃菜 优质 斤",
            "￥6.90/3.00斤",
            "订",
            "￥6.90/3.00斤",
            "付款信息",
            "商品多退少补合计￥6.90",
        ])
        parsed = core.parse_meicai_text(text)
        self.assertEqual(len(parsed["items"]), 1)
        it = parsed["items"][0]
        self.assertEqual(it["品名"], "大娃娃菜")
        self.assertIsNone(it["数量"])
        self.assertIsNone(it["金额"])
        self.assertTrue(it["实收缺失"])
        self.assertTrue(any("实收" in s for s in parsed["suspects"]))
        self.assertNotEqual(it["数量"], 3.0)
        self.assertNotEqual(it["金额"], 6.9)

    def test_multi_segment_pending_kept_when_next_name_arrives(self):
        """连续两段「品名 + 下一行标价」都没有「收」，第三条品名不能把前两段盖掉。"""
        text = "\n".join([
            "美菜",
            "送达时间 2026-09-04 07:12:48",
            "订单号：2998060810 复制",
            "大娃娃菜 优质 斤",
            "￥6.90/3.00斤",
            "订",
            "￥6.90/3.00斤",
            "鸡毛菜 普通 斤",
            "￥6.54/3.00斤",
            "订",
            "￥6.54/3.00斤",
            "鲜鸡蛋 红壳 中码 托(30个)",
            "付款信息",
            "商品多退少补合计￥20.00",
        ])
        parsed = core.parse_meicai_text(text)
        self.assertEqual(
            [it["品名"] for it in parsed["items"]],
            ["大娃娃菜", "鸡毛菜", "鲜鸡蛋 红壳"],
        )
        for it in parsed["items"]:
            self.assertIsNone(it["数量"])
            self.assertIsNone(it["金额"])
            self.assertTrue(it["实收缺失"])
        self.assertNotEqual(parsed["items"][0]["数量"], 3.0)
        self.assertNotEqual(parsed["items"][0]["金额"], 6.9)
        self.assertNotEqual(parsed["items"][1]["数量"], 3.0)
        self.assertNotEqual(parsed["items"][1]["金额"], 6.54)
        self.assertTrue(any("实收" in s for s in parsed["suspects"]))
        self.assertEqual(parsed["paid"], 20.0)

    def test_shot3_tiled_fixture_fills_recv_and_drops_noise(self):
        text = _read_text(SHOT3_TILED_OCR)
        parsed = core.parse_meicai_text(text)
        self.assertTrue(core.is_meicai_ocr(text))
        self.assertEqual(parsed["date"], "2026-09-04")
        self.assertEqual(parsed["paid"], 87.34)
        cabbage = next(it for it in parsed["items"] if it["品名"] == "大娃娃菜")
        self.assertAlmostEqual(cabbage["数量"], 3.04, places=2)
        self.assertAlmostEqual(cabbage["金额"], 6.99, places=2)
        self.assertFalse(cabbage["实收缺失"])
        greens = next(it for it in parsed["items"] if it["品名"] == "鸡毛菜")
        self.assertAlmostEqual(greens["数量"], 3.0, places=2)
        self.assertAlmostEqual(greens["金额"], 6.54, places=2)
        egg = next(it for it in parsed["items"] if it["品名"] == "鲜鸡蛋 红壳")
        self.assertAlmostEqual(egg["数量"], 1.0, places=2)
        self.assertAlmostEqual(egg["金额"], 20.08, places=2)
        self.assertEqual(egg["单位"], "托")
        claw = next(it for it in parsed["items"] if it["品名"] == "冻鸡爪 L号 大爪/长爪")
        self.assertAlmostEqual(claw["数量"], 1.0, places=2)
        self.assertAlmostEqual(claw["金额"], 28.69, places=2)
        self.assertEqual(claw["单位"], "袋")
        herb = next(it for it in parsed["items"] if it["品名"] == "鱼腥草")
        self.assertAlmostEqual(herb["数量"], 1.0, places=2)
        self.assertAlmostEqual(herb["金额"], 6.18, places=2)
        beans = [it for it in parsed["items"] if it["品名"] == "优质黄豆" and it["金额"] == 4.66]
        self.assertEqual(len(beans), 1)
        self.assertEqual(beans[0]["单位"], "袋")
        names = [it["品名"] for it in parsed["items"]]
        blob = "\n".join(names)
        for word in ("已完成", "美菜配送", "退款详情", "送达时间", "信任签收", "查看更多物流", "司机姓名"):
            self.assertNotIn(word, blob)
        self.assertIsNone(re.search(r"\(\s*\d+\s*,\s*\d+\s*\)", blob))
        # 订货价 6.90/3.00 不是大娃娃菜的实收
        self.assertFalse(any(
            it["品名"] == "大娃娃菜" and it["金额"] == 6.9 and it["数量"] == 3.0
            for it in parsed["items"]
        ))

    def test_shot3_full_image_does_not_invent_recv(self):
        text = _read_text(SHOT3_OCR)
        parsed = core.parse_meicai_text(text)
        self.assertTrue(core.is_meicai_ocr(text))
        self.assertEqual(parsed["date"], "2026-09-04")
        self.assertEqual(parsed["paid"], 87.34)
        self.assertGreaterEqual(len(parsed["items"]), 8)
        self.assertTrue(any("大娃娃菜" in it["品名"] for it in parsed["items"]))
        self.assertNotIn("共9类", [it["品名"] for it in parsed["items"]])
        for it in parsed["items"]:
            self.assertTrue(it["品名"])
            self.assertIsNone(it["数量"])
            self.assertIsNone(it["金额"])
            self.assertTrue(it["实收缺失"])

    def test_quality_prefers_tiled_shot3_and_more_recv(self):
        tiled = _read_text(SHOT3_TILED_OCR)
        full = _read_text(SHOT3_OCR)
        self.assertGreater(core.meicai_ocr_quality(tiled), core.meicai_ocr_quality(full))
        self.assertIs(core.choose_meicai_ocr_text([tiled, full]), tiled)
        # 同分留前面的；整图实收更多时选整图。选的是原文，金额仍由解析读「收」。
        self.assertIs(core.choose_meicai_ocr_text([full, full]), full)
        weak = "\n".join([
            "美菜配送",
            "送达时间 2026-09-04 07:12:48",
            "订单号：2998060810 复制",
            "(1,2),(3,4)",
            "大娃娃菜 斤",
            "￥6.90/3.00斤",
            "多退少补",
        ])
        rich = "\n".join([
            "美菜",
            "送达时间 2026-09-04 07:12:48",
            "订单号：2998060810 复制",
            "大娃娃菜 斤 ￥6.90/3.00斤",
            "订 ￥6.90/3.00斤",
            "收 ￥6.99/3.04斤",
            "鸡毛菜 斤 ￥6.54/3.00斤",
            "收 ￥6.54/3.00斤",
            "商品多退少补合计￥13.53",
        ])
        self.assertGreater(core.meicai_ocr_quality(rich), core.meicai_ocr_quality(weak))
        chosen = core.choose_meicai_ocr_text([weak, rich])
        self.assertIs(chosen, rich)
        item = next(it for it in core.parse_meicai_text(chosen)["items"] if it["品名"] == "大娃娃菜")
        self.assertAlmostEqual(item["金额"], 6.99, places=2)
        self.assertAlmostEqual(item["数量"], 3.04, places=2)


if __name__ == "__main__":
    unittest.main()
