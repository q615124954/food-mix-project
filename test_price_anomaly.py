# -*- coding: utf-8 -*-
"""P0：数量丢小数点（鸡毛菜 5.8→58）要按单价异常标数量/金额存疑，不自动改数。

单价是 金额÷数量 反算的，58 斤 28 元反算出 0.48/斤，「数量×单价≠金额」永远不触发，
所以拿历史单价（物料表）或票面单价比，偏离超过 3 倍才标。
OCR 原文来自 samples/gold/ocr/（真实 HunyuanOCR 返回，可离线复现）。
"""
import os
import sys
import unittest

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(BASE, "webapp"))
import app as webapp
import pipeline_min as core

GOLD = os.path.join(BASE, "samples", "gold")


def _ocr_items(day):
    with open(os.path.join(GOLD, "ocr", f"receipt_2026-08-{day}.ocr.txt"), encoding="utf-8") as f:
        text = f.read()
    items = core.parse_line_items(core.parse_table(text))
    items, _ = core.derive_price(items)
    return items


class PriceAnomalyRule(unittest.TestCase):
    def test_hist_price_flags_lost_decimal(self):
        why = core.price_anomaly({"数量": 58, "金额": 28}, hist_price=4.83)
        self.assertIsNotNone(why)
        self.assertIn("小数点", why)

    def test_hist_price_ok_after_fix(self):
        self.assertIsNone(core.price_anomaly({"数量": 5.8, "金额": 28}, hist_price=4.83))

    def test_no_reference_no_flag(self):
        self.assertIsNone(core.price_anomaly({"数量": 58, "金额": 28}))

    def test_ticket_price_used_without_history(self):
        self.assertIsNotNone(core.price_anomaly({"数量": 58, "金额": 28, "票面单价": 4.8}))
        # 手写抹零（3×3.8=11.4 记 12）不算异常
        self.assertIsNone(core.price_anomaly({"数量": 3, "金额": 12, "票面单价": 3.8}))

    def test_history_overrides_bad_ticket_price(self):
        # 票面单价 3.9 被 OCR 成 39，但反算单价与历史一致 → 数量/金额没问题，不标
        self.assertIsNone(core.price_anomaly({"数量": 2, "金额": 8, "票面单价": 39}, hist_price=4.0))


class GoldReceipt0813(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._old = webapp._paths.get("material_file", "")
        webapp._paths["material_file"] = os.path.join(GOLD, "material_gold.xlsx")
        webapp.load_legal_names()

    @classmethod
    def tearDownClass(cls):
        webapp._paths["material_file"] = cls._old
        webapp.load_legal_names()

    def test_history_loaded(self):
        self.assertAlmostEqual(webapp._hist_prices["鸡毛菜"], 28 / 5.8, places=2)

    def test_jimaocai_flagged_only(self):
        items = _ocr_items("13")
        jm = [it for it in items if it["品名"] == "鸡毛菜"]
        self.assertEqual(len(jm), 1)
        self.assertEqual(jm[0]["数量"], 58)  # OCR 原值保留，不自动改
        suspects = webapp.compute_suspects("2026-08-13", items)
        self.assertTrue(jm[0]["数量存疑"])
        self.assertTrue(jm[0]["金额存疑"])
        self.assertIn("小数点", jm[0]["存疑说明"])
        self.assertTrue(any("单价异常" in s and "鸡毛菜" in s for s in suspects))
        # 其余有历史单价的物料行，数量/金额都不因单价被标（数量没回退）
        others = [it["品名"] for it in items
                  if it.get("存疑说明") and it["品名"] != "鸡毛菜"]
        self.assertEqual(others, [])

    def test_manual_fix_clears_flag(self):
        items = _ocr_items("13")
        jm = next(it for it in items if it["品名"] == "鸡毛菜")
        jm["数量"] = 5.8
        core.derive_price(items)
        webapp.compute_suspects("2026-08-13", items)
        self.assertFalse(jm["数量存疑"])
        self.assertAlmostEqual(jm["单价"], 4.83, places=2)

    def test_other_days_no_false_alarm(self):
        for day in ("11", "12"):
            items = _ocr_items(day)
            webapp.compute_suspects(f"2026-08-{day}", items)
            flagged = [it["品名"] for it in items
                       if it.get("存疑说明") and webapp.normalize_pin(it["品名"]) in webapp._legal_names]
            self.assertEqual(flagged, [], day)


if __name__ == "__main__":
    unittest.main()
