# -*- coding: utf-8 -*-
"""P1：金标里出现过的 OCR 错名要归到正确品名；认不出的、纯数字品名、合计缺失/不符都要标出来交人工。

OCR 原文来自 samples/gold/ocr/（真实 HunyuanOCR 返回，可离线复现）。
"""
import copy
import os
import sys
import unittest

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(BASE, "webapp"))
import app as webapp
import name_match
import pipeline_min as core

GOLD = os.path.join(BASE, "samples", "gold")
CANDS = ["香葱", "香菜", "娃娃菜", "金针菇", "猪腱子肉", "去皮蒜米", "蒜头", "薄荷叶", "细面"]


def _rows(day):
    with open(os.path.join(GOLD, "ocr", f"receipt_2026-08-{day}.ocr.txt"), encoding="utf-8") as f:
        return core.parse_table(f.read())


def _run(day):
    """按 ocr_one 的顺序跑：解析 → 单价 → 品名纠错 → 存疑 → 合计校验（不调 OCR/LLM）"""
    rows = _rows(day)
    items, _ = core.derive_price(core.parse_line_items(rows))
    for it in items:
        raw = it["品名"]
        it["原品名"] = raw
        if raw:
            best, conf, src, _top = name_match.match(
                raw, webapp._legal_name_list, category="蔬菜",
                units=webapp._legal_units, unit=it.get("单位"))
            if src == "variant" or src.startswith("exact"):
                it["品名"] = best
        it["纠正来源"] = ""
    suspects = webapp.compute_suspects(f"2026-08-{day}", items)
    tip = webapp.order_total_suspect(core.sum_amount(items), core.extract_order_total(rows))
    if tip:
        suspects.append(tip)
    return items, suspects


class DefaultVariants(unittest.TestCase):
    def setUp(self):
        self._old = name_match._VARIANTS
        name_match._VARIANTS = copy.deepcopy(name_match.DEFAULT_VARIANTS)

    def tearDown(self):
        name_match._VARIANTS = self._old

    def test_gold_mappings(self):
        cases = {"香蕊": "香葱", "香芯": "香葱", "搭娃亲": "娃娃菜", "全针蕊": "金针菇",
                 "腱肉": "猪腱子肉", "腊肉": "猪腱子肉", "蒜半": "去皮蒜米",
                 "大蒜头": "蒜头", "薄荷": "薄荷叶", "细面条": "细面"}
        for raw, want in cases.items():
            best, conf, src, _ = name_match.match(raw, CANDS)
            self.assertEqual((best, src), (want, "variant"), raw)

    def test_real_larou_not_rewritten(self):
        best, conf, src, _ = name_match.match("腊肉", CANDS + ["腊肉"])
        self.assertEqual((best, src), ("腊肉", "exact"))

    def test_xiangcai_not_from_xiangrui(self):
        self.assertNotIn("香蕊", name_match.DEFAULT_VARIANTS["蔬菜"]["香菜"])
        best, _, _, _ = name_match.match("香菜", CANDS)
        self.assertEqual(best, "香菜")

    def test_migrate_old_json(self):
        old = {"蔬菜": {"香菜": ["香蕊", "香芯", "用户学的"]}, "肉类": {}}
        self.assertTrue(name_match._migrate(old))
        self.assertEqual(old["蔬菜"]["香菜"], ["用户学的"])
        self.assertIn("香蕊", old["蔬菜"]["香葱"])
        self.assertIn("腊肉", old["蔬菜"]["猪腱子肉"])
        self.assertFalse(name_match._migrate(old))


class ParseAndTotals(unittest.TestCase):
    def test_numeric_name_kept_and_blank_name_kept(self):
        rows = [["5", "10", "500克", "10", "1.5", "15", ""],
                ["6", "", "斤", "2", "3", "6", ""]]
        items = core.parse_line_items(rows)
        self.assertEqual([(i["品名"], i["数量"], i["金额"]) for i in items],
                         [("10", 10, 15), ("", 2, 6)])
        suspects = webapp.compute_suspects("2026-08-13", items)
        self.assertTrue(all(i["品名存疑"] for i in items))
        self.assertTrue(any("纯数字" in s for s in suspects))

    def test_order_total_forms(self):
        self.assertEqual(core.extract_order_total(_rows("11")), 141)   # 计110 … 计141 取最后
        self.assertEqual(core.extract_order_total(_rows("13")), 213)   # 总没 168 … 计:213 取最后
        self.assertIsNone(core.extract_order_total(_rows("12")))       # 没写「计」，不猜
        self.assertEqual(core.extract_order_total([["合计", "", "", "99.5"]]), 99.5)

    def test_order_total_suspect(self):
        self.assertIn("未识别", webapp.order_total_suspect(71, None))
        self.assertIn("不符", webapp.order_total_suspect(212, 213))
        self.assertIsNone(webapp.order_total_suspect(141, 141))


class GoldReceipts(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._old_v = name_match._VARIANTS
        name_match._VARIANTS = copy.deepcopy(name_match.DEFAULT_VARIANTS)
        cls._old = webapp._paths.get("material_file", "")
        webapp._paths["material_file"] = os.path.join(GOLD, "material_gold.xlsx")
        webapp.load_legal_names()

    @classmethod
    def tearDownClass(cls):
        name_match._VARIANTS = cls._old_v
        webapp._paths["material_file"] = cls._old
        webapp.load_legal_names()

    def _get(self, items, raw):
        return [i for i in items if i["原品名"] == raw]

    def test_0811(self):
        items, suspects = _run("11")
        self.assertEqual(self._get(items, "香蕊")[0]["品名"], "香葱")
        self.assertEqual(self._get(items, "搭娃亲")[0]["品名"], "娃娃菜")
        self.assertEqual(self._get(items, "全针蕊")[0]["品名"], "金针菇")
        self.assertEqual(self._get(items, "腱肉")[0]["品名"], "猪腱子肉")

    def test_0812(self):
        items, suspects = _run("12")
        names = {i["原品名"]: i["品名"] for i in items}
        self.assertEqual(names["香蕊"], "香葱")
        self.assertEqual(names["香菜"], "香菜")
        self.assertEqual(names["大蒜头"], "蒜头")
        self.assertEqual(names["蒜半"], "去皮蒜米")
        self.assertTrue(any("订单合计未识别" in s for s in suspects))

    def test_0813(self):
        items, suspects = _run("13")
        names = {i["原品名"]: i["品名"] for i in items}
        self.assertEqual(names["香芯"], "香葱")
        self.assertEqual(names["香菜"], "香菜")
        self.assertEqual(names["腊肉"], "猪腱子肉")
        self.assertEqual(names["薄荷"], "薄荷叶")
        self.assertEqual(names["细面条"], "细面")
        salt = self._get(items, "10")
        self.assertEqual(len(salt), 1)
        self.assertTrue(salt[0]["品名存疑"])
        self.assertEqual((salt[0]["数量"], salt[0]["金额"]), (10, 15))
        self.assertTrue(any("纯数字" in s for s in suspects))
        self.assertTrue(any("订单合计 213" in s and "不符" in s for s in suspects))
        # 金标里出现过的物料错名都纠正了；员工餐那几行（五花肉、油麦菜…）不在物料表，照样标品名存疑
        legal = [i for i in items if webapp.normalize_pin(i["品名"]) in webapp._legal_names]
        self.assertEqual([i["品名"] for i in legal if i["品名存疑"]], [])
        self.assertTrue(self._get(items, "五花肉")[0]["品名存疑"])


if __name__ == "__main__":
    unittest.main()
