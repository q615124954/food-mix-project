# -*- coding: utf-8 -*-
"""高分弱相似不得把「香菜」静默纠成「香葱」，应保留原名并提示人工。"""
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(ROOT, "webapp"))
sys.path.insert(0, ROOT)

import name_match
import app as webapp


class NameMatchConfusableTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self._variants_file = name_match.VARIANTS_FILE
        self._variants = name_match._VARIANTS
        name_match.VARIANTS_FILE = os.path.join(self.tmp.name, "name_variants.json")
        name_match._VARIANTS = None
        self._legal_names = webapp._legal_names
        self._legal_list = list(webapp._legal_name_list)
        self._legal_units = dict(webapp._legal_units)

    def tearDown(self):
        name_match.VARIANTS_FILE = self._variants_file
        name_match._VARIANTS = self._variants
        webapp._legal_names = self._legal_names
        webapp._legal_name_list = self._legal_list
        webapp._legal_units = self._legal_units

    def _use_catalog(self, names, unit="斤"):
        webapp._legal_name_list = list(names)
        webapp._legal_names = {webapp.normalize_pin(n) for n in names}
        webapp._legal_units = {n: unit for n in names}

    def test_xiangcai_not_silently_rewritten_to_xiangcong(self):
        """单位把首字母分加成 0.9 时，香菜不得静默变成香葱，并带人工提示。"""
        candidates = ["香葱", "小葱", "生菜"]
        units = {c: "斤" for c in candidates}
        best, conf, source, _top = name_match.match(
            "香菜", candidates, category="蔬菜", units=units, unit="斤")
        adopted, prompt = name_match.correction_action("香菜", best, conf, source)

        silent = adopted == "香葱" and not prompt
        self.assertFalse(silent)
        self.assertLess(conf, name_match.AUTO_CONF)
        self.assertEqual(conf, name_match.WEAK_SCORE_CAP)
        self.assertEqual(source, name_match.REVIEW_SOURCE)
        self.assertTrue(prompt)
        self.assertEqual(adopted, "香菜")

        # 候选里没有香菜，但保留下来的原名仍视为合法，存疑只能来自提示标志
        self._use_catalog(candidates)
        webapp._legal_names.add(webapp.normalize_pin("香菜"))
        corrected, src, _suggestion, _action = webapp.apply_veg_name_correction("香菜", unit="斤")
        self.assertEqual(corrected, "香菜")
        self.assertNotEqual(corrected, "香葱")
        self.assertEqual(src, name_match.REVIEW_SOURCE)
        item = {"品名": corrected, "纠正来源": src, "勾稽": ""}
        suspects = webapp.compute_suspects("2026-10-13", [item])
        self.assertTrue(item["品名存疑"])
        self.assertTrue(any("品名存疑" in s for s in suspects))

    def test_same_class_initials_not_silently_rewritten(self):
        """同类：白菜 / 白茶同首字母，单位加成后同样不得静默替换。"""
        candidates = ["白茶", "菠菜"]
        units = {c: "斤" for c in candidates}
        best, conf, source, _top = name_match.match(
            "白菜", candidates, category="蔬菜", units=units, unit="斤")
        adopted, prompt = name_match.correction_action("白菜", best, conf, source)
        self.assertFalse(adopted == "白茶" and not prompt)
        self.assertEqual(adopted, "白菜")
        self.assertTrue(prompt)
        self.assertEqual(source, name_match.REVIEW_SOURCE)
        self.assertLess(conf, name_match.AUTO_CONF)

    def test_homophone_still_auto_corrects(self):
        """全拼相同的错字（芥兰→芥蓝）仍自动纠正，不标成易混存疑。"""
        best, conf, source, _top = name_match.match(
            "芥兰", ["芥蓝"], category="蔬菜", units={"芥蓝": "斤"}, unit="斤")
        adopted, prompt = name_match.correction_action("芥兰", best, conf, source)
        self.assertEqual(adopted, "芥蓝")
        self.assertFalse(prompt)
        self.assertGreaterEqual(conf, name_match.AUTO_CONF)
        self.assertNotEqual(source, name_match.REVIEW_SOURCE)

        self._use_catalog(["芥蓝"])
        corrected, src, _suggestion, _action = webapp.apply_veg_name_correction("芥兰", unit="斤")
        self.assertEqual(corrected, "芥蓝")
        self.assertEqual(src, "fuzzy")
        item = {"品名": corrected, "纠正来源": src, "勾稽": ""}
        webapp.compute_suspects("2026-10-13", [item])
        self.assertFalse(item["品名存疑"])

    def test_exact_xiangcai_stays_when_listed(self):
        best, conf, source, _top = name_match.match(
            "香菜", ["香菜", "香葱"], category="蔬菜",
            units={"香菜": "斤", "香葱": "斤"}, unit="斤")
        adopted, prompt = name_match.correction_action("香菜", best, conf, source)
        self.assertEqual(best, "香菜")
        self.assertEqual(adopted, "香菜")
        self.assertFalse(prompt)
        self.assertEqual(conf, 1.0)
        self.assertEqual(source, "exact")

    def test_variant_xiangrui_maps_to_xiangcong(self):
        # 金标核对：8/11、8/12 的「香蕊」和 8/13 的「香芯」单价都是 4 元/斤，对应香葱（香菜约 7.5 元/斤）
        best, conf, source, _top = name_match.match(
            "香蕊", ["香菜", "香葱"], category="蔬菜")
        adopted, prompt = name_match.correction_action("香蕊", best, conf, source)
        self.assertEqual(best, "香葱")
        self.assertEqual(adopted, "香葱")
        self.assertFalse(prompt)
        self.assertGreaterEqual(conf, name_match.AUTO_CONF)
        self.assertEqual(source, "variant")


if __name__ == "__main__":
    unittest.main()
