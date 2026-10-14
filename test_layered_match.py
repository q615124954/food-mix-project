# -*- coding: utf-8 -*-
"""分层品名匹配：确认回流、海天/老抽不得归生抽、金标 8/12 归属、加载失败要看得见。

不硬编码三张样票的品名清单。物料标准名和代码从金标物料表读出来。
OCR 原文用 samples/gold/ocr，不调在线 OCR，也不在默认路径调大模型。
"""
import json
import os
import sys
import tempfile
import unittest

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(BASE, "webapp"))
import app as webapp
import name_match
import pipeline_min as core

GOLD = os.path.join(BASE, "samples", "gold")


def _rows(day):
    with open(os.path.join(GOLD, "ocr", f"receipt_2026-08-{day}.ocr.txt"), encoding="utf-8") as f:
        return core.parse_table(f.read())


class _Isolated(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self._variants_file = name_match.VARIANTS_FILE
        self._variants = name_match._VARIANTS
        self._load_error = name_match._LOAD_ERROR
        self._paths = dict(webapp._paths)
        self._state = webapp._state
        self._name_dict = webapp._name_dict
        self._code_map = webapp._code_map
        self._save_state = webapp.save_state
        self._save_code_map = webapp.save_code_map
        self._dict_file = webapp.NAME_DICT_FILE
        self._code_file = webapp.CODE_MAP_FILE
        self._chat = core.call_chat
        name_match.VARIANTS_FILE = os.path.join(self.tmp.name, "name_variants.json")
        name_match._VARIANTS = None
        name_match._LOAD_ERROR = ""
        webapp.NAME_DICT_FILE = os.path.join(self.tmp.name, "product_dict.json")
        webapp.CODE_MAP_FILE = os.path.join(self.tmp.name, "code_map.json")
        webapp._name_dict = {"蔬菜": {}, "肉类": {}}
        webapp._code_map = {"肉类": {}, "蔬菜": {}}
        webapp._state = {"invoices": []}
        webapp.save_state = lambda: None
        out = os.path.join(self.tmp.name, "out")
        os.makedirs(out)
        webapp._paths.update(
            output_dir=out,
            material_file=os.path.join(GOLD, "material_gold.xlsx"),
        )
        webapp.load_legal_names()
        self.assertFalse(webapp._catalog_error, webapp._catalog_error)
        self.assertTrue(webapp._legal_codes.get("蒜头"))
        self.assertTrue(webapp._legal_codes.get("去皮蒜米"))

    def tearDown(self):
        core.call_chat = self._chat
        webapp.save_state = self._save_state
        webapp.save_code_map = self._save_code_map
        webapp.NAME_DICT_FILE = self._dict_file
        webapp.CODE_MAP_FILE = self._code_file
        webapp._name_dict = self._name_dict
        webapp._code_map = self._code_map
        webapp._state = self._state
        webapp._paths.clear()
        webapp._paths.update(self._paths)
        name_match.VARIANTS_FILE = self._variants_file
        name_match._VARIANTS = self._variants
        name_match._LOAD_ERROR = self._load_error
        webapp.load_legal_names()


class AliasWriteback(_Isolated):
    def test_confirm_once_then_same_ocr_name_autohits(self):
        """复核页确认一次 OCR 名 → 标准名/物料代码，下次同一 OCR 名走别名自动命中。"""
        ocr_name = "蒜脑袋"
        std = "蒜头"
        code = webapp._legal_codes[std]
        before = name_match.layered_match(ocr_name, webapp._legal_name_list, codes=webapp._legal_codes)
        self.assertNotEqual(before["action"], "auto")
        self.assertNotEqual(before["name"], std)

        webapp._state = {"invoices": [{
            "filename": "learn.jpg", "status": "pending", "ocr_done": True, "items": [],
        }]}
        items = [{"品名": std, "原品名": ocr_name, "品名编辑": True,
                  "数量": 1, "单价": 6, "金额": 6}]
        payload = webapp.ConfirmPayload(
            date="2026-10-14", category="蔬菜", items=items, paid_total=6)
        self.assertEqual(webapp.confirm("learn.jpg", payload), {"ok": True})

        # 下一轮进程：只从磁盘把别名读回来，纠错字典清掉，证明命中来自别名表而不是当次内存
        webapp._name_dict = {"蔬菜": {}, "肉类": {}}
        name_match._VARIANTS = None
        name_match._LOAD_ERROR = ""
        name_match.load_variants()
        self.assertEqual(name_match.alias_code("蔬菜", ocr_name), code)

        again = name_match.layered_match(
            ocr_name, webapp._legal_name_list, category="蔬菜", codes=webapp._legal_codes)
        self.assertEqual(again["action"], "auto")
        self.assertEqual(again["name"], std)
        self.assertGreaterEqual(again["confidence"], name_match.AUTO_CONF)
        self.assertIn(again["source"], ("variant", "alias"))

        # 标准名改了字、物料代码还在：仍按代码命中，不靠三张样票的品名硬编码
        renamed = std + "（规格调整）"
        hit = name_match.layered_match(
            ocr_name, [renamed, "香菜"], category="蔬菜", codes={renamed: code})
        self.assertEqual((hit["action"], hit["name"], hit["source"]), ("auto", renamed, "alias"))

        fresh = {"品名": ocr_name, "类别": "蔬菜", "单位": "斤", "数量": 1, "金额": 6}
        webapp.correct_item_names([fresh])
        self.assertEqual(fresh["品名"], std)
        self.assertIn(fresh["纠正来源"], ("variant", "alias", "dict"))


class DarkSoyBan(_Isolated):
    def test_haitian_laochou_never_auto_maps_to_shengchou(self):
        raws = ["消天老抽酱油", "海天老抽", "老抽酱油", "消天酱油"]
        for raw in raws:
            best, conf, src, _top = name_match.match(raw, webapp._legal_name_list, category="蔬菜")
            adopted, _prompt = name_match.correction_action(raw, best, conf, src)
            decision = name_match.layered_match(raw, webapp._legal_name_list, category="蔬菜")
            self.assertNotEqual(decision["action"], "auto", raw)
            self.assertNotIn("生抽", adopted or "")
            self.assertNotEqual(decision.get("suggestion"), "生抽", raw)
            self.assertNotEqual(decision.get("suggestion"), "鲜味生抽 1.9升", raw)
            forced = name_match.correction_action(raw, "生抽", 0.99, "fuzzy")
            self.assertEqual(forced, (raw, True), raw)
            forced_decision = name_match.decide(
                raw, "生抽", 0.99, "fuzzy", [("生抽", 0.99), ("鲜味生抽 1.9升", 0.96)])
            self.assertEqual(forced_decision["action"], "keep", raw)
            self.assertEqual(forced_decision["name"], raw, raw)
            self.assertEqual(forced_decision["suggestion"], "", raw)

        # 别名表里即使已经写了错误配对，也不自动采纳
        name_match._VARIANTS = {"蔬菜": {"生抽": ["消天老抽酱油"]}, "肉类": {}}
        name_match._LOAD_ERROR = ""
        best, conf, src, _ = name_match.match("消天老抽酱油", ["生抽", "蒜头"], category="蔬菜")
        adopted, _prompt = name_match.correction_action("消天老抽酱油", best, conf, src)
        self.assertNotEqual(adopted, "生抽")
        self.assertFalse(name_match.add_variant("蔬菜", "生抽", "消天老抽酱油", code="ZC-FAKE"))
        self.assertEqual(name_match.alias_code("蔬菜", "消天老抽酱油"), "")

    def test_confirm_does_not_learn_dark_soy_as_light_soy(self):
        webapp._state = {"invoices": [{
            "filename": "soy.jpg", "status": "pending", "ocr_done": True, "items": [],
        }]}
        items = [{"品名": "生抽", "原品名": "消天老抽酱油", "品名编辑": True,
                  "数量": 1, "单价": 13, "金额": 13}]
        payload = webapp.ConfirmPayload(
            date="2026-10-14", category="蔬菜", items=items, paid_total=13)
        self.assertEqual(webapp.confirm("soy.jpg", payload), {"ok": True})
        self.assertNotEqual(
            (webapp._name_dict.get("蔬菜") or {}).get("消天老抽酱油"), "生抽")
        name_match._VARIANTS = None
        name_match._LOAD_ERROR = ""
        name_match.load_variants()
        self.assertEqual(name_match.alias_code("蔬菜", "消天老抽酱油"), "")
        decision = name_match.layered_match(
            "消天老抽酱油", webapp._legal_name_list, codes=webapp._legal_codes)
        self.assertNotEqual(decision["action"], "auto")
        self.assertNotEqual(decision["name"], "生抽")


class GoldAugust12(_Isolated):
    def test_only_soy_sauce_row_is_staff_garlic_lands_as_material(self):
        """8/12：大蒜头/蒜半进物料；老抽一类拿不准才是员工餐。默认路径不调大模型。"""
        def no_llm(*_a, **_k):
            raise AssertionError("默认路径不应调用大模型")
        core.call_chat = no_llm
        name_match._VARIANTS = json.loads(json.dumps(name_match.DEFAULT_VARIANTS))
        name_match._LOAD_ERROR = ""

        items, _ = core.derive_price(core.parse_line_items(_rows("12")))
        webapp.correct_item_names(items)
        webapp.compute_suspects("2026-08-12", items)
        by = {it["原品名"]: it for it in items}

        garlic = by["大蒜头"]
        self.assertEqual(garlic["品名"], "蒜头")
        self.assertEqual(garlic["归属"], webapp.MATERIAL)
        self.assertFalse(garlic["归属待确认"])

        clove = by["蒜半"]
        self.assertEqual(clove["品名"], "去皮蒜米")
        self.assertEqual(clove["归属"], webapp.MATERIAL)

        soy = by["消天老抽酱油"]
        self.assertEqual(soy["品名"], "消天老抽酱油")
        self.assertNotIn("生抽", soy["品名"])
        self.assertNotEqual(soy.get("建议品名"), "生抽")
        self.assertEqual(soy["归属"], webapp.STAFF)
        self.assertTrue(soy["归属待确认"])

        staff = [it["原品名"] for it in items if it["归属"] == webapp.STAFF]
        self.assertEqual(staff, ["消天老抽酱油"])


class LoadFailure(_Isolated):
    def test_corrupt_variant_store_is_visible_and_not_mass_staff(self):
        with open(name_match.VARIANTS_FILE, "w", encoding="utf-8") as f:
            f.write("{")
        name_match._VARIANTS = None
        name_match._LOAD_ERROR = ""
        name_match.load_variants()
        with open(name_match.VARIANTS_FILE, encoding="utf-8") as f:
            self.assertEqual(f.read(), "{")  # 坏文件不能被空表覆盖
        err = name_match.variants_load_error()
        self.assertIn("品名变体表加载失败", err)
        self.assertEqual(webapp.classify_item("大蒜头", "蔬菜")[0], webapp.MATERIAL)
        self.assertEqual(webapp.classify_item("没见过的菜", "蔬菜")[0], webapp.MATERIAL)
        self.assertEqual(webapp.classify_item("海底捞", "蔬菜")[0], webapp.STAFF)
        item = {"品名": "大蒜头", "类别": "蔬菜", "数量": 1, "金额": 6}
        webapp.correct_item_names([item])
        webapp.compute_suspects("2026-08-12", [item])
        self.assertEqual(item["品名"], "大蒜头")
        self.assertEqual(item["归属"], webapp.MATERIAL)
        state = webapp.get_state()
        self.assertTrue(any("品名变体表加载失败" in s for s in state["catalog_errors"]))

    def test_missing_or_broken_material_table_is_visible(self):
        missing = os.path.join(self.tmp.name, "missing.xlsx")
        webapp._paths["material_file"] = missing
        webapp.load_legal_names()
        self.assertIn("物料明细表加载失败", webapp._catalog_error)
        self.assertIn("文件不存在", webapp._catalog_error)
        self.assertEqual(webapp.classify_item("香葱", "蔬菜")[0], webapp.MATERIAL)
        self.assertEqual(webapp.classify_item("没见过的菜", "蔬菜")[0], webapp.MATERIAL)

        broken = os.path.join(self.tmp.name, "broken.xlsx")
        with open(broken, "wb") as f:
            f.write(b"this is not a workbook")
        webapp._paths["material_file"] = broken
        webapp.load_legal_names()
        self.assertIn("物料明细表加载失败", webapp._catalog_error)
        self.assertEqual(webapp._legal_name_list, [])
        self.assertEqual(webapp.classify_item("随便一个品名", "蔬菜")[0], webapp.MATERIAL)
        state = webapp.get_state()
        self.assertTrue(any("物料明细表加载失败" in s for s in state["catalog_errors"]))

        page = os.path.join(BASE, "webapp", "static", "index.html")
        with open(page, encoding="utf-8") as f:
            html = f.read()
        self.assertIn('id="catalog-alert"', html)
        self.assertIn("catalog_errors", html)


class LlmOnlyWhenTied(_Isolated):
    def test_single_mid_score_suggests_without_llm(self):
        name_match._VARIANTS = {"蔬菜": {}, "肉类": {}}
        name_match._LOAD_ERROR = ""
        calls = []

        def boom(*_a, **_k):
            calls.append(1)
            raise AssertionError("单个中分不应调用大模型")

        core.call_chat = boom
        _best, _conf, _src, top = name_match.match("大蒜头", webapp._legal_name_list)
        self.assertEqual(len(name_match.similar_mid_candidates(top, "大蒜头")), 1)
        corrected, src, suggestion, action = webapp.apply_veg_name_correction("大蒜头", unit="斤")
        self.assertEqual(calls, [])
        self.assertEqual(corrected, "大蒜头")
        self.assertEqual((src, action, suggestion), ("suggest", "suggest", "蒜头"))

    def test_two_or_three_close_mids_may_call_llm(self):
        name_match._VARIANTS = {"蔬菜": {}, "肉类": {}}
        name_match._LOAD_ERROR = ""
        _best, _conf, _src, top = name_match.match("小葱", webapp._legal_name_list)
        close = name_match.similar_mid_candidates(top, "小葱")
        self.assertIn(len(close), (2, 3), close)

        calls = []

        def fake_chat(messages, model=None):
            calls.append(messages)
            return '{"目标品名":"香葱","置信度":0.7,"理由":"测试"}'

        core.call_chat = fake_chat
        corrected, src, _suggestion, action = webapp.apply_veg_name_correction("小葱")
        self.assertEqual(len(calls), 1)
        self.assertEqual((corrected, src, action), ("香葱", "llm", "llm"))

        def down(*_a, **_k):
            raise RuntimeError("llm down")

        core.call_chat = down
        corrected, src, suggestion, action = webapp.apply_veg_name_correction("小葱")
        self.assertEqual(corrected, "小葱")
        self.assertEqual(action, "suggest")
        self.assertTrue(suggestion)
        self.assertNotEqual(src, "llm")


if __name__ == "__main__":
    unittest.main()
