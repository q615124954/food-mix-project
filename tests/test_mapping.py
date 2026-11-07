from foodmix.catalog import Catalog, Item, ParsedSheet
from foodmix.mapping import Alias, map_name
from foodmix.numbers import norm_name
from foodmix.settings import load_sheet_layout


def catalog_with(*rows):
    items = []
    sheets = {}
    for sheet, row, code, name, unit in rows:
        item = Item(sheet, row, code, name, unit, norm_name(name))
        items.append(item)
        sheets.setdefault(sheet, ParsedSheet(sheet, 1, {}, []))
        sheets[sheet].items.append(item)
    return Catalog(items, sheets, [], "memory")


def test_alias_maps_marketing_name_to_catalog_name():
    catalog = catalog_with(("日物料采购", 3, "V", "小葱", "斤"), ("日物料采购", 4, "V2", "香葱", "斤"))
    aliases = [Alias(norm_name("无敌美味小小葱"), "小葱", "蔬菜")]
    hit = map_name("无敌美味小小葱", "蔬菜", catalog, aliases, load_sheet_layout())
    assert hit.status == "alias"
    assert hit.name == "小葱"


def test_alias_to_a_missing_name_does_not_guess_a_neighbor():
    catalog = catalog_with(("日物料采购", 4, "V2", "香葱", "斤"))
    aliases = [Alias(norm_name("无敌美味小小葱"), "小葱", "蔬菜")]
    hit = map_name("无敌美味小小葱", "蔬菜", catalog, aliases, load_sheet_layout())
    assert hit.status == "unmapped"
    assert hit.name != "香葱"
    assert "没有该品名" in hit.reason


def test_cola_is_not_in_the_catalog():
    catalog = catalog_with(("日物料采购", 3, "V", "小葱", "斤"))
    hit = map_name("可乐", "蔬菜", catalog, [], load_sheet_layout())
    assert hit.status == "unmapped"
    assert "找不到" in hit.reason
    assert hit.name == "可乐"


def test_exact_name_and_model_suggestion_stay_on_the_line_sheet():
    catalog = catalog_with(
        ("日物料采购", 3, "V", "小葱", "斤"),
        ("肉类采购", 4, "M", "五花肉", "斤"),
    )
    layout = load_sheet_layout()
    exact = map_name("小葱", "蔬菜", catalog, [], layout)
    assert exact.status == "exact"
    assert exact.name == "小葱"
    accepted = map_name(
        "鲜五花", "肉类", catalog, [], layout,
        suggestion={"name": "五花肉", "confidence": 0.91},
    )
    assert accepted.status == "model"
    assert accepted.name == "五花肉"
    rejected = map_name(
        "可乐", "蔬菜", catalog, [], layout,
        suggestion={"name": "雪碧", "confidence": 0.99},
    )
    assert rejected.status == "unmapped"
    assert rejected.name == "可乐"
    assert "已拒绝" in rejected.reason
    low = map_name(
        "葱花", "蔬菜", catalog, [], layout,
        suggestion={"name": "小葱", "confidence": 0.4},
    )
    assert low.status == "unmapped"
    assert low.name == "葱花"
    assert low.suggestion == "小葱"


def test_vegetable_alias_does_not_cross_into_meat():
    catalog = catalog_with(("日物料采购", 3, "V", "小葱", "斤"), ("肉类采购", 4, "M", "五花肉", "斤"))
    aliases = [Alias(norm_name("无敌美味小小葱"), "小葱", "蔬菜")]
    hit = map_name("无敌美味小小葱", "肉类", catalog, aliases, load_sheet_layout())
    assert hit.status == "unmapped"
    assert hit.name != "小葱"
