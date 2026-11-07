import json
from decimal import Decimal

from foodmix.units import convert_quantity, load_rules


def test_same_unit_is_not_rewritten():
    outcome = convert_quantity("小葱", Decimal("2"), "斤", "斤", load_rules())
    assert outcome.ok
    assert outcome.status == "same"
    assert outcome.qty == Decimal("2")


def test_egg_board_converts_only_for_eggs():
    rules = load_rules()
    egg = convert_quantity("鸡蛋", Decimal("2"), "板", "个", rules)
    assert egg.ok
    assert egg.status == "converted"
    assert egg.qty == Decimal("60")
    bag = convert_quantity("煲汤袋", Decimal("1"), "板", "个", rules)
    assert not bag.ok
    assert bag.status == "mismatch"
    assert bag.qty is None
    scallion = convert_quantity("香葱", Decimal("1"), "板", "个", rules)
    assert not scallion.ok


def test_matching_catalog_unit_does_not_apply_a_rule():
    # 物料表本身就是「板」时，不能再按「板→个」乘一遍。
    outcome = convert_quantity("鸡蛋", Decimal("2"), "板", "板", load_rules())
    assert outcome.status == "same"
    assert outcome.qty == Decimal("2")


def test_spelling_fold_is_not_a_silent_kilo_conversion():
    rules = load_rules()
    folded = convert_quantity("小葱", Decimal("1"), "個", "个", rules)
    assert folded.status == "same"
    kilos = convert_quantity("五花肉", Decimal("1"), "公斤", "斤", rules)
    assert not kilos.ok
    assert kilos.status == "mismatch"


def test_rule_without_names_is_ignored(tmp_path):
    path = tmp_path / "units.json"
    path.write_text(json.dumps({
        "rules": [
            {"applies_to": [], "from_unit": "板", "to_unit": "个", "factor": 30},
            {"applies_to": ["鸡蛋"], "from_unit": "板", "to_unit": "个", "factor": 30},
        ]
    }, ensure_ascii=False), encoding="utf-8")
    rules = load_rules(path)
    assert len(rules) == 1
    assert not convert_quantity("香葱", Decimal("1"), "板", "个", rules).ok
