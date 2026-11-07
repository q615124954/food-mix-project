from decimal import Decimal

from foodmix.catalog import load_catalog
from foodmix.plan import BillLine, build_plan
from foodmix.settings import load_sheet_layout
from foodmix.units import load_rules as load_unit_rules
from tests.conftest import SAMPLE_DATE, make_material


def line(invoice, name, category, qty="1", unit="斤", amount="1"):
    return BillLine(invoice, SAMPLE_DATE, name, name, Decimal(qty), unit, Decimal(amount), category)


def test_rows_go_to_their_own_sheets():
    # build_plan 没有整单类别参数，只看每一行。
    import tempfile
    from pathlib import Path
    folder = Path(tempfile.mkdtemp())
    path = make_material(folder / "m.xlsx")
    catalog = load_catalog(path, load_sheet_layout())
    writes, skipped = build_plan(
        [
            line("a", "小葱", "蔬菜", "2", "斤", "5"),
            line("a", "五花肉", "肉类", "1", "斤", "20"),
            line("a", "五花肉", "蔬菜", "9", "斤", "9"),
            BillLine("a", SAMPLE_DATE, "可乐", "可乐", Decimal("1"), "瓶", Decimal("3"), "蔬菜"),
            BillLine("a", SAMPLE_DATE, "未知", "未知", Decimal("1"), "斤", Decimal("1"), ""),
        ],
        catalog,
        load_unit_rules(),
        load_sheet_layout(),
    )
    landed = {(item.sheet, item.name, item.qty) for item in writes}
    assert ("日物料采购", "小葱", Decimal("2")) in landed
    assert ("肉类采购", "五花肉", Decimal("1")) in landed
    assert not any(item.name == "五花肉" and item.sheet == "日物料采购" for item in writes)
    reasons = " ".join(item.reason for item in skipped)
    assert "找不到" in reasons
    assert "不是蔬菜或肉类" in reasons


def test_shipped_layout_does_not_target_staff_meal():
    layout = load_sheet_layout()
    assert layout.category_sheets == {"蔬菜": "日物料采购", "肉类": "肉类采购"}
    assert "员工餐" in layout.never_write_sheets
    assert "员工餐" not in layout.writable_sheets
