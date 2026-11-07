from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path

from foodmix.numbers import norm_name, parse_decimal
from foodmix.settings import CONFIG_DIR, read_json

# 只折叠写法差异。公斤和斤不是同一个单位，没有规则就不能换。
UNIT_FOLD = {
    "個": "个",
    "隻": "只",
    "KG": "kg",
    "Kg": "kg",
    "ｋｇ": "kg",
}


def fold_unit(value) -> str:
    text = "".join(str(value or "").split())
    return UNIT_FOLD.get(text, text)


@dataclass(frozen=True)
class UnitRule:
    applies_to: frozenset[str]
    from_unit: str
    to_unit: str
    factor: Decimal
    note: str


@dataclass(frozen=True)
class UnitOutcome:
    ok: bool
    status: str
    qty: Decimal | None
    unit: str
    reason: str


def load_rules(path: Path | None = None) -> list[UnitRule]:
    raw = read_json(path or (CONFIG_DIR / "unit_conversions.json"), {"rules": []})
    rules: list[UnitRule] = []
    for item in raw.get("rules") or []:
        names = frozenset(norm_name(name) for name in (item.get("applies_to") or []) if norm_name(name))
        # 空的 applies_to 会把「板」误套到所有货，直接忽略。
        if not names:
            continue
        factor = parse_decimal(item.get("factor"))
        source = fold_unit(item.get("from_unit"))
        target = fold_unit(item.get("to_unit"))
        if factor is None or factor <= 0 or not source or not target or source == target:
            continue
        rules.append(UnitRule(names, source, target, factor, str(item.get("note") or "")))
    return rules


def convert_quantity(name: str, qty: Decimal | None, bill_unit: str, catalog_unit: str, rules: list[UnitRule]) -> UnitOutcome:
    """换到物料表上的单位。对不上且没有这条品名的规则时，不换、不写。"""
    target = fold_unit(catalog_unit)
    source = fold_unit(bill_unit)
    if qty is None:
        return UnitOutcome(False, "missing", None, target, "缺少数量，未写入")
    if not target:
        return UnitOutcome(False, "missing", None, "", "物料表该品名没有单位，拒绝落表")
    if not source:
        return UnitOutcome(False, "missing", None, str(catalog_unit).strip(), "单据行没有单位，拒绝落表")
    if source == target:
        return UnitOutcome(True, "same", qty, str(catalog_unit).strip(), "单位与物料表一致")
    key = norm_name(name)
    for rule in rules:
        if key in rule.applies_to and rule.from_unit == source and rule.to_unit == target:
            converted = qty * rule.factor
            return UnitOutcome(
                True,
                "converted",
                converted,
                str(catalog_unit).strip(),
                f"{qty}{source} × {rule.factor} = {converted}{target}",
            )
    return UnitOutcome(
        False,
        "mismatch",
        None,
        target,
        f"单位「{bill_unit}」与物料表「{catalog_unit}」不一致，且没有适用于「{name}」的换算规则，未写入",
    )
