from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from foodmix.catalog import Catalog
from foodmix.numbers import norm_name
from foodmix.settings import CONFIG_DIR, SheetLayout, read_json


@dataclass(frozen=True)
class Alias:
    alias: str
    name: str
    category: str


@dataclass(frozen=True)
class MapHit:
    name: str
    status: str
    reason: str
    suggestion: str = ""
    confidence: float | None = None


def load_aliases(paths: list[Path] | None = None) -> list[Alias]:
    if paths is None:
        paths = [CONFIG_DIR / "aliases.json"]
    found: list[Alias] = []
    for path in paths:
        if not path.is_file():
            continue
        raw = read_json(path, {"aliases": []})
        for item in raw.get("aliases") or []:
            alias = norm_name(item.get("alias"))
            name = str(item.get("name") or "").strip()
            if not alias or not name:
                continue
            found.append(Alias(alias, name, str(item.get("category") or "").strip()))
    return found


def lookup_alias(raw_name: str, category: str, aliases: list[Alias]) -> str | None:
    key = norm_name(raw_name)
    generic = None
    for alias in aliases:
        if alias.alias != key:
            continue
        if alias.category == category:
            return alias.name
        if not alias.category and generic is None:
            generic = alias.name
    return generic


def _sheet_for(category: str, layout: SheetLayout) -> str | None:
    if category not in {"蔬菜", "肉类"}:
        return None
    sheet = layout.category_sheets.get(category)
    if not sheet or sheet in layout.never_write_sheets:
        return None
    return sheet


def map_name(
    raw_name: str,
    category: str,
    catalog: Catalog,
    aliases: list[Alias],
    layout: SheetLayout,
    *,
    suggestion: dict | None = None,
    threshold: float = 0.8,
) -> MapHit:
    """别名、表内原名、模型建议，都只认这一行类别对应的那张表。"""
    sheet = _sheet_for(category, layout)
    if sheet is None:
        return MapHit(str(raw_name or "").strip(), "unmapped", "行类别不是蔬菜或肉类，不能归表")
    kept = str(raw_name or "").strip()
    alias_target = lookup_alias(raw_name, category, aliases)
    if alias_target:
        item = catalog.find(sheet, alias_target)
        if item:
            return MapHit(item.name, "alias", f"别名「{kept}」→「{item.name}」")
        return MapHit(kept, "unmapped", f"别名指向「{alias_target}」，但「{sheet}」中没有该品名")
    item = catalog.find(sheet, raw_name)
    if item:
        return MapHit(item.name, "exact", "与物料表品名一致")
    if suggestion and str(suggestion.get("name") or "").strip():
        target = str(suggestion["name"]).strip()
        try:
            confidence = float(suggestion.get("confidence") or 0)
        except (TypeError, ValueError):
            confidence = 0.0
        hit = catalog.find(sheet, target)
        if hit and confidence >= threshold:
            return MapHit(hit.name, "model", f"模型将「{kept}」映射为「{hit.name}」", suggestion=hit.name, confidence=confidence)
        if hit:
            return MapHit(
                kept,
                "unmapped",
                f"模型建议「{hit.name}」，置信度 {confidence:g} 低于 {threshold:g}，未自动改名",
                suggestion=hit.name,
                confidence=confidence,
            )
        return MapHit(kept, "unmapped", f"模型建议「{target}」不在「{sheet}」，已拒绝")
    return MapHit(kept, "unmapped", f"「{kept or '（空品名）'}」在「{sheet}」中找不到，跳过不写")


def resolve_review(
    raw_name: str,
    chosen_name: str,
    category: str,
    catalog: Catalog | None,
    aliases: list[Alias],
    layout: SheetLayout,
    *,
    previous: dict | None = None,
    threshold: float = 0.8,
) -> MapHit:
    chosen = str(chosen_name or "").strip()
    if catalog is None:
        if not chosen:
            return MapHit("", "unmapped", "物料表未加载，且没有填写品名")
        return MapHit(chosen, "manual", "物料表未加载，尚未核对是否能落表")
    previous = previous or {}
    if (
        previous.get("map_status") == "model"
        and previous.get("name") == chosen
        and previous.get("category") == category
        and chosen
    ):
        sheet = _sheet_for(category, layout)
        item = catalog.find(sheet, chosen) if sheet else None
        if item:
            confidence = previous.get("confidence")
            return MapHit(item.name, "model", str(previous.get("map_reason") or "沿用模型映射"), suggestion=item.name, confidence=confidence)
    auto = map_name(raw_name, category, catalog, aliases, layout, suggestion=None, threshold=threshold)
    if not chosen:
        if auto.status == "unmapped":
            return auto
        return MapHit("", "unmapped", "人工清空品名，不落表")
    if auto.status in {"alias", "exact"} and auto.name == chosen:
        return auto
    sheet = _sheet_for(category, layout)
    if sheet is None:
        return MapHit(chosen, "unmapped", "行类别不是蔬菜或肉类，不能归表")
    item = catalog.find(sheet, chosen)
    if not item:
        return MapHit(chosen, "unmapped", f"「{chosen}」在「{sheet}」中找不到，跳过不写")
    return MapHit(item.name, "manual", f"人工指定「{item.name}」")
