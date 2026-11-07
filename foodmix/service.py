from __future__ import annotations

import threading
from datetime import datetime
from pathlib import Path

from foodmix.catalog import Catalog, load_catalog
from foodmix.errors import AppError
from foodmix.extract import parse_extraction
from foodmix.mapping import load_aliases, map_name, resolve_review
from foodmix.numbers import dump_decimal, norm_date, parse_decimal
from foodmix.plan import BillLine, build_plan
from foodmix.settings import (
    CONFIG_DIR,
    Paths,
    SheetLayout,
    load_model_config,
    load_paths,
    load_sheet_layout,
    local_aliases_path,
    prompt_text,
    save_paths,
    state_path,
)
from foodmix.store import read_state, write_state
from foodmix.sync_excel import apply_plan
from foodmix.units import load_rules
from foodmix.vision import MappingClient, VisionClient

IMAGE_EXT = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}
MIME = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".webp": "image/webp",
    ".bmp": "image/bmp",
}


def _blank_line(row: dict) -> dict:
    return {
        "line_no": row["line_no"],
        "raw_name": row["raw_name"],
        "name": "",
        "qty": dump_decimal(row["qty"]),
        "unit": row["unit"],
        "amount": dump_decimal(row["amount"]),
        "category": row["category"],
        "map_status": "unmapped",
        "map_reason": "",
        "suggestion": "",
        "confidence": None,
        "issues": list(row.get("issues") or []),
    }


def _apply_hit(line: dict, hit) -> None:
    line["name"] = hit.name
    line["map_status"] = hit.status
    line["map_reason"] = hit.reason
    line["suggestion"] = hit.suggestion
    line["confidence"] = hit.confidence


class Workspace:
    def __init__(self):
        self._lock = threading.Lock()
        self._state = None
        self._state_path: Path | None = None
        self._state_mtime: float | None = None
        self._catalog: Catalog | None = None
        self._catalog_key: tuple | None = None

    def _bind_state(self) -> dict:
        path = state_path()
        mtime = path.stat().st_mtime if path.exists() else None
        if self._state is None or self._state_path != path or self._state_mtime != mtime:
            self._state = read_state(path)
            self._state_path = path
            self._state_mtime = mtime
        return self._state

    def _save_state(self, state: dict) -> None:
        path = state_path()
        write_state(path, state)
        self._state = state
        self._state_path = path
        self._state_mtime = path.stat().st_mtime

    def _layout(self) -> SheetLayout:
        return load_sheet_layout()

    def _rules(self):
        return load_rules()

    def _aliases(self):
        # 本地复核学到的别名盖过仓库里的示例。
        return load_aliases([local_aliases_path(), CONFIG_DIR / "aliases.json"])

    def _catalog_for(self, layout: SheetLayout) -> tuple[Catalog | None, str]:
        material = load_paths().material_file
        if not material:
            return None, "未配置物料明细表"
        path = Path(material)
        if not path.is_file():
            return None, f"物料明细表不存在：{material}"
        key = (str(path.resolve()), path.stat().st_mtime)
        if self._catalog is not None and self._catalog_key == key:
            return self._catalog, ""
        try:
            catalog = load_catalog(path, layout)
        except AppError as exc:
            self._catalog = None
            self._catalog_key = None
            return None, exc.message
        self._catalog = catalog
        self._catalog_key = key
        return catalog, ""

    def health(self) -> dict:
        config = load_model_config()
        return {
            "ok": True,
            "phase": 1,
            "vision_model": config.vision_model,
            "api_base": config.api_base,
            "has_api_key": config.has_key,
            "writes_staff_meal": False,
        }

    def get_settings(self) -> dict:
        paths = load_paths()
        return {"input_dir": paths.input_dir, "material_file": paths.material_file}

    def update_settings(self, input_dir: str, material_file: str) -> dict:
        save_paths(Paths(str(input_dir or "").strip(), str(material_file or "").strip()))
        self._catalog = None
        self._catalog_key = None
        return self.get_settings()

    def scan(self) -> dict:
        with self._lock:
            folder = load_paths().input_dir
            if not folder:
                raise AppError("请先填写电子单目录")
            root = Path(folder)
            if not root.is_dir():
                raise AppError(f"电子单目录不存在：{folder}")
            root_resolved = root.resolve()
            found = []
            for path in sorted(root.iterdir(), key=lambda item: item.name):
                if not path.is_file() or path.suffix.lower() not in IMAGE_EXT:
                    continue
                if "/" in path.name or "\\" in path.name or ".." in path.name:
                    continue
                resolved = path.resolve()
                if resolved.parent != root_resolved:
                    continue
                found.append(resolved)
            state = self._bind_state()
            previous = {invoice["filename"]: invoice for invoice in state["invoices"]}
            invoices = []
            for path in found:
                kept = previous.get(path.name)
                if kept:
                    kept["path"] = str(path)
                    invoices.append(kept)
                else:
                    invoices.append({
                        "id": path.name,
                        "filename": path.name,
                        "path": str(path),
                        "status": "pending",
                        "date": "",
                        "supplier": "",
                        "error": "",
                        "warning": "",
                        "confirmed_at": "",
                        "lines": [],
                    })
            state["invoices"] = invoices
            self._save_state(state)
            return self._list_locked(state)

    def list_invoices(self) -> dict:
        with self._lock:
            return self._list_locked(self._bind_state())

    def _list_locked(self, state: dict) -> dict:
        items = []
        for invoice in state["invoices"]:
            items.append({
                "id": invoice["id"],
                "filename": invoice["filename"],
                "status": invoice["status"],
                "date": invoice.get("date") or "",
                "supplier": invoice.get("supplier") or "",
                "line_count": len(invoice.get("lines") or []),
                "error": invoice.get("error") or "",
                "warning": invoice.get("warning") or "",
                "confirmed_at": invoice.get("confirmed_at") or "",
            })
        counts = {
            "total": len(items),
            "pending": sum(1 for item in items if item["status"] == "pending"),
            "recognized": sum(1 for item in items if item["status"] == "recognized"),
            "confirmed": sum(1 for item in items if item["status"] == "confirmed"),
            "error": sum(1 for item in items if item["status"] == "error"),
        }
        return {"invoices": items, "counts": counts}

    def get_invoice(self, invoice_id: str) -> dict:
        with self._lock:
            state = self._bind_state()
            return self._public(self._must(state, invoice_id))

    def image_file(self, invoice_id: str) -> Path:
        with self._lock:
            state = self._bind_state()
            invoice = self._must(state, invoice_id)
            path = Path(invoice.get("path") or "")
            root = load_paths().input_dir
            if not root or not path.is_file():
                raise AppError("图片不存在", status=404)
            root_resolved = Path(root).resolve()
            resolved = path.resolve()
            if resolved.parent != root_resolved:
                raise AppError("图片不在电子单目录里", status=404)
            return resolved

    def recognize(self, invoice_id: str, vision=None, suggestor=None) -> dict:
        with self._lock:
            state = self._bind_state()
            invoice = self._must(state, invoice_id)
            if invoice["status"] == "confirmed":
                raise AppError("已确认的单据已锁定。要重新识别，请先取消确认", status=409)
            path = Path(invoice.get("path") or "")
            if not path.is_file():
                raise AppError("图片文件不存在", status=404)
            config = load_model_config()
            client = vision or VisionClient(config)
            try:
                text = client.extract(path.read_bytes(), MIME.get(path.suffix.lower(), "image/jpeg"), prompt_text("extract_lines.md"))
                extracted = parse_extraction(text)
            except AppError as exc:
                invoice["status"] = "error"
                invoice["error"] = exc.message
                self._save_state(state)
                raise
            layout = self._layout()
            catalog, catalog_error = self._catalog_for(layout)
            aliases = self._aliases()
            lines = []
            queries = []
            for row in extracted["lines"]:
                line = _blank_line(row)
                if not row["category"]:
                    line["name"] = row["raw_name"]
                    line["map_reason"] = "这一行没有蔬菜/肉类，请人工选择。整单类别不会自动套用"
                elif catalog is None:
                    line["name"] = row["raw_name"]
                    line["map_reason"] = catalog_error or "物料表未加载"
                else:
                    hit = map_name(
                        row["raw_name"], row["category"], catalog, aliases, layout,
                        suggestion=None, threshold=config.map_confidence,
                    )
                    _apply_hit(line, hit)
                    if hit.status == "unmapped" and row["raw_name"]:
                        sheet = layout.category_sheets[row["category"]]
                        queries.append({
                            "key": str(row["line_no"]),
                            "raw_name": row["raw_name"],
                            "category": row["category"],
                            "candidates": [item.name for item in catalog.names_for(row["category"], layout)],
                            "sheet": sheet,
                        })
                lines.append(line)
            warning = ""
            if queries:
                try:
                    mapper = suggestor or MappingClient()
                    suggestions = mapper.suggest_batch(queries, prompt_text("map_names.md"), config)
                except AppError as exc:
                    suggestions = {}
                    warning = f"品名映射模型不可用，未自动改名：{exc.message}"
                by_no = {line["line_no"]: line for line in lines}
                for query in queries:
                    line = by_no.get(int(query["key"]))
                    if line is None or line["map_status"] != "unmapped":
                        continue
                    suggestion = suggestions.get(query["key"]) if suggestions else None
                    if not suggestion:
                        continue
                    hit = map_name(
                        line["raw_name"], line["category"], catalog, aliases, layout,
                        suggestion=suggestion, threshold=config.map_confidence,
                    )
                    _apply_hit(line, hit)
            invoice["lines"] = lines
            if extracted["date"]:
                invoice["date"] = extracted["date"]
            invoice["supplier"] = extracted["supplier"]
            invoice["status"] = "recognized"
            invoice["error"] = ""
            invoice["warning"] = warning
            invoice["confirmed_at"] = ""
            self._save_state(state)
            return self._public(invoice)

    def remap(self, invoice_id: str, suggestor=None) -> dict:
        with self._lock:
            state = self._bind_state()
            invoice = self._must(state, invoice_id)
            if invoice["status"] == "confirmed":
                raise AppError("已确认的单据已锁定，取消确认后才能重新匹配", status=409)
            if invoice["status"] == "pending":
                raise AppError("请先识别这张电子单")
            self._remap_lines(invoice, suggestor)
            invoice["status"] = "recognized"
            self._save_state(state)
            return self._public(invoice)

    def _remap_lines(self, invoice: dict, suggestor=None) -> None:
        layout = self._layout()
        catalog, catalog_error = self._catalog_for(layout)
        aliases = self._aliases()
        config = load_model_config()
        queries = []
        for line in invoice["lines"]:
            if line.get("map_status") == "manual" and catalog is not None:
                sheet = layout.category_sheets.get(line.get("category") or "")
                if sheet and catalog.find(sheet, line.get("name") or ""):
                    continue
            if catalog is None or line.get("category") not in {"蔬菜", "肉类"}:
                line["map_status"] = "unmapped"
                line["map_reason"] = catalog_error or "行类别不是蔬菜或肉类，不能归表"
                if not line.get("name"):
                    line["name"] = line.get("raw_name") or ""
                continue
            hit = map_name(
                line.get("raw_name") or "", line["category"], catalog, aliases, layout,
                suggestion=None, threshold=config.map_confidence,
            )
            _apply_hit(line, hit)
            if hit.status == "unmapped" and line.get("raw_name"):
                queries.append({
                    "key": str(line["line_no"]),
                    "raw_name": line["raw_name"],
                    "category": line["category"],
                    "candidates": [item.name for item in catalog.names_for(line["category"], layout)],
                })
        if not queries:
            invoice["warning"] = ""
            return
        try:
            mapper = suggestor or MappingClient()
            suggestions = mapper.suggest_batch(queries, prompt_text("map_names.md"), config)
            invoice["warning"] = ""
        except AppError as exc:
            suggestions = {}
            invoice["warning"] = f"品名映射模型不可用，未自动改名：{exc.message}"
        by_no = {line["line_no"]: line for line in invoice["lines"]}
        for query in queries:
            line = by_no.get(int(query["key"]))
            suggestion = suggestions.get(query["key"]) if suggestions else None
            if line is None or not suggestion or catalog is None:
                continue
            hit = map_name(
                line.get("raw_name") or "", line["category"], catalog, aliases, layout,
                suggestion=suggestion, threshold=config.map_confidence,
            )
            _apply_hit(line, hit)

    def save(self, invoice_id: str, date: str, lines: list[dict]) -> dict:
        with self._lock:
            state = self._bind_state()
            invoice = self._must(state, invoice_id)
            if invoice["status"] == "confirmed":
                raise AppError("已确认的单据已锁定，取消确认后才能修改", status=409)
            if invoice["status"] == "pending":
                raise AppError("请先识别这张电子单")
            self._replace_lines(invoice, date, lines)
            invoice["status"] = "recognized"
            self._save_state(state)
            return self._public(invoice)

    def confirm(self, invoice_id: str, date: str, lines: list[dict]) -> dict:
        """只锁定复核结果。不打开、不修改物料明细表。"""
        with self._lock:
            state = self._bind_state()
            invoice = self._must(state, invoice_id)
            if invoice["status"] in {"pending", "error"}:
                raise AppError("请先识别这张电子单")
            if invoice["status"] != "confirmed":
                if lines is None:
                    lines = [
                        {
                            "line_no": line["line_no"],
                            "name": line.get("name") or "",
                            "qty": line.get("qty"),
                            "unit": line.get("unit") or "",
                            "amount": line.get("amount"),
                            "category": line.get("category") or "",
                        }
                        for line in invoice.get("lines") or []
                    ]
                self._replace_lines(invoice, date if date is not None else invoice.get("date") or "", lines)
                self._learn_aliases(invoice)
            elif lines:
                raise AppError("已确认的单据已锁定，取消确认后才能修改", status=409)
            invoice["status"] = "confirmed"
            invoice["confirmed_at"] = invoice.get("confirmed_at") or datetime.now().isoformat(timespec="seconds")
            invoice["error"] = ""
            self._save_state(state)
            return self._public(invoice)

    def unconfirm(self, invoice_id: str) -> dict:
        with self._lock:
            state = self._bind_state()
            invoice = self._must(state, invoice_id)
            if invoice["status"] == "confirmed":
                invoice["status"] = "recognized"
                invoice["confirmed_at"] = ""
                self._save_state(state)
            return self._public(invoice)

    def sync(self) -> dict:
        with self._lock:
            state = self._bind_state()
            confirmed = [invoice for invoice in state["invoices"] if invoice.get("status") == "confirmed"]
            if not confirmed:
                raise AppError("没有已确认的单据，未写入物料明细")
            material = load_paths().material_file
            if not material:
                raise AppError("请先填写物料明细表路径")
            path = Path(material)
            layout = self._layout()
            catalog, catalog_error = self._catalog_for(layout)
            if catalog is None:
                raise AppError(catalog_error or "物料表无法读取")
            bill_lines = []
            for invoice in confirmed:
                for line in invoice.get("lines") or []:
                    bill_lines.append(BillLine(
                        invoice_id=invoice["id"],
                        date=norm_date(invoice.get("date")),
                        raw_name=str(line.get("raw_name") or ""),
                        name=str(line.get("name") or ""),
                        qty=parse_decimal(line.get("qty")),
                        unit=str(line.get("unit") or ""),
                        amount=parse_decimal(line.get("amount")),
                        category=str(line.get("category") or ""),
                    ))
            writes, skipped = build_plan(bill_lines, catalog, self._rules(), layout)
            ledger = state.get("ledger") or {}
            owned = ledger.get("cells") or []
            if str(ledger.get("material_file") or "") != str(path):
                owned = []
            try:
                owned_now, backup = apply_plan(path, writes, owned, layout)
            except PermissionError as exc:
                raise AppError("物料明细表被占用，请先关闭 WPS 或 Excel 再同步") from exc
            self._catalog = None
            self._catalog_key = None
            result = {
                "synced_at": datetime.now().isoformat(timespec="seconds"),
                "backup": backup,
                "material_file": str(path),
                "invoices": [invoice["id"] for invoice in confirmed],
                "written": [
                    {
                        "date": write.date,
                        "sheet": write.sheet,
                        "name": write.name,
                        "code": write.code,
                        "unit": write.unit,
                        "qty": dump_decimal(write.qty),
                        "amount": dump_decimal(write.amount),
                        "invoices": write.invoice_ids,
                    }
                    for write in writes
                ],
                "skipped": [
                    {
                        "invoice_id": skip.invoice_id,
                        "date": skip.date,
                        "raw_name": skip.raw_name,
                        "name": skip.name,
                        "category": skip.category,
                        "sheet": skip.sheet,
                        "reason": skip.reason,
                    }
                    for skip in skipped
                ],
                "staff_meal_written": False,
            }
            state["ledger"] = {"material_file": str(path), "cells": owned_now}
            state["last_sync"] = result
            self._save_state(state)
            return result

    def last_sync(self) -> dict:
        with self._lock:
            state = self._bind_state()
            return state.get("last_sync") or {"written": [], "skipped": [], "staff_meal_written": False}

    def catalog_public(self) -> dict:
        layout = self._layout()
        catalog, error = self._catalog_for(layout)
        sheets = {"蔬菜": [], "肉类": []}
        if catalog is not None:
            for category in sheets:
                sheets[category] = [
                    {"name": item.name, "unit": item.unit, "code": item.code}
                    for item in catalog.names_for(category, layout)
                ]
        return {"error": error, "sheets": sheets, "missing_sheets": catalog.missing_sheets if catalog else []}

    def _must(self, state: dict, invoice_id: str) -> dict:
        for invoice in state["invoices"]:
            if invoice["id"] == invoice_id:
                return invoice
        raise AppError("找不到这张电子单", status=404)

    def _replace_lines(self, invoice: dict, date: str, incoming: list[dict]) -> None:
        if date is not None:
            text = str(date).strip()
            normalized_date = norm_date(text) if text else ""
            if text and not normalized_date:
                raise AppError("日期格式应为 YYYY-MM-DD")
            invoice["date"] = normalized_date
        layout = self._layout()
        catalog, _error = self._catalog_for(layout)
        aliases = self._aliases()
        config = load_model_config()
        previous = {int(line["line_no"]): line for line in invoice.get("lines") or []}
        used = set()
        rebuilt = []
        next_no = max([int(line["line_no"]) for line in previous.values()] or [0]) + 1
        for raw in incoming or []:
            requested = int(raw.get("line_no") or 0)
            current = previous.get(requested)
            if current is None or requested in used:
                line_no = next_no
                next_no += 1
                current = {
                    "raw_name": "",
                    "issues": [],
                    "suggestion": "",
                    "confidence": None,
                    "map_status": "",
                    "name": "",
                    "category": "",
                    "map_reason": "",
                }
            else:
                line_no = requested
            used.add(line_no)
            category = str(raw.get("category") or "").strip()
            if category not in {"蔬菜", "肉类"}:
                category = ""
            hit = resolve_review(
                current.get("raw_name") or "",
                str(raw.get("name") or ""),
                category,
                catalog,
                aliases,
                layout,
                previous={**current, "category": current.get("category"), "name": current.get("name")},
                threshold=config.map_confidence,
            )
            # 人工改了类别后，旧的模型结论不能跨表沿用。
            if current.get("map_status") == "model" and current.get("category") != category:
                hit = resolve_review(
                    current.get("raw_name") or "",
                    str(raw.get("name") or ""),
                    category,
                    catalog,
                    aliases,
                    layout,
                    previous=None,
                    threshold=config.map_confidence,
                )
            qty_raw = raw.get("qty")
            amount_raw = raw.get("amount")
            issues = [issue for issue in (current.get("issues") or []) if issue == "缺少品名"]
            if qty_raw not in (None, "") and parse_decimal(qty_raw) is None:
                issues.append("数量无法解析")
            if amount_raw not in (None, "") and parse_decimal(amount_raw) is None:
                issues.append("金额无法解析")
            rebuilt.append({
                "line_no": line_no,
                "raw_name": current.get("raw_name") or "",
                "name": hit.name,
                "qty": dump_decimal(parse_decimal(qty_raw)),
                "unit": str(raw.get("unit") or "").strip(),
                "amount": dump_decimal(parse_decimal(amount_raw)),
                "category": category,
                "map_status": hit.status,
                "map_reason": hit.reason,
                "suggestion": hit.suggestion or current.get("suggestion") or "",
                "confidence": hit.confidence,
                "issues": issues,
            })
        invoice["lines"] = rebuilt

    def _learn_aliases(self, invoice: dict) -> None:
        layout = self._layout()
        catalog, _error = self._catalog_for(layout)
        if catalog is None:
            return
        path = local_aliases_path()
        from foodmix.settings import read_json
        raw = read_json(path, {"aliases": []})
        aliases = list(raw.get("aliases") or [])
        index = {
            ("".join(str(item.get("alias") or "").split()), str(item.get("category") or "")): item
            for item in aliases
        }
        changed = False
        for line in invoice.get("lines") or []:
            raw_name = str(line.get("raw_name") or "").strip()
            name = str(line.get("name") or "").strip()
            category = str(line.get("category") or "")
            if not raw_name or not name or category not in {"蔬菜", "肉类"}:
                continue
            if "".join(raw_name.split()) == "".join(name.split()):
                continue
            sheet = layout.category_sheets.get(category)
            if not sheet or catalog.find(sheet, name) is None:
                continue
            key = ("".join(raw_name.split()), category)
            current = index.get(key)
            if current is None:
                item = {"alias": raw_name, "name": name, "category": category}
                aliases.append(item)
                index[key] = item
                changed = True
            elif current.get("name") != name:
                current["name"] = name
                changed = True
        if changed:
            path.parent.mkdir(parents=True, exist_ok=True)
            import json
            path.write_text(json.dumps({"aliases": aliases}, ensure_ascii=False, indent=2), encoding="utf-8")

    def _public(self, invoice: dict) -> dict:
        layout = self._layout()
        catalog, error = self._catalog_for(layout)
        rules = self._rules()
        lines = []
        for line in invoice.get("lines") or []:
            preview = self._preview(invoice, line, catalog, rules, layout)
            lines.append({
                "line_no": line["line_no"],
                "raw_name": line.get("raw_name") or "",
                "name": line.get("name") or "",
                "qty": line.get("qty"),
                "unit": line.get("unit") or "",
                "amount": line.get("amount"),
                "category": line.get("category") or "",
                "map_status": line.get("map_status") or "unmapped",
                "map_reason": line.get("map_reason") or "",
                "suggestion": line.get("suggestion") or "",
                "confidence": line.get("confidence"),
                "issues": line.get("issues") or [],
                "preview": preview,
            })
        return {
            "id": invoice["id"],
            "filename": invoice["filename"],
            "status": invoice["status"],
            "date": invoice.get("date") or "",
            "supplier": invoice.get("supplier") or "",
            "error": invoice.get("error") or "",
            "warning": invoice.get("warning") or "",
            "confirmed_at": invoice.get("confirmed_at") or "",
            "catalog_error": error,
            "lines": lines,
        }

    def _preview(self, invoice: dict, line: dict, catalog: Catalog | None, rules, layout: SheetLayout) -> dict:
        empty = {"will_write": False, "sheet": "", "write_qty": None, "write_unit": "", "write_amount": None, "reason": ""}
        if catalog is None:
            empty["reason"] = "物料表未加载，还不能判断能否落表"
            return empty
        bill = BillLine(
            invoice_id=invoice["id"],
            date=norm_date(invoice.get("date")),
            raw_name=str(line.get("raw_name") or ""),
            name=str(line.get("name") or ""),
            qty=parse_decimal(line.get("qty")),
            unit=str(line.get("unit") or ""),
            amount=parse_decimal(line.get("amount")),
            category=str(line.get("category") or ""),
        )
        writes, skipped = build_plan([bill], catalog, rules, layout)
        if writes:
            write = writes[0]
            return {
                "will_write": True,
                "sheet": write.sheet,
                "write_qty": dump_decimal(write.qty),
                "write_unit": write.unit,
                "write_amount": dump_decimal(write.amount),
                "reason": f"将写入「{write.sheet}」",
            }
        empty["reason"] = skipped[0].reason if skipped else "不写入"
        empty["sheet"] = skipped[0].sheet if skipped else ""
        return empty
