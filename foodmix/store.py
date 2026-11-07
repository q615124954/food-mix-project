from __future__ import annotations

import json
from pathlib import Path


def empty_state() -> dict:
    return {
        "invoices": [],
        "ledger": {"material_file": "", "cells": []},
        "last_sync": None,
    }


def read_state(path: Path) -> dict:
    if not path.exists():
        return empty_state()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return empty_state()
    if not isinstance(data, dict):
        return empty_state()
    data.setdefault("invoices", [])
    ledger = data.get("ledger")
    if not isinstance(ledger, dict):
        data["ledger"] = {"material_file": "", "cells": []}
    else:
        ledger.setdefault("material_file", "")
        ledger.setdefault("cells", [])
    data.setdefault("last_sync", None)
    return data


def write_state(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(path)
