from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

from foodmix.errors import AppError

ROOT = Path(__file__).resolve().parent.parent
CONFIG_DIR = ROOT / "config"
PROMPTS_DIR = ROOT / "prompts"
WEBAPP_DIR = ROOT / "webapp"

load_dotenv(ROOT / ".env")


def _env_path(name: str, default: Path) -> Path:
    raw = os.getenv(name, "").strip()
    return Path(raw) if raw else default


def state_path() -> Path:
    return _env_path("FOODMIX_STATE_PATH", WEBAPP_DIR / "state.json")


def settings_path() -> Path:
    return _env_path("FOODMIX_SETTINGS_PATH", WEBAPP_DIR / "settings.json")


def local_aliases_path() -> Path:
    return _env_path("FOODMIX_ALIASES_PATH", WEBAPP_DIR / "aliases.local.json")


def read_json(path: Path, default):
    if not path.exists():
        return default
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


@dataclass(frozen=True)
class ModelConfig:
    api_base: str
    vision_model: str
    map_confidence: float
    temperature: float

    @property
    def api_key(self) -> str:
        return (os.getenv("HUNYUAN_API_KEY") or os.getenv("MOLLI_ARK_API_KEY") or "").strip()

    @property
    def has_key(self) -> bool:
        return bool(self.api_key)


def load_model_config() -> ModelConfig:
    raw = read_json(CONFIG_DIR / "models.json", {})
    return ModelConfig(
        api_base=os.getenv("VISION_API_BASE", str(raw.get("api_base") or "https://ai.gitee.com/v1")).rstrip("/"),
        vision_model=os.getenv("VISION_MODEL", str(raw.get("vision_model") or "Qwen3.5-Flash")),
        map_confidence=float(os.getenv("MAP_CONFIDENCE", raw.get("map_confidence", 0.8))),
        temperature=float(os.getenv("VISION_TEMPERATURE", raw.get("temperature", 0))),
    )


def prompt_text(name: str) -> str:
    directory = _env_path("FOODMIX_PROMPTS_DIR", PROMPTS_DIR)
    path = directory / name
    if not path.is_file():
        raise AppError(f"找不到提示词文件：{path}")
    return path.read_text(encoding="utf-8")


@dataclass(frozen=True)
class SheetLayout:
    name_column: int
    unit_column: int
    code_column: int
    amount_column_offset: int
    category_sheets: dict[str, str]
    writable_sheets: tuple[str, ...]
    never_write_sheets: frozenset[str]


def load_sheet_layout() -> SheetLayout:
    raw = read_json(CONFIG_DIR / "sheet_layout.json", {})
    category_sheets = {str(k): str(v) for k, v in (raw.get("category_sheets") or {}).items()}
    writable = tuple(str(item) for item in (raw.get("writable_sheets") or []))
    never = frozenset(str(item) for item in (raw.get("never_write_sheets") or []))
    for category, sheet in category_sheets.items():
        if sheet in never:
            raise AppError(f"配置错误：{category} 指向一期禁止写入的「{sheet}」")
        if sheet not in writable:
            raise AppError(f"配置错误：{category} 的「{sheet}」不在可写表里")
    return SheetLayout(
        name_column=int(raw.get("name_column") or 4),
        unit_column=int(raw.get("unit_column") or 5),
        code_column=int(raw.get("code_column") or 2),
        amount_column_offset=int(raw.get("amount_column_offset") or 2),
        category_sheets=category_sheets,
        writable_sheets=writable,
        never_write_sheets=never,
    )


@dataclass(frozen=True)
class Paths:
    input_dir: str
    material_file: str


def load_paths() -> Paths:
    raw = read_json(settings_path(), {})
    return Paths(
        input_dir=str(raw.get("input_dir") or os.getenv("INPUT_DIR") or "").strip(),
        material_file=str(raw.get("material_file") or os.getenv("MATERIAL_FILE") or "").strip(),
    )


def save_paths(paths: Paths) -> None:
    path = settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {"input_dir": paths.input_dir, "material_file": paths.material_file}
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
