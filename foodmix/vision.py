from __future__ import annotations

import base64
import io

import httpx

from foodmix.errors import AppError
from foodmix.extract import parse_json_blob
from foodmix.settings import ModelConfig


def _shrink(data: bytes, mime: str) -> tuple[bytes, str]:
    try:
        from PIL import Image, ImageOps
        image = ImageOps.exif_transpose(Image.open(io.BytesIO(data)))
        if max(image.size) > 2048:
            image.thumbnail((2048, 2048))
        if image.mode not in ("RGB", "L"):
            image = image.convert("RGB")
        buffer = io.BytesIO()
        image.save(buffer, format="JPEG", quality=85)
        return buffer.getvalue(), "image/jpeg"
    except Exception:
        return data, mime or "application/octet-stream"


def _message_text(content) -> str:
    if isinstance(content, list):
        parts = []
        for part in content:
            if isinstance(part, dict):
                parts.append(str(part.get("text") or ""))
            else:
                parts.append(str(part))
        return "".join(parts)
    return str(content or "")


def chat_text(config: ModelConfig, payload: dict) -> str:
    if not config.api_key:
        raise AppError("未配置视觉模型密钥。请设置环境变量 HUNYUAN_API_KEY 或 MOLLI_ARK_API_KEY")
    url = config.api_base.rstrip("/") + "/chat/completions"
    headers = {"Authorization": f"Bearer {config.api_key}"}
    try:
        response = httpx.post(url, json=payload, headers=headers, timeout=120)
    except httpx.HTTPError as exc:
        raise AppError(f"模型请求失败：{exc.__class__.__name__}") from exc
    if response.status_code >= 400:
        raise AppError(f"模型接口返回 {response.status_code}")
    try:
        body = response.json()
        content = body["choices"][0]["message"]["content"]
    except (ValueError, KeyError, IndexError, TypeError) as exc:
        raise AppError("模型返回里没有文本内容") from exc
    return _message_text(content)


class VisionClient:
    def __init__(self, config: ModelConfig):
        self.config = config

    def extract(self, image_bytes: bytes, mime: str, prompt: str) -> str:
        shrunk, shrunk_mime = _shrink(image_bytes, mime)
        encoded = base64.b64encode(shrunk).decode("ascii")
        payload = {
            "model": self.config.vision_model,
            "temperature": self.config.temperature,
            "messages": [
                {"role": "system", "content": prompt},
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": "识别这张电子采购单。只输出 JSON。"},
                        {"type": "image_url", "image_url": {"url": f"data:{shrunk_mime};base64,{encoded}"}},
                    ],
                },
            ],
        }
        return chat_text(self.config, payload)


class MappingClient:
    def suggest_batch(self, queries: list[dict], prompt: str, config: ModelConfig) -> dict[str, dict]:
        if not queries or not config.api_key:
            return {}
        payload = {
            "model": config.vision_model,
            "temperature": config.temperature,
            "messages": [
                {"role": "system", "content": prompt},
                {"role": "user", "content": _query_text(queries)},
            ],
        }
        parsed = parse_json_blob(chat_text(config, payload))
        found: dict[str, dict] = {}
        for match in parsed.get("matches") or []:
            if not isinstance(match, dict):
                continue
            key = str(match.get("key") or "")
            if key:
                found[key] = match
        return found


def _query_text(queries: list[dict]) -> str:
    import json
    items = []
    for query in queries:
        items.append({
            "key": query["key"],
            "raw_name": query["raw_name"],
            "category": query["category"],
            "candidates": query["candidates"],
        })
    return json.dumps({"items": items}, ensure_ascii=False)
