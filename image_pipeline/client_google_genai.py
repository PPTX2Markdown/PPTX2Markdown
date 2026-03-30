"""Low-level Google Generative Language API client."""

from __future__ import annotations

import base64
import json
import os
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any, Dict, Optional

from dotenv import find_dotenv, load_dotenv

from .image_preprocess import gemini_ready_image_path


class GoogleGenAIClientError(RuntimeError):
    """Raised when a Google GenAI request fails."""


_DOTENV_LOADED = False


def _load_project_dotenv() -> None:
    global _DOTENV_LOADED
    if _DOTENV_LOADED:
        return
    dotenv_path = find_dotenv(usecwd=True)
    if dotenv_path:
        load_dotenv(dotenv_path=dotenv_path, override=False)
    _DOTENV_LOADED = True


def resolve_api_key(api_key: Optional[str], env_name: str) -> str:
    direct = str(api_key or "").strip()
    if direct:
        return direct
    _load_project_dotenv()
    return os.getenv(env_name, "").strip()


def extract_text(payload: Dict[str, Any]) -> str:
    texts = []
    for candidate in payload.get("candidates", []):
        if not isinstance(candidate, dict):
            continue
        content = candidate.get("content")
        if not isinstance(content, dict):
            continue
        for part in content.get("parts", []):
            if not isinstance(part, dict):
                continue
            text = part.get("text")
            if isinstance(text, str) and text.strip():
                texts.append(text)
    return "\n".join(texts).strip()


def _extract_error_message(payload: Dict[str, Any]) -> Optional[str]:
    error = payload.get("error")
    if not isinstance(error, dict):
        return None
    message = error.get("message")
    if isinstance(message, str) and message.strip():
        return message.strip()
    return None


def generate_content(
    image_path: Path,
    *,
    model_id: str,
    prompt: str,
    max_output_tokens: int,
    api_key: str,
) -> Dict[str, Any]:
    with gemini_ready_image_path(image_path) as (request_image_path, mime_type):
        encoded_image = base64.b64encode(request_image_path.read_bytes()).decode("ascii")

    payload = {
        "contents": [
            {
                "role": "user",
                "parts": [
                    {"text": prompt},
                    {
                        "inlineData": {
                            "mimeType": mime_type,
                            "data": encoded_image,
                        }
                    },
                ],
            }
        ],
        "generationConfig": {
            "temperature": 0,
            "maxOutputTokens": max(1, int(max_output_tokens)),
        },
    }

    endpoint = (
        "https://generativelanguage.googleapis.com/v1beta/models/"
        f"{urllib.parse.quote(model_id, safe='')}:generateContent"
        f"?key={urllib.parse.quote(api_key, safe='')}"
    )
    request = urllib.request.Request(
        endpoint,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json; charset=utf-8"},
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            response_payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        message = body
        try:
            parsed = json.loads(body)
            message = _extract_error_message(parsed) or body
        except Exception:
            pass
        raise GoogleGenAIClientError(f"HTTP {exc.code}: {message}") from exc
    except Exception as exc:  # noqa: BLE001
        raise GoogleGenAIClientError(f"{type(exc).__name__}: {exc}") from exc

    response_error = _extract_error_message(response_payload)
    if response_error:
        raise GoogleGenAIClientError(response_error)
    return response_payload
