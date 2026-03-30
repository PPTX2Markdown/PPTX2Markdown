"""Low-level Google Generative Language API client."""

from __future__ import annotations

import base64
import email.utils
import json
import os
import random
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Optional

from dotenv import find_dotenv, load_dotenv

from .image_preprocess import gemini_ready_image_path


class GoogleGenAIClientError(RuntimeError):
    """Raised when a Google GenAI request fails."""


_RATE_LIMIT_LOCK = threading.Lock()
_NEXT_REQUEST_AT = 0.0
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


def _wait_for_rate_limit(min_request_interval_sec: float) -> None:
    global _NEXT_REQUEST_AT

    interval = max(0.0, float(min_request_interval_sec))
    if interval <= 0:
        return

    while True:
        with _RATE_LIMIT_LOCK:
            now = time.monotonic()
            wait_sec = _NEXT_REQUEST_AT - now
            if wait_sec <= 0:
                _NEXT_REQUEST_AT = now + interval
                return
        time.sleep(min(wait_sec, 0.25))


def _parse_retry_after(value: Optional[str]) -> Optional[float]:
    raw = str(value or "").strip()
    if not raw:
        return None

    try:
        return max(0.0, float(raw))
    except ValueError:
        pass

    try:
        parsed = email.utils.parsedate_to_datetime(raw)
    except (TypeError, ValueError, IndexError, OverflowError):
        return None

    if parsed is None:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    delta = (parsed - datetime.now(timezone.utc)).total_seconds()
    return max(0.0, delta)


def _is_retryable_http_status(status_code: int) -> bool:
    return status_code == 429 or 500 <= status_code < 600


def _compute_backoff_delay(
    *,
    attempt_index: int,
    retry_after: Optional[str],
    base_backoff_sec: float,
    max_backoff_sec: float,
) -> float:
    retry_after_sec = _parse_retry_after(retry_after)
    if retry_after_sec is not None:
        return retry_after_sec

    capped_base = max(0.1, float(base_backoff_sec))
    capped_max = max(capped_base, float(max_backoff_sec))
    exponential = min(capped_max, capped_base * (2 ** max(0, attempt_index)))
    jitter = random.uniform(0.0, min(1.0, exponential * 0.25))
    return exponential + jitter


def generate_content(
    image_path: Path,
    *,
    model_id: str,
    prompt: str,
    max_output_tokens: int,
    api_key: str,
    max_retries: int = 5,
    base_backoff_sec: float = 2.0,
    max_backoff_sec: float = 30.0,
    min_request_interval_sec: float = 0.0,
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

    attempts = max(0, int(max_retries)) + 1
    last_error: Optional[Exception] = None
    for attempt_index in range(attempts):
        _wait_for_rate_limit(min_request_interval_sec)
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

            if attempt_index + 1 < attempts and _is_retryable_http_status(exc.code):
                delay = _compute_backoff_delay(
                    attempt_index=attempt_index,
                    retry_after=exc.headers.get("Retry-After"),
                    base_backoff_sec=base_backoff_sec,
                    max_backoff_sec=max_backoff_sec,
                )
                time.sleep(delay)
                continue
            raise GoogleGenAIClientError(f"HTTP {exc.code}: {message}") from exc
        except urllib.error.URLError as exc:
            last_error = exc
            if attempt_index + 1 < attempts:
                delay = _compute_backoff_delay(
                    attempt_index=attempt_index,
                    retry_after=None,
                    base_backoff_sec=base_backoff_sec,
                    max_backoff_sec=max_backoff_sec,
                )
                time.sleep(delay)
                continue
            raise GoogleGenAIClientError(f"{type(exc).__name__}: {exc}") from exc
        except Exception as exc:  # noqa: BLE001
            last_error = exc
            raise GoogleGenAIClientError(f"{type(exc).__name__}: {exc}") from exc

        response_error = _extract_error_message(response_payload)
        if response_error:
            raise GoogleGenAIClientError(response_error)
        return response_payload

    if last_error is not None:
        raise GoogleGenAIClientError(f"{type(last_error).__name__}: {last_error}") from last_error
    raise GoogleGenAIClientError("Gemini request failed without a response payload")
