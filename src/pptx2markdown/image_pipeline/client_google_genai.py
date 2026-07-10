"""Low-level Google Generative Language API client."""

from __future__ import annotations

import base64
import email.utils
import json
import logging
import os
import random
import re
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


logger = logging.getLogger(__name__)

_RATE_LIMIT_LOCK = threading.Lock()
_NEXT_REQUEST_AT = 0.0
_DOTENV_LOADED = False
_QUOTA_EXCEEDED_MESSAGE: Optional[str] = None


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


def _quota_exceeded_message() -> Optional[str]:
    with _RATE_LIMIT_LOCK:
        return _QUOTA_EXCEEDED_MESSAGE


def _mark_quota_exceeded(message: str, *, model_id: str) -> None:
    global _QUOTA_EXCEEDED_MESSAGE

    normalized = str(message or "").strip() or "Gemini quota exceeded."
    with _RATE_LIMIT_LOCK:
        first_time = _QUOTA_EXCEEDED_MESSAGE is None
        _QUOTA_EXCEEDED_MESSAGE = normalized
    if first_time:
        logger.warning(
            "  [image-vlm] Gemini quota exceeded for this run; all remaining Gemini image requests will fall back without retry. (model=%s)",
            model_id,
        )


def _defer_next_request(delay_sec: float) -> None:
    global _NEXT_REQUEST_AT

    delay = max(0.0, float(delay_sec))
    if delay <= 0:
        return

    with _RATE_LIMIT_LOCK:
        _NEXT_REQUEST_AT = max(_NEXT_REQUEST_AT, time.monotonic() + delay)


def _sleep_with_progress(delay_sec: float, *, image_name: str, model_id: str) -> None:
    remaining = max(0.0, float(delay_sec))
    if remaining <= 0:
        return

    # Keep the process chatty during long quota waits so it does not look hung.
    while remaining > 0:
        if remaining > 10:
            chunk = min(10.0, remaining)
        elif remaining > 5:
            chunk = 5.0
        else:
            chunk = remaining
        logger.info(
            "  [image-vlm] Gemini waiting: %s (model=%s, remaining=%.1fs)",
            image_name,
            model_id,
            remaining,
        )
        time.sleep(chunk)
        remaining = max(0.0, remaining - chunk)


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


def _parse_retry_delay_from_text(value: Optional[str]) -> Optional[float]:
    raw = str(value or "").strip()
    if not raw:
        return None

    match = re.search(r"Please retry in\s+([0-9]+(?:\.[0-9]+)?)s", raw, flags=re.IGNORECASE)
    if not match:
        return None
    try:
        return max(0.0, float(match.group(1)))
    except ValueError:
        return None


def _is_retryable_http_status(status_code: int) -> bool:
    return status_code == 429 or 500 <= status_code < 600


def _is_quota_exceeded_message(message: Optional[str]) -> bool:
    normalized = str(message or "").strip().lower()
    return "quota exceeded" in normalized or "billing details" in normalized


def _compute_backoff_delay(
    *,
    attempt_index: int,
    retry_after: Optional[str],
    error_text: Optional[str],
    base_backoff_sec: float,
    max_backoff_sec: float,
) -> float:
    retry_after_sec = _parse_retry_after(retry_after)
    if retry_after_sec is not None:
        return retry_after_sec

    retry_from_text_sec = _parse_retry_delay_from_text(error_text)
    if retry_from_text_sec is not None:
        return retry_from_text_sec

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
    quota_message = _quota_exceeded_message()
    if quota_message is not None:
        logger.info(
            "  [image-vlm] Gemini skipped due to earlier quota exhaustion: %s (model=%s)",
            image_path.name,
            model_id,
        )
        raise GoogleGenAIClientError(quota_message)

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
        logger.info(
            "  [image-vlm] Gemini request: %s (model=%s, attempt=%d/%d)",
            image_path.name,
            model_id,
            attempt_index + 1,
            attempts,
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

            if exc.code == 429 and _is_quota_exceeded_message(message):
                _mark_quota_exceeded(message, model_id=model_id)
                raise GoogleGenAIClientError(f"HTTP {exc.code}: {message}") from exc

            if attempt_index + 1 < attempts and _is_retryable_http_status(exc.code):
                delay = _compute_backoff_delay(
                    attempt_index=attempt_index,
                    retry_after=exc.headers.get("Retry-After"),
                    error_text=message,
                    base_backoff_sec=base_backoff_sec,
                    max_backoff_sec=max_backoff_sec,
                )
                _defer_next_request(delay)
                logger.warning(
                    "  [image-vlm] Gemini retry scheduled: %s (model=%s, status=%s, wait=%.1fs, next_attempt=%d/%d)",
                    image_path.name,
                    model_id,
                    exc.code,
                    delay,
                    attempt_index + 2,
                    attempts,
                )
                _sleep_with_progress(delay, image_name=image_path.name, model_id=model_id)
                continue
            raise GoogleGenAIClientError(f"HTTP {exc.code}: {message}") from exc
        except urllib.error.URLError as exc:
            last_error = exc
            if attempt_index + 1 < attempts:
                delay = _compute_backoff_delay(
                    attempt_index=attempt_index,
                    retry_after=None,
                    error_text=str(exc),
                    base_backoff_sec=base_backoff_sec,
                    max_backoff_sec=max_backoff_sec,
                )
                _defer_next_request(delay)
                logger.warning(
                    "  [image-vlm] Gemini retry scheduled: %s (model=%s, error=%s, wait=%.1fs, next_attempt=%d/%d)",
                    image_path.name,
                    model_id,
                    type(exc).__name__,
                    delay,
                    attempt_index + 2,
                    attempts,
                )
                _sleep_with_progress(delay, image_name=image_path.name, model_id=model_id)
                continue
            raise GoogleGenAIClientError(f"{type(exc).__name__}: {exc}") from exc
        except Exception as exc:  # noqa: BLE001
            last_error = exc
            raise GoogleGenAIClientError(f"{type(exc).__name__}: {exc}") from exc

        response_error = _extract_error_message(response_payload)
        if response_error:
            raise GoogleGenAIClientError(response_error)
        logger.info(
            "  [image-vlm] Gemini response received: %s (model=%s)",
            image_path.name,
            model_id,
        )
        return response_payload

    if last_error is not None:
        raise GoogleGenAIClientError(f"{type(last_error).__name__}: {last_error}") from last_error
    raise GoogleGenAIClientError("Gemini request failed without a response payload")
