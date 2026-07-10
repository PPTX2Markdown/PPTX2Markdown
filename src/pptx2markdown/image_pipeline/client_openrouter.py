"""Low-level OpenRouter Chat Completions client for image markdown extraction."""

from __future__ import annotations

import base64
import json
import logging
import os
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any, Dict, List, Optional

from .client_google_genai import (
    _compute_backoff_delay,
    _defer_next_request,
    _sleep_with_progress,
    _wait_for_rate_limit,
)


class OpenRouterClientError(RuntimeError):
    """Raised when an OpenRouter request fails."""


logger = logging.getLogger(__name__)

_MIME_BY_SUFFIX = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".webp": "image/webp",
    ".gif": "image/gif",
}


def _image_mime_type(image_path: Path) -> str:
    return _MIME_BY_SUFFIX.get(image_path.suffix.lower(), "image/png")


def extract_output_text(payload: Dict[str, Any]) -> str:
    texts: List[str] = []
    for choice in payload.get("choices", []):
        if not isinstance(choice, dict):
            continue
        message = choice.get("message")
        if not isinstance(message, dict):
            continue
        content = message.get("content")
        if isinstance(content, str) and content.strip():
            texts.append(content.strip())
            continue
        if isinstance(content, list):
            for part in content:
                if not isinstance(part, dict):
                    continue
                text = part.get("text")
                if isinstance(text, str) and text.strip():
                    texts.append(text.strip())
    return "\n".join(texts).strip()


def _extract_error_message(payload: Dict[str, Any]) -> Optional[str]:
    error = payload.get("error")
    if not isinstance(error, dict):
        return None
    message = error.get("message")
    if isinstance(message, str) and message.strip():
        return message.strip()
    return None


def _request_headers(api_key: str) -> Dict[str, str]:
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json; charset=utf-8",
    }
    site_url = os.getenv("OPENROUTER_SITE_URL", "").strip()
    app_name = os.getenv("OPENROUTER_APP_NAME", "").strip()
    if site_url:
        headers["HTTP-Referer"] = site_url
    if app_name:
        headers["X-Title"] = app_name
    return headers


def generate_chat_completion(
    image_path: Path,
    *,
    model_id: str,
    prompt: str,
    max_output_tokens: int,
    api_key: str,
    max_retries: int = 2,
    base_backoff_sec: float = 2.0,
    max_backoff_sec: float = 30.0,
    min_request_interval_sec: float = 0.0,
) -> Dict[str, Any]:
    encoded_image = base64.b64encode(image_path.read_bytes()).decode("ascii")
    image_url = f"data:{_image_mime_type(image_path)};base64,{encoded_image}"
    payload = {
        "model": model_id,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": image_url}},
                ],
            }
        ],
        "temperature": 0,
        "max_tokens": max(1, int(max_output_tokens)),
    }
    request = urllib.request.Request(
        "https://openrouter.ai/api/v1/chat/completions",
        data=json.dumps(payload).encode("utf-8"),
        headers=_request_headers(api_key),
        method="POST",
    )

    attempts = max(0, int(max_retries)) + 1
    for attempt_index in range(attempts):
        _wait_for_rate_limit(min_request_interval_sec)
        logger.info(
            "  [image-vlm] OpenRouter request: %s (model=%s, attempt=%d/%d)",
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

            if attempt_index + 1 < attempts and (exc.code == 429 or 500 <= exc.code < 600):
                delay = _compute_backoff_delay(
                    attempt_index=attempt_index,
                    retry_after=exc.headers.get("Retry-After"),
                    error_text=message,
                    base_backoff_sec=base_backoff_sec,
                    max_backoff_sec=max_backoff_sec,
                )
                _defer_next_request(delay)
                logger.warning(
                    "  [image-vlm] OpenRouter retry scheduled: %s (model=%s, status=%s, wait=%.1fs, next_attempt=%d/%d)",
                    image_path.name,
                    model_id,
                    exc.code,
                    delay,
                    attempt_index + 2,
                    attempts,
                )
                _sleep_with_progress(delay, image_name=image_path.name, model_id=model_id)
                continue
            raise OpenRouterClientError(f"HTTP {exc.code}: {message}") from exc
        except urllib.error.URLError as exc:
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
                    "  [image-vlm] OpenRouter retry scheduled: %s (model=%s, error=%s, wait=%.1fs, next_attempt=%d/%d)",
                    image_path.name,
                    model_id,
                    type(exc).__name__,
                    delay,
                    attempt_index + 2,
                    attempts,
                )
                _sleep_with_progress(delay, image_name=image_path.name, model_id=model_id)
                continue
            raise OpenRouterClientError(f"{type(exc).__name__}: {exc}") from exc

        response_error = _extract_error_message(response_payload)
        if response_error:
            raise OpenRouterClientError(response_error)
        logger.info(
            "  [image-vlm] OpenRouter response received: %s (model=%s)",
            image_path.name,
            model_id,
        )
        return response_payload

    raise OpenRouterClientError("OpenRouter request failed without a response payload")
