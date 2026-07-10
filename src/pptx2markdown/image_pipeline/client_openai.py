"""Low-level OpenAI Responses API client for image markdown extraction."""

from __future__ import annotations

import base64
import json
import logging
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any, Dict, Optional

from .client_google_genai import (
    _compute_backoff_delay,
    _defer_next_request,
    _sleep_with_progress,
    _wait_for_rate_limit,
)


class OpenAIClientError(RuntimeError):
    """Raised when an OpenAI request fails."""


logger = logging.getLogger(__name__)


def extract_output_text(payload: Dict[str, Any]) -> str:
    output_text = payload.get("output_text")
    if isinstance(output_text, str) and output_text.strip():
        return output_text.strip()

    texts = []
    for item in payload.get("output", []):
        if not isinstance(item, dict):
            continue
        for content in item.get("content", []):
            if not isinstance(content, dict):
                continue
            text = content.get("text")
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


def generate_response(
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
    suffix = image_path.suffix.lower().lstrip(".") or "png"
    encoded_image = base64.b64encode(image_path.read_bytes()).decode("ascii")
    image_url = f"data:image/{suffix};base64,{encoded_image}"
    payload = {
        "model": model_id,
        "input": [
            {
                "role": "user",
                "content": [
                    {"type": "input_text", "text": prompt},
                    {"type": "input_image", "image_url": image_url},
                ],
            }
        ],
        "max_output_tokens": max(1, int(max_output_tokens)),
    }
    request = urllib.request.Request(
        "https://api.openai.com/v1/responses",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json; charset=utf-8",
        },
        method="POST",
    )

    attempts = max(0, int(max_retries)) + 1
    for attempt_index in range(attempts):
        _wait_for_rate_limit(min_request_interval_sec)
        logger.info(
            "  [image-vlm] OpenAI request: %s (model=%s, attempt=%d/%d)",
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
                    "  [image-vlm] OpenAI retry scheduled: %s "
                    "(model=%s, status=%s, wait=%.1fs, next_attempt=%d/%d)",
                    image_path.name,
                    model_id,
                    exc.code,
                    delay,
                    attempt_index + 2,
                    attempts,
                )
                _sleep_with_progress(delay, image_name=image_path.name, model_id=model_id)
                continue
            raise OpenAIClientError(f"HTTP {exc.code}: {message}") from exc
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
                    "  [image-vlm] OpenAI retry scheduled: %s "
                    "(model=%s, error=%s, wait=%.1fs, next_attempt=%d/%d)",
                    image_path.name,
                    model_id,
                    type(exc).__name__,
                    delay,
                    attempt_index + 2,
                    attempts,
                )
                _sleep_with_progress(delay, image_name=image_path.name, model_id=model_id)
                continue
            raise OpenAIClientError(f"{type(exc).__name__}: {exc}") from exc

        response_error = _extract_error_message(response_payload)
        if response_error:
            raise OpenAIClientError(response_error)
        logger.info(
            "  [image-vlm] OpenAI response received: %s (model=%s)",
            image_path.name,
            model_id,
        )
        return response_payload

    raise OpenAIClientError("OpenAI request failed without a response payload")
