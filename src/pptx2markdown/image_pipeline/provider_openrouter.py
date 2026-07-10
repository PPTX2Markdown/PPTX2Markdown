"""OpenRouter image provider adapter for the image pipeline."""

from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Optional

from .client_google_genai import resolve_api_key
from .client_openrouter import OpenRouterClientError, extract_output_text, generate_chat_completion
from .constants import (
    DEFAULT_GEMINI_BASE_BACKOFF_SEC,
    DEFAULT_GEMINI_MAX_BACKOFF_SEC,
    DEFAULT_GEMINI_MAX_RETRIES,
    DEFAULT_GEMINI_MIN_REQUEST_INTERVAL_SEC,
    DEFAULT_OPENROUTER_API_KEY_ENV,
    DEFAULT_OPENROUTER_MODEL,
)
from .markdown_postprocess import normalize_markdown
from .schemas import ImageMarkdownResult, ModelResolution


logger = logging.getLogger(__name__)


class OpenRouterImageProvider:
    provider_name = "openrouter"

    @staticmethod
    def _read_int_env(name: str, default: int) -> int:
        raw = os.getenv(name, "").strip()
        if not raw:
            return default
        try:
            return max(0, int(raw))
        except ValueError:
            return default

    @staticmethod
    def _read_float_env(name: str, default: float) -> float:
        raw = os.getenv(name, "").strip()
        if not raw:
            return default
        try:
            return max(0.0, float(raw))
        except ValueError:
            return default

    def resolve_model_id(self, model_spec: Optional[str]) -> ModelResolution:
        normalized = str(model_spec or "").strip()
        if not normalized:
            normalized = DEFAULT_OPENROUTER_MODEL
        return normalized, normalized

    def extract_markdown(
        self,
        image_path: Path,
        *,
        model_spec: Optional[str],
        prompt: str,
        max_new_tokens: int,
        api_key: Optional[str] = None,
        api_key_env: Optional[str] = None,
    ) -> ImageMarkdownResult:
        model_alias, model_id = self.resolve_model_id(model_spec)
        effective_api_key_env = str(api_key_env or DEFAULT_OPENROUTER_API_KEY_ENV).strip() or DEFAULT_OPENROUTER_API_KEY_ENV
        resolved_api_key = resolve_api_key(api_key, effective_api_key_env)
        if not resolved_api_key:
            return {
                "status": "error",
                "provider": self.provider_name,
                "file": str(image_path),
                "model_alias": model_alias,
                "model_id": model_id,
                "error": f"OpenRouter API key not found. Set {effective_api_key_env}.",
            }

        try:
            response_payload = generate_chat_completion(
                image_path,
                model_id=model_id,
                prompt=prompt,
                max_output_tokens=max_new_tokens,
                api_key=resolved_api_key,
                max_retries=self._read_int_env("OPENROUTER_MAX_RETRIES", DEFAULT_GEMINI_MAX_RETRIES),
                base_backoff_sec=self._read_float_env(
                    "OPENROUTER_BASE_BACKOFF_SEC",
                    DEFAULT_GEMINI_BASE_BACKOFF_SEC,
                ),
                max_backoff_sec=self._read_float_env(
                    "OPENROUTER_MAX_BACKOFF_SEC",
                    DEFAULT_GEMINI_MAX_BACKOFF_SEC,
                ),
                min_request_interval_sec=self._read_float_env(
                    "OPENROUTER_MIN_REQUEST_INTERVAL_SEC",
                    DEFAULT_GEMINI_MIN_REQUEST_INTERVAL_SEC,
                ),
            )
        except OpenRouterClientError as exc:
            return {
                "status": "error",
                "provider": self.provider_name,
                "file": str(image_path),
                "model_alias": model_alias,
                "model_id": model_id,
                "error": str(exc),
            }

        markdown = normalize_markdown(extract_output_text(response_payload))
        if not markdown.strip():
            logger.info(
                "  [image-vlm] OpenRouter produced no document-worthy markdown: %s (model=%s)",
                image_path.name,
                model_id,
            )
            return {
                "status": "no_markdown",
                "provider": self.provider_name,
                "file": str(image_path),
                "model_alias": model_alias,
                "model_id": model_id,
                "reason": "not_document_worthy",
                "fallback": "image_link",
            }

        logger.info(
            "  [image-vlm] OpenRouter markdown extracted: %s (model=%s)",
            image_path.name,
            model_id,
        )
        return {
            "status": "markdown",
            "provider": self.provider_name,
            "file": str(image_path),
            "model_alias": model_alias,
            "model_id": model_id,
            "prompt": prompt,
            "max_new_tokens": max(1, int(max_new_tokens)),
            "markdown": markdown,
        }
