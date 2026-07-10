"""Google Gemini provider adapter for the image pipeline."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional

from .client_google_genai import (
    GoogleGenAIClientError,
    extract_text,
    generate_content,
    resolve_api_key,
)
from .constants import (
    DEFAULT_GEMINI_API_KEY_ENV,
    DEFAULT_GEMINI_BASE_BACKOFF_SEC,
    DEFAULT_GEMINI_MAX_BACKOFF_SEC,
    DEFAULT_GEMINI_MAX_RETRIES,
    DEFAULT_GEMINI_MIN_REQUEST_INTERVAL_SEC,
    DEFAULT_GEMINI_MODEL,
)
from .markdown_postprocess import normalize_markdown
from .provider_base import (
    ImageMarkdownResult,
    ModelResolution,
    api_key_missing_result,
    client_error_result,
    markdown_result,
    no_markdown_result,
    read_float_env,
    read_int_env,
)

logger = logging.getLogger(__name__)


class GoogleGeminiImageProvider:
    provider_name = "gemini"

    def resolve_model_id(self, model_spec: Optional[str]) -> ModelResolution:
        normalized = str(model_spec or "").strip()
        if normalized.startswith("models/"):
            normalized = normalized.split("/", 1)[1].strip()
        if not normalized:
            normalized = DEFAULT_GEMINI_MODEL
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
        effective_api_key_env = (
            str(api_key_env or DEFAULT_GEMINI_API_KEY_ENV).strip() or DEFAULT_GEMINI_API_KEY_ENV
        )
        resolved_api_key = resolve_api_key(api_key, effective_api_key_env)
        if not resolved_api_key:
            return api_key_missing_result(
                provider=self.provider_name,
                image_path=image_path,
                model_alias=model_alias,
                model_id=model_id,
                provider_label="Gemini",
                api_key_env=effective_api_key_env,
            )

        try:
            response_payload = generate_content(
                image_path,
                model_id=model_id,
                prompt=prompt,
                max_output_tokens=max_new_tokens,
                api_key=resolved_api_key,
                max_retries=read_int_env("GEMINI_MAX_RETRIES", DEFAULT_GEMINI_MAX_RETRIES),
                base_backoff_sec=read_float_env(
                    "GEMINI_BASE_BACKOFF_SEC",
                    DEFAULT_GEMINI_BASE_BACKOFF_SEC,
                ),
                max_backoff_sec=read_float_env(
                    "GEMINI_MAX_BACKOFF_SEC",
                    DEFAULT_GEMINI_MAX_BACKOFF_SEC,
                ),
                min_request_interval_sec=read_float_env(
                    "GEMINI_MIN_REQUEST_INTERVAL_SEC",
                    DEFAULT_GEMINI_MIN_REQUEST_INTERVAL_SEC,
                ),
            )
        except GoogleGenAIClientError as exc:
            return client_error_result(
                provider=self.provider_name,
                image_path=image_path,
                model_alias=model_alias,
                model_id=model_id,
                error=str(exc),
            )

        markdown = normalize_markdown(extract_text(response_payload))
        if not markdown.strip():
            logger.info(
                "  [image-vlm] Gemini produced no document-worthy markdown: %s (model=%s)",
                image_path.name,
                model_id,
            )
            return no_markdown_result(
                provider=self.provider_name,
                image_path=image_path,
                model_alias=model_alias,
                model_id=model_id,
            )

        logger.info(
            "  [image-vlm] Gemini markdown extracted: %s (model=%s)",
            image_path.name,
            model_id,
        )
        return markdown_result(
            provider=self.provider_name,
            image_path=image_path,
            model_alias=model_alias,
            model_id=model_id,
            prompt=prompt,
            max_new_tokens=max_new_tokens,
            markdown=markdown,
        )
