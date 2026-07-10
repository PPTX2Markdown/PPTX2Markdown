"""Provider abstraction for image-to-markdown backends."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Dict, Optional, Protocol, Tuple

ImageMarkdownResult = Dict[str, Any]
ModelResolution = Tuple[str, str]


class ImageMarkdownProvider(Protocol):
    provider_name: str

    def resolve_model_id(self, model_spec: Optional[str]) -> ModelResolution: ...

    def extract_markdown(
        self,
        image_path: Path,
        *,
        model_spec: Optional[str],
        prompt: str,
        max_new_tokens: int,
        api_key: Optional[str] = None,
        api_key_env: Optional[str] = None,
    ) -> ImageMarkdownResult: ...


def read_int_env(name: str, default: int) -> int:
    raw = os.getenv(name, "").strip()
    if not raw:
        return default
    try:
        return max(0, int(raw))
    except ValueError:
        return default


def read_float_env(name: str, default: float) -> float:
    raw = os.getenv(name, "").strip()
    if not raw:
        return default
    try:
        return max(0.0, float(raw))
    except ValueError:
        return default


def api_key_missing_result(
    *,
    provider: str,
    image_path: Path,
    model_alias: str,
    model_id: str,
    provider_label: str,
    api_key_env: str,
) -> ImageMarkdownResult:
    return {
        "status": "error",
        "provider": provider,
        "file": str(image_path),
        "model_alias": model_alias,
        "model_id": model_id,
        "error": f"{provider_label} API key not found. Set {api_key_env}.",
    }


def client_error_result(
    *,
    provider: str,
    image_path: Path,
    model_alias: str,
    model_id: str,
    error: str,
) -> ImageMarkdownResult:
    return {
        "status": "error",
        "provider": provider,
        "file": str(image_path),
        "model_alias": model_alias,
        "model_id": model_id,
        "error": error,
    }


def no_markdown_result(
    *,
    provider: str,
    image_path: Path,
    model_alias: str,
    model_id: str,
) -> ImageMarkdownResult:
    return {
        "status": "no_markdown",
        "provider": provider,
        "file": str(image_path),
        "model_alias": model_alias,
        "model_id": model_id,
        "reason": "not_document_worthy",
        "fallback": "image_link",
    }


def markdown_result(
    *,
    provider: str,
    image_path: Path,
    model_alias: str,
    model_id: str,
    prompt: str,
    max_new_tokens: int,
    markdown: str,
) -> ImageMarkdownResult:
    return {
        "status": "markdown",
        "provider": provider,
        "file": str(image_path),
        "model_alias": model_alias,
        "model_id": model_id,
        "prompt": prompt,
        "max_new_tokens": max(1, int(max_new_tokens)),
        "markdown": markdown,
    }
