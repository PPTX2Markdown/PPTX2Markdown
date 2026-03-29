#!/usr/bin/env python3
"""Public orchestration entry points for the image-to-markdown pipeline."""

from __future__ import annotations

import time
from pathlib import Path
from typing import Dict, Optional

from .constants import (
    DEFAULT_GEMINI_API_KEY_ENV,
    DEFAULT_GEMINI_MODEL,
    DEFAULT_MAX_NEW_TOKENS,
    DEFAULT_PROMPT,
    DEFAULT_PROVIDER,
)
from .image_preprocess import is_supported_image_suffix, prepared_image_path
from .provider_registry import IMAGE_VLM_PROVIDERS, get_provider, normalize_provider
from .schemas import ImageMarkdownResult, ModelResolution


_RESULT_CACHE: Dict[str, ImageMarkdownResult] = {}


def resolve_model_id(model_spec: Optional[str], *, provider: str = DEFAULT_PROVIDER) -> ModelResolution:
    return get_provider(provider).resolve_model_id(model_spec)


def _build_cache_key(
    image_path: Path,
    *,
    provider: str,
    model_spec: Optional[str],
    prompt: str,
    max_new_tokens: int,
    api_key_env: str,
) -> str:
    return "||".join(
        [
            str(image_path),
            provider,
            str(model_spec or ""),
            str(max(1, int(max_new_tokens))),
            prompt,
            str(api_key_env if provider == "gemini" else ""),
        ]
    )


def _error_result(
    *,
    file_path: Path,
    provider: str,
    model_spec: Optional[str],
    error: str,
    elapsed_sec: float,
) -> ImageMarkdownResult:
    try:
        model_alias, model_id = resolve_model_id(model_spec, provider=provider)
    except Exception:
        fallback_model = str(model_spec or "").strip()
        model_alias = fallback_model
        model_id = fallback_model

    return {
        "status": "error",
        "provider": provider,
        "file": str(file_path),
        "model_alias": model_alias,
        "model_id": model_id,
        "error": error,
        "elapsed_sec": round(elapsed_sec, 3),
    }


def extract_markdown_from_image(
    image_path: Path,
    *,
    model_spec: Optional[str],
    prompt: str = DEFAULT_PROMPT,
    max_new_tokens: int = DEFAULT_MAX_NEW_TOKENS,
    provider: str = DEFAULT_PROVIDER,
    gemini_api_key: Optional[str] = None,
    gemini_api_key_env: str = DEFAULT_GEMINI_API_KEY_ENV,
) -> ImageMarkdownResult:
    normalized_provider = normalize_provider(provider)
    provider_impl = get_provider(normalized_provider)
    resolved_path = image_path.expanduser().resolve()
    cache_key = _build_cache_key(
        resolved_path,
        provider=normalized_provider,
        model_spec=model_spec,
        prompt=prompt,
        max_new_tokens=max_new_tokens,
        api_key_env=gemini_api_key_env,
    )
    cached = _RESULT_CACHE.get(cache_key)
    if cached is not None:
        return cached

    if not resolved_path.exists() or not resolved_path.is_file():
        out: ImageMarkdownResult = {
            "status": "error",
            "file": str(resolved_path),
            "error": f"file not found: {resolved_path}",
        }
        _RESULT_CACHE[cache_key] = out
        return out

    suffix = resolved_path.suffix.lower()
    if not is_supported_image_suffix(suffix):
        out = {
            "status": "error",
            "file": str(resolved_path),
            "error": f"unsupported image extension: {suffix or '(none)'}",
        }
        _RESULT_CACHE[cache_key] = out
        return out

    started_at = time.perf_counter()
    try:
        with prepared_image_path(resolved_path) as prepared_path:
            out = provider_impl.extract_markdown(
                prepared_path,
                model_spec=model_spec,
                prompt=prompt,
                max_new_tokens=max_new_tokens,
                api_key=gemini_api_key,
                api_key_env=gemini_api_key_env,
            )
            out["file"] = str(resolved_path)
            out["elapsed_sec"] = round(time.perf_counter() - started_at, 3)
            _RESULT_CACHE[cache_key] = out
            return out
    except Exception as exc:  # noqa: BLE001
        out = _error_result(
            file_path=resolved_path,
            provider=normalized_provider,
            model_spec=model_spec,
            error=f"{type(exc).__name__}: {exc}",
            elapsed_sec=time.perf_counter() - started_at,
        )
        _RESULT_CACHE[cache_key] = out
        return out
