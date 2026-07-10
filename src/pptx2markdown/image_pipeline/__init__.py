"""Image pipeline package for image-to-markdown conversion."""

from .service import (
    DEFAULT_GEMINI_API_KEY_ENV,
    DEFAULT_GEMINI_MODEL,
    DEFAULT_MAX_NEW_TOKENS,
    DEFAULT_OPENAI_API_KEY_ENV,
    DEFAULT_OPENAI_MODEL,
    DEFAULT_OPENROUTER_API_KEY_ENV,
    DEFAULT_OPENROUTER_MODEL,
    DEFAULT_PROMPT,
    DEFAULT_PROVIDER,
    IMAGE_VLM_PROVIDERS,
    extract_markdown_from_image,
    normalize_provider,
    resolve_model_id,
)

__all__ = [
    "DEFAULT_GEMINI_API_KEY_ENV",
    "DEFAULT_GEMINI_MODEL",
    "DEFAULT_MAX_NEW_TOKENS",
    "DEFAULT_OPENAI_API_KEY_ENV",
    "DEFAULT_OPENAI_MODEL",
    "DEFAULT_OPENROUTER_API_KEY_ENV",
    "DEFAULT_OPENROUTER_MODEL",
    "DEFAULT_PROMPT",
    "DEFAULT_PROVIDER",
    "IMAGE_VLM_PROVIDERS",
    "extract_markdown_from_image",
    "normalize_provider",
    "resolve_model_id",
]
