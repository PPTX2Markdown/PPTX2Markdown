"""Provider registry for pluggable image markdown backends."""

from __future__ import annotations

from typing import Dict, Optional

from .constants import DEFAULT_PROVIDER
from .provider_base import ImageMarkdownProvider
from .provider_google_gemini import GoogleGeminiImageProvider
from .provider_openai import OpenAIImageProvider
from .provider_openrouter import OpenRouterImageProvider
from .provider_qwen_local import QwenLocalImageProvider

_PROVIDER_REGISTRY: Dict[str, ImageMarkdownProvider] = {
    "local": QwenLocalImageProvider(),
    "gemini": GoogleGeminiImageProvider(),
    "openai": OpenAIImageProvider(),
    "openrouter": OpenRouterImageProvider(),
}

IMAGE_VLM_PROVIDERS = frozenset(_PROVIDER_REGISTRY)


def normalize_provider(provider: Optional[str]) -> str:
    normalized = str(provider or DEFAULT_PROVIDER).strip().lower()
    if normalized not in _PROVIDER_REGISTRY:
        raise ValueError(
            f"unsupported image VLM provider: {provider}. "
            f"Expected one of: {', '.join(sorted(_PROVIDER_REGISTRY))}"
        )
    return normalized


def get_provider(provider: Optional[str]) -> ImageMarkdownProvider:
    return _PROVIDER_REGISTRY[normalize_provider(provider)]
