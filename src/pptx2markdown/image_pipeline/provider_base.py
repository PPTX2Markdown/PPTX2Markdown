"""Provider abstraction for image-to-markdown backends."""

from __future__ import annotations

from pathlib import Path
from typing import Optional, Protocol

from .schemas import ImageMarkdownResult, ModelResolution


class ImageMarkdownProvider(Protocol):
    provider_name: str

    def resolve_model_id(self, model_spec: Optional[str]) -> ModelResolution:
        ...

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
        ...
