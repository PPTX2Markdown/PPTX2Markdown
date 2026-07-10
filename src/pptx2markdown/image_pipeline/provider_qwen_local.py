"""Local Qwen-based image markdown provider."""

from __future__ import annotations

import warnings
from pathlib import Path
from typing import Any, Dict, Optional

from .constants import QWEN_VL_MODELS
from .markdown_postprocess import normalize_markdown
from .schemas import ImageMarkdownResult, ModelResolution


_MODEL_CACHE: Dict[str, Dict[str, Any]] = {}

warnings.filterwarnings(
    "ignore",
    message=r"Using `TRANSFORMERS_CACHE` is deprecated and will be removed in v5 of Transformers\..*",
    category=FutureWarning,
)


def _pick_torch_dtype(torch: Any) -> Any:
    if torch.cuda.is_available():
        try:
            if torch.cuda.is_bf16_supported():
                return torch.bfloat16
        except Exception:
            pass
        return torch.float16
    return torch.float32


class QwenLocalImageProvider:
    provider_name = "local"

    def resolve_model_id(self, model_spec: Optional[str]) -> ModelResolution:
        normalized = str(model_spec or "").strip()
        if not normalized:
            raise ValueError("image VLM model is required for local provider")

        alias = normalized.lower()
        if alias in QWEN_VL_MODELS:
            return alias, QWEN_VL_MODELS[alias]

        if "/" in normalized:
            return normalized, normalized

        raise ValueError(
            f"unsupported image VLM model: {model_spec}. "
            f"Expected one of: {', '.join(sorted(QWEN_VL_MODELS))} or a Hugging Face model id"
        )

    def _load_runtime(self, model_spec: Optional[str]) -> Dict[str, Any]:
        model_alias, model_id = self.resolve_model_id(model_spec)
        cached = _MODEL_CACHE.get(model_id)
        if cached is not None:
            return cached

        import torch
        from transformers import AutoProcessor, Qwen2_5_VLForConditionalGeneration
        from transformers.utils import logging as transformers_logging

        try:
            transformers_logging.disable_progress_bar()
        except Exception:
            pass
        transformers_logging.set_verbosity_error()

        try:
            from huggingface_hub.utils import disable_progress_bars

            disable_progress_bars()
        except Exception:
            pass

        model_kwargs: Dict[str, Any] = {
            "device_map": "auto",
            "dtype": _pick_torch_dtype(torch),
            "low_cpu_mem_usage": True,
        }
        if torch.cuda.is_available():
            model_kwargs["attn_implementation"] = "sdpa"

        model = Qwen2_5_VLForConditionalGeneration.from_pretrained(model_id, **model_kwargs)
        processor = AutoProcessor.from_pretrained(model_id, use_fast=False)
        runtime = {
            "model_alias": model_alias,
            "model_id": model_id,
            "model": model,
            "processor": processor,
            "torch": torch,
        }
        _MODEL_CACHE[model_id] = runtime
        return runtime

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
        del api_key, api_key_env

        runtime = self._load_runtime(model_spec)
        model = runtime["model"]
        processor = runtime["processor"]

        messages = [
            {
                "role": "user",
                "content": [
                    {"type": "image", "path": str(image_path)},
                    {"type": "text", "text": prompt},
                ],
            }
        ]
        inputs = processor.apply_chat_template(
            messages,
            add_generation_prompt=True,
            tokenize=True,
            return_dict=True,
            return_tensors="pt",
        )
        inputs = inputs.to(model.device)
        generated_ids = model.generate(**inputs, max_new_tokens=max(1, int(max_new_tokens)))
        trimmed_ids = generated_ids[:, inputs.input_ids.shape[1] :]
        decoded = processor.batch_decode(
            trimmed_ids,
            skip_special_tokens=True,
            clean_up_tokenization_spaces=True,
        )
        markdown = normalize_markdown(decoded[0] if decoded else "")

        if not markdown.strip():
            return {
                "status": "no_markdown",
                "provider": self.provider_name,
                "file": str(image_path),
                "model_alias": runtime["model_alias"],
                "model_id": runtime["model_id"],
                "reason": "not_document_worthy",
                "fallback": "image_link",
            }

        return {
            "status": "markdown",
            "provider": self.provider_name,
            "file": str(image_path),
            "model_alias": runtime["model_alias"],
            "model_id": runtime["model_id"],
            "prompt": prompt,
            "max_new_tokens": max(1, int(max_new_tokens)),
            "markdown": markdown,
        }
