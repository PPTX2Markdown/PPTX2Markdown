#!/usr/bin/env python3
"""Qwen2.5-VL based image-to-markdown service."""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Dict, Iterator, Optional, Tuple


QWEN_VL_MODELS = {
    "3b": "Qwen/Qwen2.5-VL-3B-Instruct",
    "7b": "Qwen/Qwen2.5-VL-7B-Instruct",
}

DEFAULT_PROMPT = (
    "Convert this image into Markdown.\n"
    "- Return Markdown only.\n"
    "- Preserve visible headings, paragraphs, bullet lists, numbered lists, tables, and code-like text.\n"
    "- If the image is a chart, diagram, infographic, or screenshot, summarize the visible content in clean Markdown.\n"
    "- If some text is unreadable, omit it instead of guessing.\n"
    "- Do not wrap the answer in triple backticks."
)

DEFAULT_MAX_NEW_TOKENS = 1024

_SUPPORTED_IMAGE_SUFFIXES = {
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".bmp",
    ".webp",
    ".tif",
    ".tiff",
}
_VECTOR_IMAGE_SUFFIXES = {
    ".emf",
    ".wmf",
}

_MODEL_CACHE: Dict[str, Dict[str, Any]] = {}
_RESULT_CACHE: Dict[str, Dict[str, Any]] = {}


def resolve_model_id(model_spec: str) -> Tuple[str, str]:
    normalized = str(model_spec or "").strip()
    if not normalized:
        raise ValueError("image VLM model is required")

    alias = normalized.lower()
    if alias in QWEN_VL_MODELS:
        return alias, QWEN_VL_MODELS[alias]

    if "/" in normalized:
        return normalized, normalized

    raise ValueError(
        f"unsupported image VLM model: {model_spec}. "
        f"Expected one of: {', '.join(sorted(QWEN_VL_MODELS))}"
    )


def _is_supported_image_suffix(suffix: str) -> bool:
    return suffix in _SUPPORTED_IMAGE_SUFFIXES or suffix in _VECTOR_IMAGE_SUFFIXES


def _resolve_vector_converter() -> Optional[str]:
    env_path = os.getenv("SOFFICE_PATH", "").strip()
    if env_path and Path(env_path).exists():
        return env_path

    if os.name == "nt":
        for candidate in (
            "C:/Program Files/LibreOffice/program/soffice.exe",
            "C:/Program Files (x86)/LibreOffice/program/soffice.exe",
        ):
            if Path(candidate).exists():
                return candidate

    for candidate in ("soffice", "libreoffice"):
        found = shutil.which(candidate)
        if found:
            return found
    soffice_exe = shutil.which("soffice.exe")
    if soffice_exe:
        return soffice_exe
    return None


def _rasterize_vector_image(image_path: Path) -> Path:
    converter = _resolve_vector_converter()
    if not converter:
        raise RuntimeError("LibreOffice is required to rasterize vector images such as EMF/WMF")

    with tempfile.TemporaryDirectory(prefix="vector_raster_") as tmpdir:
        tmp_root = Path(tmpdir)
        staged_input = tmp_root / image_path.name
        shutil.copy2(image_path, staged_input)
        proc = subprocess.run(
            [
                converter,
                "--headless",
                "--convert-to",
                "png",
                "--outdir",
                str(tmp_root),
                str(staged_input),
            ],
            text=True,
            capture_output=True,
            check=False,
        )
        if proc.returncode != 0:
            raise RuntimeError(
                "vector image rasterization failed\n"
                f"stdout={proc.stdout.strip()}\n"
                f"stderr={proc.stderr.strip()}"
            )

        output_path = tmp_root / f"{image_path.stem}.png"
        if not output_path.exists():
            pngs = sorted(tmp_root.glob("*.png"))
            if not pngs:
                raise RuntimeError(f"vector image rasterization produced no PNG output: {image_path.name}")
            output_path = pngs[0]
        persisted_output = Path(tempfile.mkdtemp(prefix="vector_raster_png_")) / output_path.name
        shutil.copy2(output_path, persisted_output)
        return persisted_output


@contextmanager
def _prepared_image_path(image_path: Path) -> Iterator[Path]:
    suffix = image_path.suffix.lower()
    if suffix in _VECTOR_IMAGE_SUFFIXES:
        raster_path = _rasterize_vector_image(image_path)
        try:
            yield raster_path
        finally:
            shutil.rmtree(raster_path.parent, ignore_errors=True)
        return
    yield image_path


def _pick_torch_dtype(torch: Any) -> Any:
    if torch.cuda.is_available():
        try:
            if torch.cuda.is_bf16_supported():
                return torch.bfloat16
        except Exception:
            pass
        return torch.float16
    return torch.float32


def _load_runtime(model_spec: str) -> Dict[str, Any]:
    model_alias, model_id = resolve_model_id(model_spec)
    cache_key = model_id
    cached = _MODEL_CACHE.get(cache_key)
    if cached is not None:
        return cached

    import torch
    from transformers import AutoProcessor, Qwen2_5_VLForConditionalGeneration

    model_kwargs: Dict[str, Any] = {
        "device_map": "auto",
        "torch_dtype": _pick_torch_dtype(torch),
        "low_cpu_mem_usage": True,
    }
    if torch.cuda.is_available():
        model_kwargs["attn_implementation"] = "sdpa"

    model = Qwen2_5_VLForConditionalGeneration.from_pretrained(model_id, **model_kwargs)
    processor = AutoProcessor.from_pretrained(model_id)
    runtime = {
        "model_alias": model_alias,
        "model_id": model_id,
        "model": model,
        "processor": processor,
        "torch": torch,
    }
    _MODEL_CACHE[cache_key] = runtime
    return runtime


def _normalize_markdown(text: str) -> str:
    normalized = text.strip()
    fence_match = re.fullmatch(r"```(?:markdown|md)?\s*(.*?)```", normalized, flags=re.DOTALL | re.IGNORECASE)
    if fence_match:
        normalized = fence_match.group(1).strip()
    return normalized.rstrip() + "\n" if normalized else ""


def extract_markdown_from_image(
    image_path: Path,
    *,
    model_spec: str,
    prompt: str = DEFAULT_PROMPT,
    max_new_tokens: int = DEFAULT_MAX_NEW_TOKENS,
) -> Dict[str, Any]:
    resolved_path = image_path.expanduser().resolve()
    cache_key = "||".join(
        [
            str(resolved_path),
            str(model_spec),
            str(max(1, int(max_new_tokens))),
            prompt,
        ]
    )
    cached = _RESULT_CACHE.get(cache_key)
    if cached is not None:
        return cached

    if not resolved_path.exists() or not resolved_path.is_file():
        out = {"status": "error", "file": str(resolved_path), "error": f"file not found: {resolved_path}"}
        _RESULT_CACHE[cache_key] = out
        return out

    suffix = resolved_path.suffix.lower()
    if not _is_supported_image_suffix(suffix):
        out = {
            "status": "error",
            "file": str(resolved_path),
            "error": f"unsupported image extension: {suffix or '(none)'}",
        }
        _RESULT_CACHE[cache_key] = out
        return out

    started_at = time.perf_counter()
    try:
        runtime = _load_runtime(model_spec)
        model = runtime["model"]
        processor = runtime["processor"]

        with _prepared_image_path(resolved_path) as prepared_path:
            messages = [
                {
                    "role": "user",
                    "content": [
                        {"type": "image", "path": str(prepared_path)},
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
            markdown = _normalize_markdown(decoded[0] if decoded else "")
    except Exception as exc:  # noqa: BLE001
        out = {
            "status": "error",
            "file": str(resolved_path),
            "model_alias": resolve_model_id(model_spec)[0],
            "model_id": resolve_model_id(model_spec)[1],
            "error": f"{type(exc).__name__}: {exc}",
            "elapsed_sec": round(time.perf_counter() - started_at, 3),
        }
        _RESULT_CACHE[cache_key] = out
        return out

    if not markdown.strip():
        out = {
            "status": "error",
            "file": str(resolved_path),
            "model_alias": runtime["model_alias"],
            "model_id": runtime["model_id"],
            "error": "model returned empty markdown",
            "elapsed_sec": round(time.perf_counter() - started_at, 3),
        }
        _RESULT_CACHE[cache_key] = out
        return out

    out = {
        "status": "markdown",
        "file": str(resolved_path),
        "model_alias": runtime["model_alias"],
        "model_id": runtime["model_id"],
        "prompt": prompt,
        "max_new_tokens": max(1, int(max_new_tokens)),
        "markdown": markdown,
        "elapsed_sec": round(time.perf_counter() - started_at, 3),
    }
    _RESULT_CACHE[cache_key] = out
    return out
