#!/usr/bin/env python3
"""Image-to-markdown service supporting local Qwen2.5-VL and Gemini API."""

from __future__ import annotations

import base64
import json
import os
import re
import shutil
import subprocess
import tempfile
import time
import urllib.error
import urllib.parse
import urllib.request
import warnings
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Dict, Iterator, Optional, Tuple


IMAGE_VLM_PROVIDERS = {"local", "gemini"}
DEFAULT_PROVIDER = "local"
DEFAULT_GEMINI_MODEL = "gemini-2.5-flash"
DEFAULT_GEMINI_API_KEY_ENV = "GEMINI_API_KEY"

QWEN_VL_MODELS = {
    "3b": "Qwen/Qwen2.5-VL-3B-Instruct",
    "7b": "Qwen/Qwen2.5-VL-7B-Instruct",
}

DEFAULT_PROMPT = (
    "Extract only document-worthy information from this image as concise Markdown for retrieval.\n\n"
    "Rules:\n"
    "1. Keep only information that is useful for search, retrieval, or understanding the document.\n"
    "2. Omit unreadable, uncertain, or purely decorative content.\n"
    "3. Do not invent missing text or details.\n"
    "4. Do not wrap the answer in triple backticks.\n\n"
    "Output policy by image type:\n"
    "- If the image is mainly a table:\n"
    "  - Recreate the table as a Markdown table.\n"
    "  - Preserve readable headers, row labels, units, and cell values.\n"
    "  - Omit cells that are unreadable instead of guessing.\n"
    "- If the image contains formulas or equations:\n"
    "  - Extract them faithfully.\n"
    "  - Use LaTeX-style math notation when possible.\n"
    "  - Keep variable names, subscripts, superscripts, and operators.\n"
    "- If the image is mainly plaintext:\n"
    "  - Transcribe the readable text in clean Markdown.\n"
    "  - Preserve headings, bullet points, and short paragraph structure when visible.\n"
    "- If the image is an informative figure, chart, diagram, screenshot, or illustration:\n"
    "  - First write exactly one concise sentence summarizing what information the image conveys.\n"
    "  - Then list any clearly readable labels, legends, axis names, key values, or embedded text in Markdown bullets.\n"
    "- If the image is decorative, redundant, or not useful for document retrieval:\n"
    "  - Answer exactly: 불필요한 정보\n\n"
    "Additional constraints:\n"
    "- Prefer faithful extraction over fluent rewriting.\n"
    "- Keep the output concise, but do not drop important entities, numbers, labels, or relationships.\n"
    "- Preserve technical terms as written in the image.\n"
    "- If only part of the image is readable, extract only that readable part."
)

DEFAULT_MAX_NEW_TOKENS = 1024
DEFAULT_TRANSPARENT_BG_GRAY = 192

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
_GEMINI_MIME_BY_SUFFIX = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".webp": "image/webp",
}

_MODEL_CACHE: Dict[str, Dict[str, Any]] = {}
_RESULT_CACHE: Dict[str, Dict[str, Any]] = {}

warnings.filterwarnings(
    "ignore",
    message=r"Using `TRANSFORMERS_CACHE` is deprecated and will be removed in v5 of Transformers\..*",
    category=FutureWarning,
)


def normalize_provider(provider: Optional[str]) -> str:
    normalized = str(provider or DEFAULT_PROVIDER).strip().lower()
    if normalized not in IMAGE_VLM_PROVIDERS:
        raise ValueError(
            f"unsupported image VLM provider: {provider}. "
            f"Expected one of: {', '.join(sorted(IMAGE_VLM_PROVIDERS))}"
        )
    return normalized


def resolve_model_id(model_spec: Optional[str], *, provider: str = DEFAULT_PROVIDER) -> Tuple[str, str]:
    normalized_provider = normalize_provider(provider)
    normalized = str(model_spec or "").strip()

    if normalized_provider == "local":
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

    if normalized.startswith("models/"):
        normalized = normalized.split("/", 1)[1].strip()
    if not normalized:
        normalized = DEFAULT_GEMINI_MODEL
    return normalized, normalized


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


def _flatten_transparent_image(image_path: Path, bg_gray: int = DEFAULT_TRANSPARENT_BG_GRAY) -> Optional[Path]:
    from PIL import Image, ImageOps

    with Image.open(image_path) as loaded:
        image = ImageOps.exif_transpose(loaded)
        try:
            image.seek(0)
        except Exception:
            pass

        has_alpha = image.mode in {"RGBA", "LA"} or "transparency" in image.info
        if not has_alpha:
            return None

        rgba = image.convert("RGBA")
        bg_gray = max(0, min(255, int(bg_gray)))
        background = Image.new("RGBA", rgba.size, (bg_gray, bg_gray, bg_gray, 255))
        composited = Image.alpha_composite(background, rgba).convert("RGB")

        persisted_output = Path(tempfile.mkdtemp(prefix="prepared_image_png_")) / f"{image_path.stem}.png"
        composited.save(persisted_output, format="PNG")
        return persisted_output


@contextmanager
def _prepared_image_path(image_path: Path) -> Iterator[Path]:
    suffix = image_path.suffix.lower()
    raster_path: Optional[Path] = None
    flattened_path: Optional[Path] = None

    if suffix in _VECTOR_IMAGE_SUFFIXES:
        raster_path = _rasterize_vector_image(image_path)

    source_path = raster_path or image_path
    flattened_path = _flatten_transparent_image(source_path)
    prepared_path = flattened_path or source_path

    try:
        yield prepared_path
    finally:
        if flattened_path is not None:
            shutil.rmtree(flattened_path.parent, ignore_errors=True)
        if raster_path is not None:
            shutil.rmtree(raster_path.parent, ignore_errors=True)


@contextmanager
def _gemini_ready_image_path(image_path: Path) -> Iterator[Tuple[Path, str]]:
    suffix = image_path.suffix.lower()
    mime_type = _GEMINI_MIME_BY_SUFFIX.get(suffix)
    if mime_type:
        yield image_path, mime_type
        return

    from PIL import Image, ImageOps

    tmp_root = Path(tempfile.mkdtemp(prefix="gemini_image_png_"))
    output_path = tmp_root / f"{image_path.stem}.png"
    try:
        with Image.open(image_path) as loaded:
            image = ImageOps.exif_transpose(loaded)
            try:
                image.seek(0)
            except Exception:
                pass

            if image.mode in {"RGBA", "LA"} or "transparency" in image.info:
                image.convert("RGBA").save(output_path, format="PNG")
            else:
                image.convert("RGB").save(output_path, format="PNG")
        yield output_path, "image/png"
    finally:
        shutil.rmtree(tmp_root, ignore_errors=True)


def _pick_torch_dtype(torch: Any) -> Any:
    if torch.cuda.is_available():
        try:
            if torch.cuda.is_bf16_supported():
                return torch.bfloat16
        except Exception:
            pass
        return torch.float16
    return torch.float32


def _load_runtime(model_spec: Optional[str]) -> Dict[str, Any]:
    model_alias, model_id = resolve_model_id(model_spec, provider="local")
    cache_key = model_id
    cached = _MODEL_CACHE.get(cache_key)
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
    _MODEL_CACHE[cache_key] = runtime
    return runtime


_PROMPT_ECHO_PHRASES = {
    "return markdown only.",
    "extract only document-worthy information from this image as concise markdown for retrieval.",
    "rules:",
    "1. keep only information that is useful for search, retrieval, or understanding the document.",
    "2. omit unreadable, uncertain, or purely decorative content.",
    "3. do not invent missing text or details.",
    "4. do not wrap the answer in triple backticks.",
    "output policy by image type:",
    "if the image is mainly a table, recreate it as a markdown table.",
    "preserve readable headers, row labels, units, and cell values.",
    "omit cells that are unreadable instead of guessing.",
    "if the image contains formulas or equations:",
    "extract them faithfully.",
    "use latex-style math notation when possible.",
    "keep variable names, subscripts, superscripts, and operators.",
    "if the image is mainly plaintext:",
    "transcribe the readable text in clean markdown.",
    "preserve headings, bullet points, and short paragraph structure when visible.",
    "if the image is an informative figure, chart, diagram, screenshot, or illustration:",
    "first write exactly one concise sentence summarizing what information the image conveys.",
    "then list any clearly readable labels, legends, axis names, key values, or embedded text in markdown bullets.",
    "if the image is decorative, redundant, or not useful for document retrieval:",
    "answer exactly: 불필요한 정보",
    "additional constraints:",
    "prefer faithful extraction over fluent rewriting.",
    "keep the output concise, but do not drop important entities, numbers, labels, or relationships.",
    "preserve technical terms as written in the image.",
    "if only part of the image is readable, extract only that readable part.",
    "do not wrap the answer in triple backticks.",
}


def _unwrap_latex_text_line(line: str) -> str:
    match = re.fullmatch(r"\$\\text\s*\{\s*(.*?)\s*\}\$", line.strip())
    if match:
        return match.group(1).strip()
    return line


def _is_prompt_echo_line(line: str) -> bool:
    normalized = _unwrap_latex_text_line(line).strip()
    if not normalized:
        return False
    normalized = re.sub(r"^[\-\*\u2022]\s*", "", normalized)
    normalized = re.sub(r"\s+", " ", normalized).strip().lower()
    return normalized in _PROMPT_ECHO_PHRASES


def _normalize_markdown(text: str) -> str:
    normalized = text.strip()
    fence_match = re.fullmatch(r"```(?:markdown|md)?\s*(.*?)```", normalized, flags=re.DOTALL | re.IGNORECASE)
    if fence_match:
        normalized = fence_match.group(1).strip()
    kept_lines = [line.rstrip() for line in normalized.splitlines() if not _is_prompt_echo_line(line)]
    normalized = "\n".join(line for line in kept_lines).strip()
    return normalized.rstrip() + "\n" if normalized else ""


def _gemini_api_key(api_key: Optional[str], env_name: str) -> str:
    direct = str(api_key or "").strip()
    if direct:
        return direct
    return os.getenv(env_name, "").strip()


def _extract_gemini_text(payload: Dict[str, Any]) -> str:
    texts = []
    for candidate in payload.get("candidates", []):
        if not isinstance(candidate, dict):
            continue
        content = candidate.get("content")
        if not isinstance(content, dict):
            continue
        for part in content.get("parts", []):
            if not isinstance(part, dict):
                continue
            text = part.get("text")
            if isinstance(text, str) and text.strip():
                texts.append(text)
    return "\n".join(texts).strip()


def _extract_gemini_error(payload: Dict[str, Any]) -> Optional[str]:
    error = payload.get("error")
    if not isinstance(error, dict):
        return None
    message = error.get("message")
    if isinstance(message, str) and message.strip():
        return message.strip()
    return None


def _request_gemini_markdown(
    image_path: Path,
    *,
    model_spec: Optional[str],
    prompt: str,
    max_new_tokens: int,
    gemini_api_key: Optional[str],
    gemini_api_key_env: str,
) -> Dict[str, Any]:
    model_alias, model_id = resolve_model_id(model_spec, provider="gemini")
    api_key = _gemini_api_key(gemini_api_key, gemini_api_key_env)
    if not api_key:
        return {
            "status": "error",
            "provider": "gemini",
            "file": str(image_path),
            "model_alias": model_alias,
            "model_id": model_id,
            "error": f"Gemini API key not found. Set {gemini_api_key_env}.",
        }

    with _gemini_ready_image_path(image_path) as (request_image_path, mime_type):
        encoded_image = base64.b64encode(request_image_path.read_bytes()).decode("ascii")

    payload = {
        "contents": [
            {
                "role": "user",
                "parts": [
                    {"text": prompt},
                    {
                        "inlineData": {
                            "mimeType": mime_type,
                            "data": encoded_image,
                        }
                    },
                ],
            }
        ],
        "generationConfig": {
            "temperature": 0,
            "maxOutputTokens": max(1, int(max_new_tokens)),
        },
    }

    endpoint = (
        "https://generativelanguage.googleapis.com/v1beta/models/"
        f"{urllib.parse.quote(model_id, safe='')}:generateContent"
        f"?key={urllib.parse.quote(api_key, safe='')}"
    )
    request = urllib.request.Request(
        endpoint,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json; charset=utf-8"},
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            response_payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        message = body
        try:
            parsed = json.loads(body)
            message = _extract_gemini_error(parsed) or body
        except Exception:
            pass
        return {
            "status": "error",
            "provider": "gemini",
            "file": str(image_path),
            "model_alias": model_alias,
            "model_id": model_id,
            "error": f"HTTP {exc.code}: {message}",
        }
    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "provider": "gemini",
            "file": str(image_path),
            "model_alias": model_alias,
            "model_id": model_id,
            "error": f"{type(exc).__name__}: {exc}",
        }

    response_error = _extract_gemini_error(response_payload)
    if response_error:
        return {
            "status": "error",
            "provider": "gemini",
            "file": str(image_path),
            "model_alias": model_alias,
            "model_id": model_id,
            "error": response_error,
        }

    markdown = _normalize_markdown(_extract_gemini_text(response_payload))
    if not markdown.strip():
        return {
            "status": "no_markdown",
            "provider": "gemini",
            "file": str(image_path),
            "model_alias": model_alias,
            "model_id": model_id,
            "reason": "not_document_worthy",
            "fallback": "image_link",
        }

    return {
        "status": "markdown",
        "provider": "gemini",
        "file": str(image_path),
        "model_alias": model_alias,
        "model_id": model_id,
        "prompt": prompt,
        "max_new_tokens": max(1, int(max_new_tokens)),
        "markdown": markdown,
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
) -> Dict[str, Any]:
    normalized_provider = normalize_provider(provider)
    resolved_path = image_path.expanduser().resolve()
    cache_key = "||".join(
        [
            str(resolved_path),
            normalized_provider,
            str(model_spec or ""),
            str(max(1, int(max_new_tokens))),
            prompt,
            str(gemini_api_key_env if normalized_provider == "gemini" else ""),
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
        with _prepared_image_path(resolved_path) as prepared_path:
            if normalized_provider == "local":
                runtime = _load_runtime(model_spec)
                model = runtime["model"]
                processor = runtime["processor"]

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
                if not markdown.strip():
                    out = {
                        "status": "no_markdown",
                        "provider": "local",
                        "file": str(resolved_path),
                        "model_alias": runtime["model_alias"],
                        "model_id": runtime["model_id"],
                        "reason": "not_document_worthy",
                        "fallback": "image_link",
                        "elapsed_sec": round(time.perf_counter() - started_at, 3),
                    }
                    _RESULT_CACHE[cache_key] = out
                    return out

                out = {
                    "status": "markdown",
                    "provider": "local",
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

            out = _request_gemini_markdown(
                prepared_path,
                model_spec=model_spec,
                prompt=prompt,
                max_new_tokens=max_new_tokens,
                gemini_api_key=gemini_api_key,
                gemini_api_key_env=gemini_api_key_env,
            )
            out["file"] = str(resolved_path)
            out["elapsed_sec"] = round(time.perf_counter() - started_at, 3)
            _RESULT_CACHE[cache_key] = out
            return out
    except Exception as exc:  # noqa: BLE001
        model_alias, model_id = resolve_model_id(model_spec, provider=normalized_provider)
        out = {
            "status": "error",
            "provider": normalized_provider,
            "file": str(resolved_path),
            "model_alias": model_alias,
            "model_id": model_id,
            "error": f"{type(exc).__name__}: {exc}",
            "elapsed_sec": round(time.perf_counter() - started_at, 3),
        }
        _RESULT_CACHE[cache_key] = out
        return out
