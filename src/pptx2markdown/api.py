"""Programmatic API for pptx2markdown.

Example:
    import pptx2markdown

    pptx2markdown.convert("deck.pptx", output_dir="out")
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Optional, Sequence, Union

PathLike = Union[str, Path]


def convert(
    inputs: Union[PathLike, Sequence[PathLike], None] = None,
    *,
    output_dir: Optional[PathLike] = None,
    work_dir: Optional[PathLike] = None,
    reading_order: str = "xml",
    headings: str = "auto",
    strict_headings: bool = True,
    placeholder_inheritance: str = "style",
    inherited_shapes: str = "visible",
    reuse_surya_cache: bool = False,
    ppt_converter: str = "auto",
    image_vlm_provider: Optional[str] = None,
    image_vlm_model: Optional[str] = None,
    image_vlm_prompt: Optional[str] = None,
    image_vlm_max_new_tokens: Optional[int] = None,
    image_vlm_api_key_env: Optional[str] = None,
    ignore_image_vlm_cache: bool = False,
    verbose: bool = False,
) -> int:
    """Convert PPTX/PPT file(s) to markdown.

    Args:
        inputs: A path or list of paths to ``.pptx``/``.ppt`` files. If omitted,
            every presentation in the current directory is processed.
        output_dir: Where converted markdown is written
            (default: ``./output``).
        work_dir: Where intermediate artifacts and caches live
            (default: ``./.pptx2markdown``).
        reading_order: ``"xml"`` (default), ``"xycut"``, or ``"surya"``
            (requires the ``pptx2markdown[surya]`` extra).
        image_vlm_provider: Enable image-to-markdown via a VLM:
            ``"gemini"``, ``"openai"``, ``"openrouter"``, or ``"local"``
            (requires the ``pptx2markdown[local-vlm]`` extra).

    Returns:
        Process-style exit code: ``0`` on success, ``1`` if any slide failed.
    """
    from pptx2markdown.image_pipeline.service import (
        DEFAULT_GEMINI_API_KEY_ENV,
        DEFAULT_MAX_NEW_TOKENS,
        DEFAULT_PROMPT,
        DEFAULT_PROVIDER,
    )
    from pptx2markdown.main_converter.run_pptx_to_markdown import (
        _build_config,
        _configure_logging,
        run,
    )

    if inputs is None:
        input_list: list[str] = []
    elif isinstance(inputs, (str, Path)):
        input_list = [str(inputs)]
    else:
        input_list = [str(p) for p in inputs]

    args = argparse.Namespace(
        inputs=input_list,
        output_dir=str(output_dir) if output_dir else None,
        work_dir=str(work_dir) if work_dir else None,
        reading_order=reading_order,
        headings=headings,
        strict=bool(strict_headings),
        pptx_inheritance=placeholder_inheritance,
        inherited_shapes=inherited_shapes,
        reuse_surya_cache=bool(reuse_surya_cache),
        ppt_converter=ppt_converter,
        image_vlm_provider=image_vlm_provider or DEFAULT_PROVIDER,
        image_vlm_model=image_vlm_model,
        image_vlm_prompt=image_vlm_prompt if image_vlm_prompt is not None else DEFAULT_PROMPT,
        image_vlm_max_new_tokens=(
            image_vlm_max_new_tokens if image_vlm_max_new_tokens is not None else DEFAULT_MAX_NEW_TOKENS
        ),
        image_vlm_api_key_env=image_vlm_api_key_env or DEFAULT_GEMINI_API_KEY_ENV,
        ignore_image_vlm_cache=bool(ignore_image_vlm_cache),
        verbose=bool(verbose),
    )
    _configure_logging(verbose=bool(verbose))
    config = _build_config(args)
    return run(config)
