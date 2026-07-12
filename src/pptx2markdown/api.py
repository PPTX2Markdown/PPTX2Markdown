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
    output_format: str = "markdown",
    headings: str = "auto",
    placeholder_inheritance: str = "style",
    inherited_shapes: str = "visible",
    ppt_converter: str = "auto",
    verbose: bool = False,
) -> int:
    """Convert PPTX/PPT file(s) to Markdown or JSON.

    Args:
        inputs: A path or list of paths to ``.pptx``/``.ppt`` files. If omitted,
            every presentation in the current directory is processed.
        output_dir: Where converted output is written
            (default: ``./output``).
        work_dir: Where intermediate artifacts and caches live
            (default: ``./.pptx2markdown``).
        output_format: ``"markdown"`` (default) or ``"json"``.
    Returns:
        Process-style exit code: ``0`` on success, ``1`` if any slide failed.
    """
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
        output_format=output_format,
        headings=headings,
        pptx_inheritance=placeholder_inheritance,
        inherited_shapes=inherited_shapes,
        ppt_converter=ppt_converter,
        verbose=bool(verbose),
    )
    _configure_logging(verbose=bool(verbose))
    config = _build_config(args)
    return run(config)
