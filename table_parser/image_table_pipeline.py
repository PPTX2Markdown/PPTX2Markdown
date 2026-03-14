#!/usr/bin/env python3
"""Backward-compatible shim for the image table pipeline.

This module re-exports the refactored implementation from:
  image_pipeline/service.py
"""

from __future__ import annotations

from pathlib import Path
import sys


_REPO_ROOT = Path(__file__).resolve().parent.parent
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

from image_pipeline.service import classify_image, extract_table_markdown_from_image  # noqa: E402

__all__ = ["classify_image", "extract_table_markdown_from_image", "main"]


def main() -> int:
    print(
        "table_parser/image_table_pipeline.py is now a compatibility shim.\n"
        "Use image_pipeline.service.extract_table_markdown_from_image instead."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
