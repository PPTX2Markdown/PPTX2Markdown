"""Command-line interface for the standalone rewrite."""

from __future__ import annotations

import argparse
import logging
from pathlib import Path
from typing import Sequence

from .api import convert_many


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pptx2markdown",
        description="Convert PPTX presentations into portable Markdown or JSON bundles.",
    )
    parser.add_argument("inputs", nargs="+", help="Local .pptx files")
    parser.add_argument(
        "-o",
        "--output-dir",
        default="output",
        help="Bundle root (default: ./output)",
    )
    parser.add_argument(
        "--output-format",
        choices=("markdown", "json"),
        default="markdown",
    )
    parser.add_argument(
        "--no-assets",
        action="store_true",
        help="Do not materialize embedded images or attachments",
    )
    parser.add_argument(
        "--overwrite",
        choices=("error", "replace"),
        default="error",
    )
    parser.add_argument(
        "--tolerant",
        action="store_true",
        help="Keep other slides when one slide cannot be parsed",
    )
    parser.add_argument("-v", "--verbose", action="store_true")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(message)s",
    )
    result = convert_many(
        [Path(value) for value in args.inputs],
        output_dir=args.output_dir,
        output_format=args.output_format,
        assets="omit" if args.no_assets else "copy",
        overwrite=args.overwrite,
        strict=not args.tolerant,
    )
    for saved in result.succeeded:
        logging.info("wrote %s", saved.document_path)
    for failure in result.failed:
        logging.error("%s: %s", failure.source, failure.error)
    return result.exit_code
