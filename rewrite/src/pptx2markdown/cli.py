"""Command-line interface for the standalone rewrite."""

from __future__ import annotations

import argparse
import json
import logging
from pathlib import Path
from typing import Sequence

from .api import convert_many
from .recipe import Recipe, recipe_json_schema
from .tables import MERGED_CELL_POLICIES


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pptx2markdown",
        description="Convert PPTX presentations into portable Markdown or JSON bundles.",
    )
    parser.add_argument("inputs", nargs="*", help="Local .pptx files")
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
        "--merged-cells",
        choices=MERGED_CELL_POLICIES,
        default="fill",
        help=(
            "What table cells covered by a merge show: fill repeats the merged value "
            "(default), header repeats it only in the header row, blank leaves them empty"
        ),
    )
    parser.add_argument(
        "--recipe",
        type=Path,
        help="JSON recipe applied after parsing (see --recipe-schema)",
    )
    parser.add_argument(
        "--recipe-schema",
        action="store_true",
        help="Print the JSON Schema of a recipe and exit",
    )
    parser.add_argument(
        "--tolerant",
        action="store_true",
        help="Keep other slides when one slide cannot be parsed",
    )
    parser.add_argument("-v", "--verbose", action="store_true")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = _parser()
    args = parser.parse_args(argv)
    if args.recipe_schema:
        print(json.dumps(recipe_json_schema(), indent=2, ensure_ascii=False))
        return 0
    if not args.inputs:
        parser.error("the following arguments are required: inputs")
    recipe = None
    if args.recipe is not None:
        try:
            recipe = Recipe.model_validate_json(args.recipe.read_text(encoding="utf-8"))
        except (OSError, ValueError) as exc:
            parser.error(f"invalid recipe {args.recipe}: {exc}")
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
        merged_cells=args.merged_cells,
        recipe=recipe,
    )
    for saved in result.succeeded:
        logging.info("wrote %s", saved.document_path)
    for failure in result.failed:
        logging.error("%s: %s", failure.source, failure.error)
    return result.exit_code
