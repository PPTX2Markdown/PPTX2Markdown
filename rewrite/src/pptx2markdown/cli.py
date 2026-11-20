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
        "--review",
        action="store_true",
        help="Print, as JSON, the conversion decisions worth checking with recipe "
        "rules for each answer, instead of writing bundles",
    )
    parser.add_argument(
        "--ask",
        metavar="REQUEST",
        help="Have a language model draft a recipe for the request (needs the agent "
        "extra and Anthropic credentials) and print it as JSON instead of converting",
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


def _print_review(args: argparse.Namespace, recipe: Recipe | None) -> int:
    from .api import parse
    from .errors import PPTX2MarkdownError
    from .review import review, review_data

    report: list[dict[str, object]] = []
    failed = False
    for value in args.inputs:
        try:
            with parse(
                value,
                include_assets=False,
                strict=not args.tolerant,
                merged_cells=args.merged_cells,
                recipe=recipe,
            ) as presentation:
                report.append(
                    {
                        "source": Path(value).name,
                        "items": review_data(review(presentation.document)),
                    }
                )
        except (PPTX2MarkdownError, OSError, ValueError) as exc:
            failed = True
            report.append({"source": Path(value).name, "error": str(exc)})
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 1 if failed else 0


def _print_draft(args: argparse.Namespace, recipe: Recipe | None) -> int:
    import sys

    from .agent import draft_recipe
    from .api import parse
    from .errors import PPTX2MarkdownError

    if len(args.inputs) != 1:
        print("--ask takes exactly one input", file=sys.stderr)
        return 2
    try:
        with parse(
            args.inputs[0],
            include_assets=False,
            strict=not args.tolerant,
            merged_cells=args.merged_cells,
            recipe=recipe,
        ) as presentation:
            draft = draft_recipe(presentation.document, args.ask)
    except (PPTX2MarkdownError, OSError, ValueError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    print(draft.summary, file=sys.stderr)
    for target in draft.unmatched:
        print(f"warning: rule target {target!r} selects nothing", file=sys.stderr)
    print(draft.recipe.model_dump_json(indent=2, exclude_none=True))
    return 0


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
    if args.review:
        return _print_review(args, recipe)
    if args.ask:
        return _print_draft(args, recipe)
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
