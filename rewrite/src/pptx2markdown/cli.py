"""Command-line interface for the standalone rewrite."""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
from pathlib import Path
from typing import Sequence

from .api import convert_many, is_lock_file
from .recipe import Recipe, merge_recipes, recipe_json_schema
from .tables import MERGED_CELL_POLICIES


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pptx2markdown",
        description="Convert PPTX presentations into portable Markdown or JSON bundles.",
    )
    parser.add_argument(
        "inputs", nargs="*", help="Local .pptx files (.ppt with --ppt-converter libreoffice)"
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        default="output",
        help="Bundle root (default: ./output)",
    )
    parser.add_argument(
        "--output-format",
        choices=("markdown", "json", "chunks"),
        help="chunks writes one JSON line per slide chunk for retrieval pipelines",
        default="markdown",
    )
    parser.add_argument(
        "--max-chunk-chars",
        type=int,
        help="With --output-format chunks, split slides longer than this at block boundaries",
    )
    parser.add_argument(
        "--jobs",
        "-j",
        type=int,
        default=1,
        help="Convert this many decks at once in worker processes (0: one per CPU)",
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
        "--headings",
        choices=("auto", "strict"),
        default="auto",
        help="strict keeps only headings from title placeholders (recipe headings=placeholder)",
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
        "--audit",
        action="store_true",
        help="Print, as JSON, text typed on the slides that the conversion does not "
        "contain, instead of writing bundles; exits with 1 when any is found",
    )
    parser.add_argument(
        "--recipe-schema",
        action="store_true",
        help="Print the JSON Schema of a recipe and exit",
    )
    parser.add_argument(
        "--ppt-converter",
        choices=("none", "libreoffice"),
        default="none",
        help="libreoffice converts legacy .ppt inputs to .pptx with LibreOffice first",
    )
    parser.add_argument(
        "--vector-images",
        choices=("keep", "png"),
        default="keep",
        help="png has LibreOffice draw EMF/WMF images as PNG, which viewers and models can read",
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
    from .review import REVIEW_SCHEMA_VERSION, review, review_data

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
                ppt_converter=args.ppt_converter,
            ) as presentation:
                report.append(
                    {
                        "source": Path(value).name,
                        "schema_version": REVIEW_SCHEMA_VERSION,
                        "items": review_data(review(presentation.document)),
                    }
                )
        except (PPTX2MarkdownError, OSError, ValueError) as exc:
            failed = True
            report.append({"source": Path(value).name, "error": str(exc)})
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 1 if failed else 0


def _print_audit(args: argparse.Namespace) -> int:
    from .api import parse
    from .audit import audit
    from .errors import PPTX2MarkdownError
    from .ppt import ppt_to_pptx

    report: list[dict[str, object]] = []
    failed = False
    for value in args.inputs:
        try:
            source: str | bytes = value
            if Path(value).suffix.casefold() == ".ppt" and args.ppt_converter == "libreoffice":
                source = ppt_to_pptx(Path(value))
            # The converter is audited, not a recipe: recipes drop content on purpose.
            with parse(
                source,
                name=f"{Path(value).stem}.pptx",
                include_assets=False,
                strict=not args.tolerant,
                merged_cells=args.merged_cells,
                ppt_converter=args.ppt_converter,
            ) as presentation:
                result = audit(source, presentation.document)
            failed = failed or not result.ok
            report.append({"source": Path(value).name, **result.to_data()})
        except (PPTX2MarkdownError, OSError, ValueError) as exc:
            failed = True
            report.append({"source": Path(value).name, "error": str(exc)})
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 1 if failed else 0


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
    lock_files = [value for value in args.inputs if is_lock_file(value)]
    for value in lock_files:
        print(
            f"warning: skipping {value}: an Office lock file (the deck is or was open), "
            "not a presentation",
            file=sys.stderr,
        )
    args.inputs = [value for value in args.inputs if not is_lock_file(value)]
    if not args.inputs:
        return 0
    if args.headings == "strict":
        recipe = merge_recipes(recipe or Recipe(), Recipe(headings="placeholder"))
    if args.review:
        return _print_review(args, recipe)
    if args.audit:
        return _print_audit(args)
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
        max_chunk_chars=args.max_chunk_chars,
        ppt_converter=args.ppt_converter,
        vector_images=args.vector_images,
        jobs=args.jobs if args.jobs > 0 else (os.cpu_count() or 1),
    )
    for saved in result.succeeded:
        logging.info("wrote %s", saved.document_path)
    for failure in result.failed:
        logging.error("%s: %s", failure.source, failure.error)
    return result.exit_code
