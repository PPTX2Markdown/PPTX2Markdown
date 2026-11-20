"""Model Context Protocol server: let an agent convert, review and adjust decks.

The tools mirror the recipe workflow: read a deck with block ids, look at the
converter's open questions, check a recipe against the deck, and get chunks
with the recipe applied. Nothing is written to disk. Run it with
``pptx2markdown-mcp`` (stdio); install with ``pip install "pptx2markdown[mcp]"``.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from pydantic import ValidationError

from .agent import document_view
from .api import parse
from .audit import audit
from .chunks import slide_chunks
from .model import ImageBlock
from .presentation import ParsedPresentation
from .recipe import Recipe, explain_recipe, recipe_json_schema
from .review import review, review_data

INSTRUCTIONS = """\
Converts PowerPoint (.pptx) files to Markdown for reading and retrieval.

Read a deck with convert_presentation: every block is preceded by a comment such as
`<!-- s3.b2 table -->` giving its id and kind. review_presentation lists the
conversion decisions worth checking; each option carries a ready-made recipe rule.
audit_presentation reports any text on the slides that the conversion lost.
presentation_images shows pictures that have no description, so you can write
description rules for them. The prepare_for_retrieval and apply_request prompts
walk through these steps.
To change the output (drop slides, set table header rows, turn text into a
heading, describe an image), write a recipe (see recipe_schema), check it with
check_recipe so every rule selects something, then pass it to convert_presentation
or presentation_chunks. Recipes never change the source file."""


def _open(
    path: str, recipe: dict[str, Any] | None, *, include_assets: bool = False
) -> ParsedPresentation:
    source = Path(path).expanduser()
    if source.suffix.casefold() != ".pptx":
        raise ValueError(f"not a .pptx file: {path}")
    if not source.is_file():
        raise ValueError(f"file not found: {path}")
    try:
        parsed_recipe = Recipe.model_validate(recipe) if recipe else None
    except ValidationError as exc:
        raise ValueError(f"invalid recipe: {exc}") from exc
    return parse(source, include_assets=include_assets, strict=False, recipe=parsed_recipe)


def convert_presentation(path: str, recipe: dict[str, Any] | None = None) -> str:
    """Convert a .pptx file to Markdown with block ids, optionally applying a recipe.

    Args:
        path: Local path to the .pptx file.
        recipe: Optional recipe (see recipe_schema) applied before rendering.
    """

    with _open(path, recipe) as presentation:
        return document_view(presentation.document)


def review_presentation(path: str) -> list[dict[str, Any]]:
    """List conversion decisions worth checking, each option with its recipe rule.

    Args:
        path: Local path to the .pptx file.
    """

    with _open(path, None) as presentation:
        return review_data(review(presentation.document))


def audit_presentation(path: str) -> dict[str, Any]:
    """Check the conversion for text typed on the slides that it does not contain.

    Returns coverage, slides missing from the conversion, and each lost piece of
    text with its slide and shape id. Losses mean the converter dropped content:
    read those slides with convert_presentation before relying on them.

    Args:
        path: Local path to the .pptx file.
    """

    with _open(path, None) as presentation:
        return audit(presentation.source_path or path, presentation.document).to_data()


def check_recipe(path: str, recipe: dict[str, Any]) -> list[dict[str, Any]]:
    """Validate a recipe and report which blocks or slides each rule selects.

    A rule that selects nothing has empty block_ids and pages; fix its target.

    Args:
        path: Local path to the .pptx file.
        recipe: The recipe to check (see recipe_schema).
    """

    try:
        parsed = Recipe.model_validate(recipe)
    except ValidationError as exc:
        raise ValueError(f"invalid recipe: {exc}") from exc
    with _open(path, None) as presentation:
        return [
            {
                "rule": match.rule,
                "target": match.target,
                "block_ids": list(match.block_ids),
                "pages": list(match.pages),
            }
            for match in explain_recipe(presentation.document, parsed)
        ]


def presentation_chunks(
    path: str, recipe: dict[str, Any] | None = None, max_chars: int | None = None
) -> list[dict[str, Any]]:
    """Slide-aligned chunks for retrieval, with citation metadata.

    Args:
        path: Local path to the .pptx file.
        recipe: Optional recipe applied first.
        max_chars: Split slides longer than this at block boundaries.
    """

    with _open(path, recipe) as presentation:
        return [
            {"text": chunk.text, "metadata": chunk.metadata}
            for chunk in slide_chunks(presentation.document, max_chars=max_chars)
        ]


# Image formats a vision model reads; others (EMF/WMF) need --vector-images png first.
_VIEWABLE = {"image/png", "image/jpeg", "image/gif", "image/webp"}
_MAX_IMAGE_BYTES = 5 * 1024 * 1024


def presentation_images(
    path: str, block_ids: list[str] | None = None, max_images: int = 8
) -> list[Any]:
    """The pictures that need a description, each after a line with its block id.

    Look at each picture and its slide, then describe it with a recipe rule such as
    {"target": "s3.b2", "description": "Bar chart: revenue doubled in 2024"}, or
    {"target": "s3.b2", "exclude": true} for decoration, and pass the recipe to
    convert_presentation or presentation_chunks.

    Args:
        path: Local path to the .pptx file.
        block_ids: Image block ids to show; by default those without a description.
        max_images: At most this many pictures per call.
    """

    from mcp.server.mcpserver import Image

    with _open(path, None, include_assets=True) as presentation:
        document = presentation.document
        wanted = (
            set(block_ids)
            if block_ids is not None
            else {item.target for item in review(document) if item.topic == "image-description"}
        )
        content: list[Any] = []
        skipped: list[str] = []
        for slide in document.slides:
            for block in slide.blocks:
                if not isinstance(block, ImageBlock) or block.id not in wanted:
                    continue
                asset = presentation.assets.descriptor(block.asset_id)
                if asset.media_type not in _VIEWABLE or asset.size > _MAX_IMAGE_BYTES:
                    skipped.append(f"{block.id} ({asset.media_type}, {asset.size} bytes)")
                    continue
                if len(content) // 2 >= max_images:
                    skipped.append(f"{block.id} (over max_images)")
                    continue
                content.append(f"{block.id} on slide {slide.page}: alt text {block.alt_text!r}")
                content.append(
                    Image(
                        data=presentation.assets.read_bytes(block.asset_id),
                        format=asset.media_type.split("/", 1)[1],
                    )
                )
        if skipped:
            content.append("Not shown: " + ", ".join(skipped))
        return content or ["No picture needs a description."]


def recipe_schema() -> dict[str, Any]:
    """The JSON Schema of a recipe, with the selector syntax and every rule action."""

    return recipe_json_schema()


def prepare_for_retrieval(path: str, purpose: str = "a search index") -> str:
    """Check, refine and chunk a deck for retrieval, ending with a reusable recipe."""

    return f"""Prepare the presentation at {path} for {purpose}.

1. Call audit_presentation. If it reports lost text, list those slides and texts in
   your final summary: the converter dropped them, and a recipe cannot bring them back.
2. Call review_presentation. For each item choose an option: the suggested one unless
   the slide, read with convert_presentation, shows otherwise.
3. Call presentation_images. Look at each picture: describe what it conveys in one or
   two sentences (chart type and trend, what a diagram or screenshot shows), or mark
   logos, backgrounds and other decoration as excluded.
4. Put every choice into one recipe (see recipe_schema) and call check_recipe; fix any
   rule that selects nothing.
5. Call presentation_chunks with the recipe.

Finish with the number of chunks, what you changed and why, and the recipe as JSON so
it can be saved next to the deck and reused (pptx2markdown deck.pptx --recipe ...)."""


def apply_request(path: str, request: str) -> str:
    """Turn a plain-language request about a deck's conversion into a checked recipe."""

    return f"""Convert the presentation at {path} as requested: {request}

Read it with convert_presentation; each block comes after a comment with its id and
kind. Read recipe_schema for what rules can do. Write the smallest recipe that does
what was asked: select blocks by id or by slide and kind, never edit wording beyond
what was asked. Call check_recipe and fix rules that select nothing, then show the
result with convert_presentation and the recipe as JSON."""


PROMPTS = (prepare_for_retrieval, apply_request)

TOOLS = (
    convert_presentation,
    review_presentation,
    audit_presentation,
    check_recipe,
    presentation_chunks,
    presentation_images,
    recipe_schema,
)


def build_server():
    """The MCP server with every tool registered."""

    try:
        from mcp.server.mcpserver import MCPServer
    except ImportError as exc:  # pragma: no cover - depends on the environment
        raise SystemExit(
            'the MCP server needs the mcp package: pip install "pptx2markdown[mcp]"'
        ) from exc
    server = MCPServer("pptx2markdown", instructions=INSTRUCTIONS)
    for tool in TOOLS:
        # Pictures go out as image content, which has no JSON form.
        server.add_tool(tool, structured_output=False if tool is presentation_images else None)
    for prompt in PROMPTS:
        server.prompt()(prompt)
    return server


def main() -> None:
    build_server().run("stdio")


if __name__ == "__main__":
    main()
