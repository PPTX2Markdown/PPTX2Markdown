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
from .chunks import slide_chunks
from .presentation import ParsedPresentation
from .recipe import Recipe, explain_recipe, recipe_json_schema
from .review import review, review_data

INSTRUCTIONS = """\
Converts PowerPoint (.pptx) files to Markdown for reading and retrieval.

Read a deck with convert_presentation: every block is preceded by a comment such as
`<!-- s3.b2 table -->` giving its id and kind. review_presentation lists the
conversion decisions worth checking; each option carries a ready-made recipe rule.
To change the output (drop slides, set table header rows, turn text into a
heading, describe an image), write a recipe (see recipe_schema), check it with
check_recipe so every rule selects something, then pass it to convert_presentation
or presentation_chunks. Recipes never change the source file."""


def _open(path: str, recipe: dict[str, Any] | None) -> ParsedPresentation:
    source = Path(path).expanduser()
    if source.suffix.casefold() != ".pptx":
        raise ValueError(f"not a .pptx file: {path}")
    if not source.is_file():
        raise ValueError(f"file not found: {path}")
    try:
        parsed_recipe = Recipe.model_validate(recipe) if recipe else None
    except ValidationError as exc:
        raise ValueError(f"invalid recipe: {exc}") from exc
    return parse(source, include_assets=False, strict=False, recipe=parsed_recipe)


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


def recipe_schema() -> dict[str, Any]:
    """The JSON Schema of a recipe, with the selector syntax and every rule action."""

    return recipe_json_schema()


TOOLS = (
    convert_presentation,
    review_presentation,
    check_recipe,
    presentation_chunks,
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
        server.add_tool(tool)
    return server


def main() -> None:
    build_server().run("stdio")


if __name__ == "__main__":
    main()
