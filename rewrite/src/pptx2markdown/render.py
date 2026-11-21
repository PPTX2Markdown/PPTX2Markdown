"""Pure Markdown and JSON renderers for the canonical document."""

from __future__ import annotations

import json
import re
from collections.abc import Callable
from typing import Literal

from .errors import AssetResolutionError
from .inline import escape_text
from .model import (
    AttachmentBlock,
    ChartBlock,
    ContentBlock,
    HeadingBlock,
    ImageBlock,
    ListBlock,
    MathBlock,
    PresentationDocument,
    SlideDocument,
    SmartArtBlock,
    TableBlock,
    TextBlock,
    UnsupportedBlock,
)

AssetResolver = Callable[[str], str]


def _escape_alt_text(text: str) -> str:
    return str(text).replace("\\", "\\\\").replace("[", r"\[").replace("]", r"\]")


def _destination(value: str) -> str:
    return f"<{value}>" if any(char.isspace() or char in "()" for char in value) else value


def _resolve_asset(
    asset_id: str,
    resolver: AssetResolver | None,
    asset_mode: Literal["error", "omit"],
) -> str | None:
    if resolver is None:
        if asset_mode == "omit":
            return None
        raise AssetResolutionError(
            f"asset {asset_id} requires a URL resolver or asset_mode='omit'"
        )
    resolved = str(resolver(asset_id)).strip()
    if not resolved:
        if asset_mode == "omit":
            return None
        raise AssetResolutionError(f"resolver returned an empty location for {asset_id}")
    return resolved.replace("\\", "/")


def _resolve_embedded_assets(
    markdown: str,
    asset_ids: tuple[str, ...],
    resolver: AssetResolver | None,
    asset_mode: Literal["error", "omit"],
) -> str:
    rendered = markdown
    for asset_id in asset_ids:
        location = _resolve_asset(asset_id, resolver, asset_mode)
        token = f"asset://{asset_id}"
        if location is None:
            # An omitted picture leaves no empty "![alt]()" behind.
            image = re.compile(r"!\[(?:\\.|[^\]\\])*\]\(" + re.escape(token) + r"\)")
            rendered = image.sub("", rendered)
        rendered = rendered.replace(token, location or "")
    rendered = re.sub(r"(?:<br>){2,}", "<br>", rendered)
    rendered = re.sub(r"(\|\s*)<br>|<br>(\s*\|)", r"\1\2", rendered)
    return rendered.strip()


def render_block_markdown(
    block: ContentBlock,
    *,
    asset_resolver: AssetResolver | None,
    asset_mode: Literal["error", "omit"],
) -> str | None:
    """Render one typed block without filesystem access."""

    if isinstance(block, TextBlock):
        return block.text.strip()
    if isinstance(block, HeadingBlock):
        return f"{'#' * block.level} {block.text.strip()}"
    if isinstance(block, ListBlock):
        lines: list[str] = []
        counters: dict[int, tuple[bool, int]] = {}
        # Open ancestors as (level, column where their text starts): a nested item must
        # be indented past its parent's marker ("1. " is wider than "- "), or
        # CommonMark does not nest it.
        parents: list[tuple[int, int]] = []
        top = min(item.level for item in block.items)
        for item in block.items:
            ordered = block.ordered if item.ordered is None else item.ordered
            while parents and parents[-1][0] >= item.level:
                parents.pop()
            column = parents[-1][1] if parents else 0
            counters = {level: c for level, c in counters.items() if level <= item.level}
            if ordered:
                kind, count = counters.get(item.level, (True, 0))
                if not kind:
                    count = 0
                if count == 0 and item.level == top:
                    count = block.start - 1
                counters[item.level] = (True, count + 1)
                marker = f"{count + 1}."
            else:
                counters[item.level] = (False, 0)
                marker = "-"
            lines.append(f"{' ' * column}{marker} {item.text.strip()}")
            parents.append((item.level, column + (4 if ordered else 2)))
        return "\n".join(lines)
    if isinstance(block, MathBlock):
        return f"$$\n{block.latex.strip()}\n$$"
    if isinstance(block, ImageBlock):
        location = _resolve_asset(block.asset_id, asset_resolver, asset_mode)
        if location is None:
            return None
        alt_text = block.description or block.alt_text or "image"
        return f"![{_escape_alt_text(alt_text)}]({_destination(location)})"
    if isinstance(block, AttachmentBlock):
        location = _resolve_asset(block.asset_id, asset_resolver, asset_mode)
        if location is None:
            return None
        label = block.label.replace("\\", "\\\\").replace("[", r"\[").replace("]", r"\]")
        return f"[attachment: {label}]({_destination(location)})"
    if isinstance(block, (ChartBlock, SmartArtBlock, TableBlock)):
        return _resolve_embedded_assets(
            block.markdown,
            block.asset_ids,
            asset_resolver,
            asset_mode,
        )
    if isinstance(block, UnsupportedBlock):
        return f"[unsupported: {block.message}]"
    raise TypeError(f"unknown content block: {type(block).__name__}")


def render_slide_markdown(
    slide: SlideDocument,
    *,
    asset_resolver: AssetResolver | None = None,
    asset_mode: Literal["error", "omit"] = "error",
    block_ids: bool = False,
    page_marker: bool = True,
) -> str:
    """Render one slide; ``block_ids`` prefixes each block with ``<!-- s3.b2 -->``."""

    parts = [f"[Page_{slide.page}]"] if page_marker else []
    if slide.hidden:
        parts.append("<!-- hidden: true -->")
    if block_ids and slide.section:
        # The annotated view names the section, so a reader can select it in a recipe.
        parts.append(f"<!-- section: {slide.section.replace('--', '- -')} -->")
    for block in slide.blocks:
        rendered = render_block_markdown(
            block,
            asset_resolver=asset_resolver,
            asset_mode=asset_mode,
        )
        if rendered:
            if block_ids and block.id:
                rendered = f"<!-- {block.id} {block.kind} -->\n{rendered}"
            parts.append(rendered)
    if slide.notes and slide.notes.strip():
        parts.extend(["[Speaker_Notes]", escape_text(slide.notes.strip())])
    return "\n\n".join(parts)


def render_presentation_markdown(
    document: PresentationDocument,
    *,
    asset_resolver: AssetResolver | None = None,
    asset_mode: Literal["error", "omit"] = "error",
    block_ids: bool = False,
) -> str:
    """Render deterministic Markdown from one canonical document.

    ``block_ids`` annotates every block with its id and kind in an HTML comment, so a
    reviewer or a language model can point recipe rules at exact blocks.
    """

    slides = [
        render_slide_markdown(
            slide,
            asset_resolver=asset_resolver,
            asset_mode=asset_mode,
            block_ids=block_ids,
        )
        for slide in document.slides
    ]
    merged = "\n\n".join(slides).strip()
    return f"{merged}\n" if merged else ""


def render_presentation_json(document: PresentationDocument) -> str:
    """Render canonical JSON without paths, timestamps or asset payloads."""

    value = document.model_dump(mode="json")
    return json.dumps(value, ensure_ascii=False, indent=2, sort_keys=True) + "\n"
