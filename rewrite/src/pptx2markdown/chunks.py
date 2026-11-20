"""Slide-aligned chunks for retrieval.

A slide is the unit a presentation is written in, so it is the natural chunk: its
title gives the context, and the page number lets an answer cite where it came
from. Generic converters join a deck into one document and lose both.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Literal

from .inline import escape_text
from .model import HeadingBlock, ImageBlock, PresentationDocument, SlideDocument
from .render import AssetResolver, render_block_markdown


@dataclass(frozen=True)
class Chunk:
    """Markdown for one slide (or part of one) with citation metadata."""

    text: str
    metadata: dict[str, Any] = field(default_factory=dict)


# Names PowerPoint gives pictures by default; they say nothing about the content.
_DEFAULT_PICTURE_NAME = re.compile(
    r"^(?:picture|image|graphic|grafik|imagen|그림|이미지|図|图片)\s*\d*$", re.IGNORECASE
)


def _image_text(block: ImageBlock) -> str | None:
    """An image without a location still carries meaning through its description."""

    text = (block.description or block.alt_text or "").strip()
    if not text or _DEFAULT_PICTURE_NAME.match(text):
        return None
    return f"[image: {escape_text(text)}]"


def _title(slide: SlideDocument) -> str | None:
    headings = [block for block in slide.blocks if isinstance(block, HeadingBlock)]
    if not headings:
        return None
    return min(headings, key=lambda block: block.level).text


def slide_chunks(
    document: PresentationDocument,
    *,
    asset_resolver: AssetResolver | None = None,
    asset_mode: Literal["error", "omit"] = "omit",
    include_notes: bool = True,
    max_chars: int | None = None,
    context: bool = True,
) -> list[Chunk]:
    """One chunk per slide, split at block boundaries when longer than ``max_chars``.

    With ``context`` every chunk starts with the deck and slide title, so a chunk
    retrieved on its own still says what it is about. Without an ``asset_resolver``
    an image is written as ``[image: description]`` when it has a meaningful
    description or alt text. A block longer than ``max_chars`` is kept whole rather
    than cut mid-table.
    """

    if max_chars is not None and max_chars <= 0:
        raise ValueError("max_chars must be positive")
    deck_title = next((title for slide in document.slides if (title := _title(slide))), None)
    chunks: list[Chunk] = []
    for slide in document.slides:
        title = _title(slide)
        pieces: list[tuple[str, str | None, str]] = []
        for block in slide.blocks:
            rendered = render_block_markdown(
                block, asset_resolver=asset_resolver, asset_mode=asset_mode
            )
            if not rendered and isinstance(block, ImageBlock):
                rendered = _image_text(block)
            if rendered:
                pieces.append((rendered, block.id, block.kind))
        if include_notes and slide.notes and slide.notes.strip():
            pieces.append((f"Speaker notes: {escape_text(slide.notes.strip())}", None, "notes"))
        if not pieces:
            continue

        header = ""
        if context:
            names = [name for name in (deck_title, title) if name]
            label = " — ".join(dict.fromkeys(names))
            header = f"<!-- {document.source.name}, slide {slide.page}"
            header += f": {label}" if label else ""
            # A "-->" inside a title would end the comment early.
            header = header.replace("-->", "- ->") + " -->\n\n"

        groups: list[list[tuple[str, str | None, str]]] = [[]]
        size = len(header)
        for piece in pieces:
            added = len(piece[0]) + 2
            if max_chars is not None and groups[-1] and size + added > max_chars:
                groups.append([])
                size = len(header)
            groups[-1].append(piece)
            size += added

        for part, group in enumerate(groups, 1):
            chunks.append(
                Chunk(
                    text=header + "\n\n".join(text for text, _, _ in group),
                    metadata={
                        "source": document.source.name,
                        "source_sha256": document.source.sha256,
                        "page": slide.page,
                        "part": part,
                        "parts": len(groups),
                        "title": title,
                        "deck_title": deck_title,
                        "hidden": slide.hidden,
                        "block_ids": [block_id for _, block_id, _ in group if block_id],
                        "kinds": sorted({kind for _, _, kind in group}),
                    },
                )
            )
    return chunks
