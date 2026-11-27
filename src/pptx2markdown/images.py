"""Image descriptions from any model: a describer callable in, a recipe out.

``describe_images_with`` hands each picture that needs a description (by default
those ``review`` asks about) to a function you supply, with the slide's Markdown
for context, and turns the answers into recipe rules. The function can call a
local vision model, an OCR engine or an internal service; nothing leaves the
machine unless it sends it.
"""

from __future__ import annotations

from collections.abc import Callable, Collection
from dataclasses import dataclass
from typing import TYPE_CHECKING

from .model import ImageBlock
from .recipe import Recipe, Rule
from .render import render_slide_markdown
from .review import review

if TYPE_CHECKING:
    from .presentation import ParsedPresentation


@dataclass(frozen=True)
class ImageRequest:
    """One picture to describe, with what the slide says around it."""

    block_id: str
    page: int
    data: bytes
    media_type: str
    alt_text: str
    slide_markdown: str


@dataclass(frozen=True)
class ImageDescription:
    """What a describer found: a description, or that the picture is decoration."""

    description: str | None = None
    decorative: bool = False


Describer = Callable[[ImageRequest], "ImageDescription | None"]


def describe_images_with(
    presentation: "ParsedPresentation",
    describe: Describer,
    *,
    block_ids: Collection[str] | None = None,
    media_types: Collection[str] | None = None,
    max_bytes: int | None = None,
) -> Recipe:
    """A recipe with the describer's answer for each picture that needs one.

    ``block_ids`` defaults to the images ``review`` flags as undescribed. Pictures
    whose media type is not in ``media_types`` (all, by default), larger than
    ``max_bytes`` or without a payload are skipped; so are those the describer
    returns None for. Decorative pictures become ``exclude`` rules, the rest
    ``description`` rules.
    """

    document = presentation.document
    wanted = (
        set(block_ids)
        if block_ids is not None
        else {item.target for item in review(document) if item.topic == "image-description"}
    )
    rules: list[Rule] = []
    for slide in document.slides:
        for block in slide.blocks:
            if not isinstance(block, ImageBlock) or block.id not in wanted:
                continue
            descriptor = presentation.assets.descriptor(block.asset_id)
            if (
                not descriptor.payload_available
                or (media_types is not None and descriptor.media_type not in media_types)
                or (max_bytes is not None and descriptor.size > max_bytes)
            ):
                continue
            result = describe(
                ImageRequest(
                    block_id=block.id,
                    page=slide.page,
                    data=presentation.assets.read_bytes(block.asset_id),
                    media_type=descriptor.media_type,
                    alt_text=block.alt_text,
                    slide_markdown=render_slide_markdown(slide, asset_mode="omit"),
                )
            )
            if result is None:
                continue
            if result.decorative:
                rules.append(Rule(target=block.id, exclude=True, reason="decorative image"))
            elif result.description and result.description.strip():
                rules.append(
                    Rule(
                        target=block.id,
                        description=result.description.strip(),
                        reason="described from the image",
                    )
                )
    return Recipe(rules=tuple(rules))
