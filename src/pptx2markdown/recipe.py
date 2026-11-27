"""Declarative conversion recipes applied to a parsed document.

A recipe never touches the parser. It is data: document-wide settings plus rules
that each select blocks or slides and say what to do with them. Applying a recipe
is a pure, deterministic transformation of the canonical document, so the same
recipe always gives the same output and can be stored, reviewed and shared. This
is the surface a person, a UI or a language model edits.

Selectors
---------
``s3.b2``                 one block, by id
``slide[3]``              slide 3 (``slide[2-5]``, ``slide[12-]``, ``slide[*]``,
                          ``slide[hidden]``)
``table``                 every table (any block kind, or ``*`` for every block)
``table[0]``              the first table on each slide (``[-1]`` is the last)
``slide[3]/table[0]``     the first table on slide 3
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

from .model import (
    BLOCK_ID_PATTERN,
    ContentBlock,
    HeadingBlock,
    ImageBlock,
    ListBlock,
    MergedCells,
    PresentationDocument,
    SlideDocument,
    TableBlock,
    TextBlock,
)
from .tables import rows_markdown

BlockKind = Literal[
    "text",
    "heading",
    "list",
    "math",
    "image",
    "chart",
    "smartart",
    "table",
    "attachment",
    "unsupported",
]
_KINDS = set(BlockKind.__args__)  # type: ignore[attr-defined]

SELECTOR_HELP = (
    "Block id like 's3.b2'; or '[slide[RANGE]/]KIND[[INDEX]]' where RANGE is N, N-M, N-, * "
    "or hidden, KIND is a block kind or *, and INDEX counts that kind on each slide from 0 "
    "(-1 is the last); or 'slide[RANGE]' alone for whole slides. 'section[NAME]' in place "
    "of 'slide[RANGE]' picks the slides of a PowerPoint section (case-insensitive)."
)

_ID = re.compile(BLOCK_ID_PATTERN)
_SLIDE = re.compile(r"^slide\[(?P<range>[^\]]+)\]$")
_SECTION = re.compile(r"^section\[(?P<name>[^\]]+)\]$")
_BLOCK = re.compile(r"^(?P<kind>[a-z]+|\*)(?:\[(?P<index>-?\d+)\])?$")
_WHOLE_EMPHASIS = re.compile(r"^(?P<mark>\*{1,3})(?P<text>[^*]+)(?P=mark)$")
_RANGE = re.compile(r"^(?P<start>\d+)(?:(?P<dash>-)(?P<end>\d+)?)?$")


@dataclass(frozen=True)
class Selector:
    """A parsed selector; see the module docstring for the syntax."""

    block_id: str | None = None
    first_page: int | None = None
    last_page: int | None = None
    hidden_only: bool = False
    kind: str | None = None
    index: int | None = None
    slides_only: bool = False
    section: str | None = None

    @classmethod
    def parse(cls, text: str) -> "Selector":
        value = text.strip()
        if _ID.match(value):
            return cls(block_id=value)
        slide_part, _, block_part = value.partition("/")
        section_match = _SECTION.match(slide_part)
        if section_match is not None:
            selector = cls(section=section_match.group("name").strip().casefold())
            if not block_part:
                return Selector(**{**selector.__dict__, "slides_only": True})
            return cls._with_block(selector, block_part, text)
        slide_match = _SLIDE.match(slide_part)
        if slide_match is None:
            if "/" in value:
                raise ValueError(f"invalid slide selector in {text!r}")
            return cls._with_block(cls(), value, text)
        selector = cls._slides(slide_match.group("range"), text)
        if not block_part:
            return Selector(**{**selector.__dict__, "slides_only": True})
        return cls._with_block(selector, block_part, text)

    @staticmethod
    def _slides(spec: str, text: str) -> "Selector":
        spec = spec.strip()
        if spec == "*":
            return Selector()
        if spec == "hidden":
            return Selector(hidden_only=True)
        match = _RANGE.match(spec)
        if match is None:
            raise ValueError(f"invalid slide range in {text!r}")
        start = int(match.group("start"))
        if start < 1:
            raise ValueError(f"slide numbers start at 1 in {text!r}")
        end = match.group("end")
        if match.group("dash") is None:
            last: int | None = start
        else:
            last = int(end) if end is not None else None
        if last is not None and last < start:
            raise ValueError(f"empty slide range in {text!r}")
        return Selector(first_page=start, last_page=last)

    @staticmethod
    def _with_block(selector: "Selector", spec: str, text: str) -> "Selector":
        match = _BLOCK.match(spec.strip())
        if match is None or (match.group("kind") != "*" and match.group("kind") not in _KINDS):
            raise ValueError(f"invalid block selector in {text!r}")
        index = match.group("index")
        return Selector(
            **{
                **selector.__dict__,
                "kind": None if match.group("kind") == "*" else match.group("kind"),
                "index": int(index) if index is not None else None,
            }
        )

    def matches_slide(self, slide: SlideDocument) -> bool:
        if self.block_id is not None:
            return self.block_id.startswith(f"s{slide.page}.")
        if self.hidden_only and not slide.hidden:
            return False
        if self.section is not None and (slide.section or "").casefold() != self.section:
            return False
        if self.first_page is not None and slide.page < self.first_page:
            return False
        return self.last_page is None or slide.page <= self.last_page

    def blocks(self, slide: SlideDocument) -> list[ContentBlock]:
        """The blocks of ``slide`` this selector picks (all of them for a slide selector)."""

        if not self.matches_slide(slide):
            return []
        if self.block_id is not None:
            return [block for block in slide.blocks if block.id == self.block_id]
        candidates = [
            block for block in slide.blocks if self.kind is None or block.kind == self.kind
        ]
        if self.index is None:
            return candidates
        try:
            return [candidates[self.index]]
        except IndexError:
            return []


class Rule(BaseModel):
    """What to do with the blocks or slides one selector picks.

    Unset fields leave the block as it is; later rules override earlier ones.
    """

    model_config = ConfigDict(extra="forbid", frozen=True)

    target: str = Field(min_length=1, description=SELECTOR_HELP)
    exclude: bool | None = Field(
        default=None, description="Drop the selected blocks, or whole slides for a slide selector."
    )
    merged_cells: MergedCells | None = Field(
        default=None,
        description="Tables: fill repeats merged values, header only in the header, blank never.",
    )
    header_rows: int | None = Field(
        default=None, ge=0, le=10, description="Tables: how many rows form the header."
    )
    caption_rows: int | None = Field(
        default=None,
        ge=0,
        le=5,
        description="Tables: how many leading rows are a title above the table, not data.",
    )
    heading_level: int | None = Field(
        default=None,
        ge=1,
        le=6,
        description="Headings: set the level. Text: turn the block into a heading.",
    )
    as_text: bool | None = Field(
        default=None, description="Headings and lists: render as plain text lines instead."
    )
    description: str | None = Field(
        default=None, min_length=1, description="Images: description used as alt text."
    )
    notes: bool | None = Field(
        default=None, description="Slide selectors: keep (true) or drop (false) speaker notes."
    )
    order: tuple[str, ...] | None = Field(
        default=None,
        description="Slide selectors: read the slide's blocks in this order of block ids; "
        "blocks left out follow in their current order.",
    )
    reason: str | None = Field(
        default=None, description="Why the rule exists; kept for review, never applied."
    )

    @field_validator("target")
    @classmethod
    def validate_target(cls, value: str) -> str:
        Selector.parse(value)
        return value.strip()

    @model_validator(mode="after")
    def validate_order(self) -> "Rule":
        if self.order is not None and not Selector.parse(self.target).slides_only:
            raise ValueError("order needs a slide selector such as 'slide[4]'")
        return self


class Recipe(BaseModel):
    """Document-wide settings and targeted rules, applied in order."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    version: Literal["1"] = "1"
    merged_cells: MergedCells | None = Field(
        default=None, description="Default merged-cell policy for every table."
    )
    include_notes: bool = True
    include_hidden: bool = True
    template_pictures: bool = Field(
        default=True,
        description="false drops pictures drawn from the slide layout or master (logos, artwork).",
    )
    headings: Literal["auto", "placeholder"] = Field(
        default="auto",
        description="placeholder keeps only headings from title placeholders; others become text.",
    )
    rules: tuple[Rule, ...] = ()


@dataclass(frozen=True)
class RuleMatch:
    """The blocks and slides a rule selects in one document."""

    rule: int
    target: str
    block_ids: tuple[str, ...]
    pages: tuple[int, ...]


def explain_recipe(document: PresentationDocument, recipe: Recipe) -> tuple[RuleMatch, ...]:
    """What every rule would select, so a reviewer or agent can check selectors first."""

    matches: list[RuleMatch] = []
    for index, rule in enumerate(recipe.rules):
        selector = Selector.parse(rule.target)
        block_ids: list[str] = []
        pages: list[int] = []
        for slide in document.slides:
            if selector.slides_only and selector.matches_slide(slide):
                pages.append(slide.page)
            block_ids.extend(block.id or "" for block in selector.blocks(slide))
        matches.append(RuleMatch(index, rule.target, tuple(block_ids), tuple(pages)))
    return tuple(matches)


def _text_lines(block: ContentBlock) -> str | None:
    if isinstance(block, HeadingBlock):
        return block.text
    if isinstance(block, ListBlock):
        return "\n".join(item.text for item in block.items)
    if isinstance(block, TextBlock):
        return block.text
    return None


def _apply_block(block: ContentBlock, rule: Rule, default_merged: MergedCells | None):
    provenance = {
        "id": block.id,
        "shape_id": block.shape_id,
        "bbox": block.bbox,
        "source_part": block.source_part,
    }
    if isinstance(block, TableBlock) and block.rows:
        header_rows = rule.header_rows if rule.header_rows is not None else block.header_rows
        captions = rule.caption_rows if rule.caption_rows is not None else block.caption_rows
        merged = rule.merged_cells or default_merged or block.merged_cells
        if (header_rows, captions, merged) != (
            block.header_rows,
            block.caption_rows,
            block.merged_cells,
        ):
            return block.model_copy(
                update={
                    "header_rows": header_rows,
                    "caption_rows": captions,
                    "merged_cells": merged,
                    "markdown": rows_markdown(
                        block.rows,
                        header_rows=header_rows,
                        merged_cells=merged,
                        caption_rows=captions,
                    ),
                }
            )
    if rule.as_text and isinstance(block, (HeadingBlock, ListBlock)):
        return TextBlock(text=_text_lines(block) or "", **provenance)
    if rule.heading_level is not None:
        if isinstance(block, HeadingBlock):
            return block.model_copy(update={"level": rule.heading_level})
        if isinstance(block, TextBlock):
            text = " ".join(block.text.split("\n"))
            # A line set entirely in bold or italic is plain once it is a heading.
            text = _WHOLE_EMPHASIS.sub(r"\g<text>", text)
            return HeadingBlock(text=text, level=rule.heading_level, **provenance)
    if rule.description is not None and isinstance(block, ImageBlock):
        return block.model_copy(update={"description": rule.description})
    return block


def apply_recipe(document: PresentationDocument, recipe: Recipe) -> PresentationDocument:
    """Return the document with the recipe applied; the input is not modified."""

    selectors = [Selector.parse(rule.target) for rule in recipe.rules]
    slides: list[SlideDocument] = []
    for slide in document.slides:
        if slide.hidden and not recipe.include_hidden:
            continue
        keep_slide = True
        keep_notes = recipe.include_notes
        order: tuple[str, ...] | None = None
        blocks: list[ContentBlock] = []
        if recipe.merged_cells is not None:
            neutral = Rule(target="*")
            source = [_apply_block(block, neutral, recipe.merged_cells) for block in slide.blocks]
        else:
            source = list(slide.blocks)
        if not recipe.template_pictures:
            # Dropped blocks become None so `source` stays aligned with slide.blocks.
            source = [
                None if isinstance(block, ImageBlock) and block.source_part != "slide" else block
                for block in source
            ]
        if recipe.headings == "placeholder":
            source = [
                TextBlock(
                    text=block.text,
                    id=block.id,
                    shape_id=block.shape_id,
                    bbox=block.bbox,
                    source_part=block.source_part,
                )
                if isinstance(block, HeadingBlock) and block.source not in (None, "placeholder")
                else block
                for block in source
            ]
        for rule, selector in zip(recipe.rules, selectors):
            if selector.slides_only and selector.matches_slide(slide):
                if rule.exclude:
                    keep_slide = False
                if rule.notes is not None:
                    keep_notes = rule.notes
                if rule.order is not None:
                    order = rule.order
        if not keep_slide:
            continue
        # Selectors always see the parsed slide, so a rule's targets do not depend on
        # what earlier rules changed.
        chosen = [{id(block) for block in selector.blocks(slide)} for selector in selectors]
        for original, block in zip(slide.blocks, source):
            if block is None:
                continue
            kept = True
            for rule, picked in zip(recipe.rules, chosen):
                if id(original) not in picked:
                    continue
                if rule.exclude:
                    kept = False
                    break
                block = _apply_block(block, rule, None)
            if kept:
                blocks.append(block)
        if order is not None:
            rank = {block_id: index for index, block_id in enumerate(order)}
            blocks.sort(key=lambda block: rank.get(block.id or "", len(rank)))
        slides.append(
            slide.model_copy(
                update={"blocks": tuple(blocks), "notes": slide.notes if keep_notes else None}
            )
        )
    return document.model_copy(update={"slides": tuple(slides)})


def recipe_json_schema() -> dict:
    """JSON Schema of a recipe, for forms and for structured model output."""

    return Recipe.model_json_schema()


def merge_recipes(*recipes: Recipe) -> Recipe:
    """One recipe doing what all of them do, later ones winning on conflicts.

    Rules are concatenated in order; a later ``merged_cells`` replaces an earlier
    one, and notes or hidden slides are kept only if every recipe keeps them.
    """

    merged = Recipe()
    for recipe in recipes:
        merged = Recipe(
            merged_cells=recipe.merged_cells or merged.merged_cells,
            headings="placeholder"
            if "placeholder" in (merged.headings, recipe.headings)
            else "auto",
            include_notes=merged.include_notes and recipe.include_notes,
            include_hidden=merged.include_hidden and recipe.include_hidden,
            template_pictures=merged.template_pictures and recipe.template_pictures,
            rules=merged.rules + recipe.rules,
        )
    return merged
