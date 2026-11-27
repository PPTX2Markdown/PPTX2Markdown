"""Where the deterministic converter is least sure, as questions with recipe answers.

The parser always decides; ``review`` points at the decisions most worth a second
look and offers each alternative as a ready-made recipe rule. A person, a UI or a
language model then only has to pick options, and the choice stays a recipe.

Every item is a decision point in one shape whatever it is about (see
docs/agent_contract.md): a stable ``id``, machine-readable ``reasons``, the page
and what to ``look`` at to decide (the slide's text, its rendering, or a picture),
an ``impact`` for budgeting, and options with stable ids. ``recipe_from_answers``
turns ``{item id: option id}`` (or free text where ``free_text`` allows it) into a
recipe, so an agent never has to copy rules by hand.
"""

from __future__ import annotations

import re
from collections.abc import Mapping, Sequence
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field

from .model import (
    ContentBlock,
    HeadingBlock,
    ImageBlock,
    ListBlock,
    PresentationDocument,
    SlideDocument,
    TableBlock,
    TableCell,
    TextBlock,
)
from .reading_order import CONTAINER_ACROSS, FRAGILE_GAP, NO_CUT, OVERLAPPING, TWO_WAY_CUT
from .recipe import Recipe, Rule

# Bumped when a field changes meaning or goes away; new fields keep the version.
REVIEW_SCHEMA_VERSION = 1
KEEP = "keep"

Topic = Literal[
    "reading-order", "heading", "subheading", "table-caption", "table-header-rows",
    "image-description",
]  # fmt: skip
# What a judge needs to decide: the slide's text (Markdown) is enough, the slide
# as rendered, or the picture itself.
Look = Literal["text", "slide", "image"]
# "high": a wrong answer corrupts content for retrieval (mixed-up order, table
# cells under the wrong header, a picture nobody can find); "low": formatting.
Impact = Literal["high", "low"]

_ORDER_REASONS = {
    TWO_WAY_CUT: "two-way-cut",
    NO_CUT: "no-cut",
    OVERLAPPING: "overlapping-boxes",
    FRAGILE_GAP: "fragile-gap",
    CONTAINER_ACROSS: "container-across-content",
}
_BOLD_LINE = re.compile(r"^\*\*(?P<text>[^*\n]+)\*\*$")
_DEFAULT_PICTURE_NAME = re.compile(
    r"^(?:picture|image|graphic|grafik|imagen|그림|이미지|図|图片)\s*\d*$", re.IGNORECASE
)


class Option(BaseModel):
    """One answer to a review question; ``rule`` is None for "keep as converted"."""

    model_config = ConfigDict(frozen=True)

    id: str
    label: str
    rule: Rule | None = None


class ReviewItem(BaseModel):
    """A converter decision worth checking, with its alternatives."""

    model_config = ConfigDict(frozen=True)

    id: str = Field(description="Stable within a conversion: '<topic>@<target>'.")
    topic: Topic
    target: str = Field(description="Recipe selector the options' rules apply to.")
    page: int = Field(description="Slide to look at.")
    question: str
    reasons: tuple[str, ...] = Field(description="Machine-readable codes for why it is asked.")
    look: Look
    impact: Impact
    options: tuple[Option, ...]
    suggested: str = Field(description="Id of the option to take when nothing settles it.")
    free_text: Literal["description"] | None = Field(
        default=None,
        description="Rule field a free-text answer fills, for questions not limited to options.",
    )
    asset_id: str | None = Field(default=None, description="The picture, when look is 'image'.")
    evidence: dict[str, Any] = Field(default_factory=dict)

    def option(self, option_id: str) -> Option | None:
        return next((option for option in self.options if option.id == option_id), None)


def _item(topic: Topic, target: str, page: int, **fields: Any) -> ReviewItem:
    return ReviewItem(id=f"{topic}@{target}", topic=topic, target=target, page=page, **fields)


def _slug(label: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", label.casefold()).strip("-")


def _text(cell: TableCell, rows: tuple[tuple[TableCell, ...], ...]) -> str:
    if cell.merged_from is not None:
        row, column = cell.merged_from
        cell = rows[row][column]
    return cell.text.replace("*", "").replace("\\", "").strip()


def _is_number(text: str) -> bool:
    """Values such as "1,200", "36 / -", "+4.2%" or "2024": digits and symbols only."""

    compact = "".join(text.split())
    return any(char.isdigit() for char in compact) and all(
        char.isdigit() or char in ".,%+-−/():$€£¥₩~" for char in compact
    )


def _label_like(cells: list[str]) -> bool:
    filled = [text for text in cells if text]
    return bool(filled) and all(not _is_number(text) and len(text) <= 40 for text in filled)


def _table_items(block: TableBlock, page: int) -> list[ReviewItem]:
    rows = block.rows
    if not rows or block.id is None:
        return []
    items: list[ReviewItem] = []
    start = block.caption_rows
    if block.caption_rows:
        title = rows[0][0].text
        items.append(
            _item(
                "table-caption",
                block.id,
                page,
                question=f"Is the merged first row {title!r} a title above the table?",
                reasons=("merged-first-row",),
                look="text",
                impact="low",
                options=(
                    Option(id=KEEP, label="Yes, a title above the table"),
                    Option(
                        id="header",
                        label="No, it is the header",
                        rule=Rule(
                            target=block.id, caption_rows=0, reason="first row is the header"
                        ),
                    ),
                ),
                suggested=KEEP,
                evidence={"title": title},
            )
        )
    header = rows[start] if start < len(rows) else ()
    second = rows[start + 1] if start + 1 < len(rows) else ()
    header_texts = [_text(cell, rows) for cell in header]
    second_texts = [_text(cell, rows) for cell in second]
    spans_columns = any(cell.col_span > 1 for cell in header if cell.merged_from is None)
    body = rows[start + 2 :]
    body_has_numbers = any(_is_number(_text(cell, rows)) for row in body for cell in row)
    if spans_columns and second and _label_like(second_texts) and body_has_numbers:
        items.append(
            _item(
                "table-header-rows",
                block.id,
                page,
                question="The header has merged groups; is the next row part of the header?",
                reasons=("grouped-header",),
                look="text",
                impact="high",
                options=(
                    Option(id=KEEP, label="One header row"),
                    Option(
                        id="two-rows",
                        label="Two header rows",
                        rule=Rule(target=block.id, header_rows=2, reason="grouped two-row header"),
                    ),
                ),
                suggested="two-rows",
                evidence={
                    "header": header_texts,
                    "next_row": second_texts,
                    # A header cell merged down into the next row makes that row header too.
                    "header_spans_rows": any(
                        cell.row_span > 1 for cell in header if cell.merged_from is None
                    ),
                },
            )
        )
    elif (
        len(header_texts) > 1
        and all(_is_number(text) for text in header_texts if text)
        and any(header_texts)
    ):
        items.append(
            _item(
                "table-header-rows",
                block.id,
                page,
                question="The first row is all numbers; does the table have a header?",
                reasons=("numeric-first-row",),
                look="text",
                impact="high",
                options=(
                    Option(id=KEEP, label="The first row is the header"),
                    Option(
                        id="no-header",
                        label="No header row",
                        rule=Rule(target=block.id, header_rows=0, reason="first row is data"),
                    ),
                ),
                suggested="no-header",
                evidence={"first_row": header_texts},
            )
        )
    return items


def _first_words(block: ContentBlock) -> str:
    text = getattr(block, "text", None) or " ".join(
        getattr(item, "text", "") for item in getattr(block, "items", ())
    )
    text = " ".join((text or block.kind).replace("*", "").replace("\\", "").split())
    return text if len(text) <= 40 else text[:37] + "..."


def _reading_order_item(slide: SlideDocument) -> ReviewItem | None:
    review_ = slide.reading_order
    if review_ is None:
        return None
    by_id = {block.id: block for block in slide.blocks if block.id}
    target = f"slide[{slide.page}]"
    orders = {KEEP: [block.id for block in slide.blocks if block.id]}
    options = [Option(id=KEEP, label="Keep the converter's order")]
    for alternative in review_.alternatives:
        option_id = _slug(alternative.label)
        if option_id in orders:
            continue
        orders[option_id] = list(alternative.block_ids)
        options.append(
            Option(
                id=option_id,
                label=f"Read it {alternative.label}",
                rule=Rule(
                    target=target,
                    order=alternative.block_ids,
                    reason=f"reading order: {alternative.label}",
                ),
            )
        )
    return _item(
        "reading-order",
        target,
        slide.page,
        question="The layout can be read in more than one order "
        f"({', '.join(review_.reasons)}). Which order is right?",
        reasons=tuple(_ORDER_REASONS.get(reason, _slug(reason)) for reason in review_.reasons),
        look="slide",
        impact="high",
        options=tuple(options),
        suggested=KEEP,
        evidence={
            # Option id -> the first words of each block, in that order.
            "orders": {
                option_id: [_first_words(by_id[block_id]) for block_id in ids if block_id in by_id]
                for option_id, ids in orders.items()
            },
        },
    )


def _heading_item(block: HeadingBlock, page: int) -> ReviewItem:
    assert block.id is not None
    return _item(
        "heading",
        block.id,
        page,
        question=f"{block.text!r} is a heading only because of its font size. Is it a heading?",
        reasons=("font-size-only",),
        look="text",
        impact="low",
        options=(
            Option(id=KEEP, label="Yes, a heading"),
            Option(
                id="text",
                label="No, plain text",
                rule=Rule(target=block.id, as_text=True, reason="large text, not a heading"),
            ),
        ),
        suggested=KEEP,
        evidence={"level": block.level},
    )


def _image_item(block: ImageBlock, page: int) -> ReviewItem | None:
    alt = block.alt_text.strip()
    if block.description or block.id is None or (alt and not _DEFAULT_PICTURE_NAME.match(alt)):
        return None
    return _item(
        "image-description",
        block.id,
        page,
        question="This image has no description. What does it show?",
        reasons=("default-picture-name" if alt else "no-alt-text",),
        look="image",
        impact="high",
        options=(
            Option(id=KEEP, label="Leave it undescribed"),
            Option(
                id="decorative",
                label="Decorative; drop it",
                rule=Rule(target=block.id, exclude=True, reason="decorative image"),
            ),
        ),
        suggested=KEEP,
        free_text="description",
        asset_id=block.asset_id,
        evidence={"alt_text": alt},
    )


def _subheading_item(blocks: tuple[ContentBlock, ...], index: int, page: int) -> ReviewItem | None:
    block = blocks[index]
    assert isinstance(block, TextBlock) and block.id is not None
    match = _BOLD_LINE.match(block.text.strip())
    following = blocks[index + 1] if index + 1 < len(blocks) else None
    introduces = isinstance(following, (ListBlock, TableBlock)) or (
        isinstance(following, TextBlock)
        and following.shape_id is not None
        and following.shape_id == block.shape_id
    )
    if not (match and len(match.group("text")) <= 60 and introduces):
        return None
    assert following is not None
    level = min(
        6,
        max((b.level for b in blocks[:index] if isinstance(b, HeadingBlock)), default=1) + 1,
    )
    return _item(
        "subheading",
        block.id,
        page,
        question=f"Is the bold line {match.group('text')!r} a subheading?",
        reasons=("bold-line-introduces-block",),
        look="text",
        impact="low",
        options=(
            Option(id=KEEP, label="Keep as bold text"),
            Option(
                id="heading",
                label=f"Make it a level-{level} heading",
                rule=Rule(
                    target=block.id,
                    heading_level=level,
                    reason="bold line introducing the next block",
                ),
            ),
        ),
        suggested="heading",
        evidence={"followed_by": following.kind},
    )


def review(document: PresentationDocument) -> tuple[ReviewItem, ...]:
    """Questions about the conversion, in reading order."""

    items: list[ReviewItem] = []
    for slide in document.slides:
        if (order_item := _reading_order_item(slide)) is not None:
            items.append(order_item)
        blocks = slide.blocks
        for index, block in enumerate(blocks):
            if block.id is None:
                continue
            item: ReviewItem | None = None
            if isinstance(block, HeadingBlock) and block.source == "font":
                item = _heading_item(block, slide.page)
            elif isinstance(block, TableBlock):
                items.extend(_table_items(block, slide.page))
            elif isinstance(block, ImageBlock):
                item = _image_item(block, slide.page)
            elif isinstance(block, TextBlock):
                item = _subheading_item(blocks, index, slide.page)
            if item is not None:
                items.append(item)
    return tuple(items)


def review_data(items: Sequence[ReviewItem]) -> list[dict[str, Any]]:
    """Review items as JSON-ready data; each option carries its recipe rule."""

    data: list[dict[str, Any]] = []
    for item in items:
        entry = item.model_dump(mode="json", exclude={"options"})
        entry["options"] = [
            {
                "id": option.id,
                "label": option.label,
                "rule": option.rule.model_dump(mode="json", exclude_none=True)
                if option.rule
                else None,
            }
            for option in item.options
        ]
        data.append(entry)
    return data


def review_report(
    items: Sequence[ReviewItem], shown: Sequence[ReviewItem] | None = None
) -> dict[str, Any]:
    """The form the CLI and MCP server return: schema version, counts, items.

    The counts (by topic, look and impact) cover all ``items``, so a caller can plan
    what to render or fetch; ``shown`` limits the items listed, as a filter would.
    """

    def count(values: list[str]) -> dict[str, int]:
        return {value: values.count(value) for value in dict.fromkeys(values)}

    return {
        "schema_version": REVIEW_SCHEMA_VERSION,
        "counts": {
            "total": len(items),
            "topic": count([item.topic for item in items]),
            "look": count([item.look for item in items]),
            "impact": count([item.impact for item in items]),
        },
        "items": review_data(items if shown is None else shown),
    }


def review_json_schema() -> dict[str, Any]:
    """The JSON Schema of one review item."""

    return ReviewItem.model_json_schema()


def recipe_from_answers(
    items: Sequence[ReviewItem],
    answers: Mapping[str, str],
    *,
    unanswered: Literal["keep", "suggested"] = "keep",
) -> Recipe:
    """The recipe for ``answers``, a map of item id to option id.

    Where an item allows ``free_text``, an answer that is no option id fills that
    rule field instead (an image's description). Items left out keep the
    converter's decision, or take their suggested option with
    ``unanswered="suggested"``. Unknown item ids and option ids raise
    ``ValueError`` listing every problem, so a caller can correct all at once.
    """

    by_id = {item.id: item for item in items}
    problems = [f"no review item {item_id!r}" for item_id in answers if item_id not in by_id]
    rules: list[Rule] = []
    for item in items:
        answer = answers.get(item.id)
        if answer is None:
            if unanswered == "keep":
                continue
            answer = item.suggested
        option = item.option(answer)
        if option is not None:
            if option.rule is not None:
                rules.append(option.rule)
        elif item.free_text is not None and answer.strip():
            rules.append(
                Rule(
                    target=item.target,
                    reason="answered in review",
                    **{item.free_text: answer.strip()},
                )
            )
        else:
            choices = ", ".join(option.id for option in item.options)
            problems.append(f"{item.id}: {answer!r} is not one of {choices}")
    if problems:
        raise ValueError("; ".join(problems))
    return Recipe(rules=tuple(rules))
