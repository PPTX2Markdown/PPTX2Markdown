"""Where the deterministic converter is least sure, as questions with recipe answers.

The parser always decides; ``review`` points at the decisions most worth a second
look and offers each alternative as a ready-made recipe rule. A person, a UI or a
language model then only has to pick options, and the choice stays a recipe.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .model import (
    HeadingBlock,
    ImageBlock,
    ListBlock,
    PresentationDocument,
    TableBlock,
    TableCell,
    TextBlock,
)
from .recipe import Rule

_BOLD_LINE = re.compile(r"^\*\*(?P<text>[^*\n]+)\*\*$")
_DEFAULT_PICTURE_NAME = re.compile(
    r"^(?:picture|image|graphic|grafik|imagen|그림|이미지|図|图片)\s*\d*$", re.IGNORECASE
)


@dataclass(frozen=True)
class Option:
    """One answer to a review question; ``rule`` is None for "keep as converted"."""

    label: str
    rule: Rule | None = None


@dataclass(frozen=True)
class ReviewItem:
    """A converter decision worth checking, with its alternatives."""

    topic: str
    target: str
    question: str
    options: tuple[Option, ...]
    suggested: int = 0
    evidence: dict[str, object] = field(default_factory=dict)


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


def _table_items(block: TableBlock) -> list[ReviewItem]:
    rows = block.rows
    if not rows or block.id is None:
        return []
    items: list[ReviewItem] = []
    start = block.caption_rows
    if block.caption_rows:
        title = rows[0][0].text
        items.append(
            ReviewItem(
                topic="table-caption",
                target=block.id,
                question=f"Is the merged first row {title!r} a title above the table?",
                options=(
                    Option("Yes, a title above the table"),
                    Option(
                        "No, it is the header",
                        Rule(target=block.id, caption_rows=0, reason="first row is the header"),
                    ),
                ),
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
            ReviewItem(
                topic="table-header-rows",
                target=block.id,
                question="The header has merged groups; is the next row part of the header?",
                options=(
                    Option("One header row"),
                    Option(
                        "Two header rows",
                        Rule(target=block.id, header_rows=2, reason="grouped two-row header"),
                    ),
                ),
                suggested=1,
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
            ReviewItem(
                topic="table-header-rows",
                target=block.id,
                question="The first row is all numbers; does the table have a header?",
                options=(
                    Option("The first row is the header"),
                    Option(
                        "No header row",
                        Rule(target=block.id, header_rows=0, reason="first row is data"),
                    ),
                ),
                suggested=1,
                evidence={"first_row": header_texts},
            )
        )
    return items


def review(document: PresentationDocument) -> tuple[ReviewItem, ...]:
    """Questions about the conversion, in reading order."""

    items: list[ReviewItem] = []
    for slide in document.slides:
        blocks = slide.blocks
        for index, block in enumerate(blocks):
            if block.id is None:
                continue
            if isinstance(block, TableBlock):
                items.extend(_table_items(block))
            elif isinstance(block, ImageBlock) and not block.description:
                alt = block.alt_text.strip()
                if not alt or _DEFAULT_PICTURE_NAME.match(alt):
                    items.append(
                        ReviewItem(
                            topic="image-description",
                            target=block.id,
                            question="This image has no description. What does it show?",
                            options=(
                                Option("Leave it undescribed"),
                                Option(
                                    "Decorative; drop it",
                                    Rule(target=block.id, exclude=True, reason="decorative image"),
                                ),
                            ),
                            evidence={"alt_text": alt},
                        )
                    )
            elif isinstance(block, TextBlock):
                match = _BOLD_LINE.match(block.text.strip())
                following = blocks[index + 1] if index + 1 < len(blocks) else None
                introduces = isinstance(following, (ListBlock, TableBlock)) or (
                    isinstance(following, TextBlock)
                    and following.shape_id is not None
                    and following.shape_id == block.shape_id
                )
                if match and len(match.group("text")) <= 60 and introduces:
                    level = min(
                        6,
                        max(
                            (b.level for b in blocks[:index] if isinstance(b, HeadingBlock)),
                            default=1,
                        )
                        + 1,
                    )
                    items.append(
                        ReviewItem(
                            topic="subheading",
                            target=block.id,
                            question=f"Is the bold line {match.group('text')!r} a subheading?",
                            options=(
                                Option("Keep as bold text"),
                                Option(
                                    f"Make it a level-{level} heading",
                                    Rule(
                                        target=block.id,
                                        heading_level=level,
                                        reason="bold line introducing the next block",
                                    ),
                                ),
                            ),
                            suggested=1,
                            evidence={"followed_by": following.kind},
                        )
                    )
    return tuple(items)


def review_data(items: tuple[ReviewItem, ...]) -> list[dict[str, object]]:
    """Review items as JSON-ready data; each option carries its recipe rule."""

    return [
        {
            "topic": item.topic,
            "target": item.target,
            "question": item.question,
            "options": [
                {
                    "label": option.label,
                    "rule": option.rule.model_dump(mode="json", exclude_none=True)
                    if option.rule
                    else None,
                }
                for option in item.options
            ],
            "suggested": item.suggested,
            "evidence": item.evidence,
        }
        for item in items
    ]
