"""Reading-order choices from any judge: only flagged slides are rendered and shown.

The parser flags slides whose layout the geometry cannot settle and keeps their
candidate orders (``SlideDocument.reading_order``). ``choose_reading_orders_with``
hands each flagged slide to a function you supply: the slide rendered as a
one-page PDF by LibreOffice (``snapshot.slide_pdf``, when a source file and
LibreOffice are available), the reasons, and each candidate order as the first
words of its blocks. The
function names the candidate it picks, and the picks come back as ``order``
recipe rules. Slides that are not flagged are never rendered, so the cost stays
with the few slides that need a look.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING

from .model import ContentBlock, SlideDocument
from .recipe import Recipe, Rule
from .snapshot import slide_pdf

if TYPE_CHECKING:
    from .presentation import ParsedPresentation

AS_CONVERTED = "as converted"


@dataclass(frozen=True)
class OrderRequest:
    """One flagged slide: how it looks, why it was flagged, and the candidate orders."""

    page: int
    reasons: tuple[str, ...]
    # Candidate label -> the first words of each block, in that order. The
    # converter's own order is listed first, under ``AS_CONVERTED``.
    candidates: dict[str, list[str]]
    # The slide as a one-page PDF, or None when it could not be rendered.
    slide_pdf: bytes | None = None


Judge = Callable[[OrderRequest], "str | None"]


def _words(block: ContentBlock) -> str:
    text = getattr(block, "text", None) or " ".join(
        getattr(item, "text", "") for item in getattr(block, "items", ())
    )
    text = " ".join((text or f"[{block.kind}]").replace("*", "").replace("\\", "").split())
    return text if len(text) <= 60 else text[:57] + "..."


def _candidates(slide: SlideDocument) -> dict[str, tuple[str, ...]]:
    assert slide.reading_order is not None
    orders = {AS_CONVERTED: tuple(block.id for block in slide.blocks if block.id)}
    for alternative in slide.reading_order.alternatives:
        orders.setdefault(alternative.label, alternative.block_ids)
    return orders


def choose_reading_orders_with(
    presentation: "ParsedPresentation",
    judge: Judge,
    *,
    source: str | Path | None = None,
    render: bool = True,
    pages: set[int] | None = None,
) -> Recipe:
    """A recipe with the judge's pick for every flagged slide (or those in ``pages``).

    ``source`` is the .pptx to render slides from; it defaults to the file the
    presentation was parsed from. With ``render=False``, or without a source or
    LibreOffice, the judge gets the candidates as text only. A judge returning
    None or ``AS_CONVERTED`` leaves the slide as converted.
    """

    source = source if source is not None else presentation.source_path
    rules: list[Rule] = []
    for slide in presentation.document.slides:
        if slide.reading_order is None or (pages is not None and slide.page not in pages):
            continue
        orders = _candidates(slide)
        by_id = {block.id: block for block in slide.blocks}
        request = OrderRequest(
            page=slide.page,
            reasons=slide.reading_order.reasons,
            candidates={
                label: [_words(by_id[block_id]) for block_id in ids if block_id in by_id]
                for label, ids in orders.items()
            },
            slide_pdf=(slide_pdf(source, slide.page) if render and source is not None else None),
        )
        choice = judge(request)
        if choice is None or choice == AS_CONVERTED or choice not in orders:
            continue
        rules.append(
            Rule(
                target=f"slide[{slide.page}]",
                order=orders[choice],
                reason=f"reading order: {choice}",
            )
        )
    return Recipe(rules=tuple(rules))
