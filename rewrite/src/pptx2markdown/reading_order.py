"""Deterministic geometry-only reading order (recursive XY-cut).

A slide is read column by column: wherever a vertical gap runs through the whole
region the columns are read left to right, each on its own. Where no column gap
exists, the topmost horizontal band is read first and the rest is cut again, so
a full-width title or divider still exposes the columns below it. A plain stack
of boxes is kept top to bottom even when the boxes are slightly offset.
"""

from __future__ import annotations

import statistics
from dataclasses import dataclass
from typing import Generic, Iterable, TypeVar

T = TypeVar("T")

EMU_PER_INCH = 914_400
# Boxes drawn by hand rarely line up exactly; an overlap of one screen pixel at
# 96 dpi still counts as a gap.
MAX_OVERLAP_EMU = EMU_PER_INCH // 96
# Single boxes whose median gap is at most half their median height form one flow.
FLOW_GAP_RATIO = 0.5


@dataclass(frozen=True)
class LayoutNode(Generic[T]):
    """A payload with an optional EMU bounding box and source order.

    A node that does not ``structure`` the layout (a background frame, a footer) is
    ignored when looking for gaps, so it cannot seal a column gap; it is read after
    the structured nodes.
    """

    payload: T
    source_order: int
    bbox: tuple[int, int, int, int] | None
    structure: bool = True


def _interval(node: LayoutNode[T], axis: int) -> tuple[int, int]:
    assert node.bbox is not None
    start, end = node.bbox[axis], node.bbox[axis + 2]
    return start, max(end, start + 1)


def _chunks(nodes: list[LayoutNode[T]], axis: int) -> list[list[LayoutNode[T]]]:
    """Groups separated by gaps along one axis; empty when there is no gap."""

    other = 1 - axis
    ordered = sorted(
        nodes,
        key=lambda node: (node.bbox[axis], node.bbox[other], node.source_order),  # type: ignore[index]
    )
    if not ordered:
        return []
    chunks = [[ordered[0]]]
    _, chunk_end = _interval(ordered[0], axis)
    for node in ordered[1:]:
        start, end = _interval(node, axis)
        if start - chunk_end >= -MAX_OVERLAP_EMU:
            chunks.append([node])
            chunk_end = end
        else:
            chunks[-1].append(node)
            chunk_end = max(chunk_end, end)
    return chunks if len(chunks) > 1 else []


def _is_single_flow(rows: list[list[LayoutNode[T]]]) -> bool:
    """One box per row, closely spaced: a stack read top to bottom."""

    if len(rows) <= 1 or any(len(row) != 1 for row in rows):
        return False
    boxes = [row[0] for row in rows]
    gaps = [
        max(0, _interval(lower, 1)[0] - _interval(upper, 1)[1])
        for upper, lower in zip(boxes, boxes[1:])
    ]
    heights = [_interval(box, 1)[1] - _interval(box, 1)[0] for box in boxes]
    return statistics.median(gaps) <= statistics.median(heights) * FLOW_GAP_RATIO


def _xycut(nodes: list[LayoutNode[T]]) -> list[LayoutNode[T]]:
    if len(nodes) <= 1:
        return list(nodes)
    columns = _chunks(nodes, axis=0)
    rows = _chunks(nodes, axis=1)
    if columns and _is_single_flow(rows):
        return [node for row in rows for node in row]
    if columns:
        return [node for column in columns for node in _xycut(column)]
    if rows:
        return _xycut(rows[0]) + _xycut([node for row in rows[1:] for node in row])
    return sorted(
        nodes,
        key=lambda node: (node.bbox[1], node.bbox[0], node.source_order),  # type: ignore[index]
    )


def geometry_reading_order(nodes: Iterable[LayoutNode[T]]) -> list[T]:
    """Order positioned shapes by XY-cut, then unpositioned and non-structural ones."""

    materialized = list(nodes)
    positioned = [node for node in materialized if node.bbox is not None and node.structure]
    unpositioned = [node for node in materialized if node.bbox is None and node.structure]
    background = [node for node in materialized if not node.structure]
    ordered = _xycut(positioned)
    ordered.extend(sorted(unpositioned, key=lambda node: node.source_order))
    ordered.extend(sorted(background, key=lambda node: node.source_order))
    return [node.payload for node in ordered]
