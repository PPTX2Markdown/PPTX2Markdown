"""Deterministic geometry-only reading order (recursive XY-cut).

A slide is read column by column: wherever a vertical gap runs through the whole
region the columns are read left to right, each on its own. Where no column gap
exists, the topmost horizontal band is read first and the rest is cut again, so
a full-width title or divider still exposes the columns below it. A plain stack
of boxes is kept top to bottom even when the boxes are slightly offset.

Containers are read as units: a group, or a shape that fully encloses other
shapes (a card's frame, a labelled box, a picture with callouts on it), is placed
among its neighbours by its own box and read through before the next unit; its
own text comes first. See docs/reading_order_design.md.
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
# Exported decks let text boxes stick out of their frame a little.
CONTAIN_TOLERANCE_EMU = EMU_PER_INCH // 10


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
    # Groups holding the node, outermost first; a group is read as one unit.
    groups: tuple[int, ...] = ()
    # The node may hold the nodes its box encloses (a frame, a labelled shape).
    container: bool = False
    # What the node encloses are labels placed on it (callouts on a picture): they
    # are scattered, not laid out in columns, and are read top to bottom.
    annotated: bool = False


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


@dataclass
class _Unit(Generic[T]):
    """A node, or a container read as a whole: its own node first, then its members."""

    bbox: tuple[int, int, int, int]
    source_order: int
    node: LayoutNode[T] | None = None
    members: list["_Unit[T]"] | None = None

    def all_nodes(self) -> list[LayoutNode[T]]:
        head = [self.node] if self.node is not None else []
        return head + [node for unit in self.members or [] for node in unit.all_nodes()]

    def structured(self) -> bool:
        """Whether anything in the unit is read: a decorative group or frame is not."""

        return any(node.structure for node in self.all_nodes())

    def nodes(self) -> list[LayoutNode[T]]:
        head = [self.node] if self.node is not None else []
        if not self.members:
            return head
        # Decoration inside a container (brackets, rules) must not split it either.
        readable = [unit for unit in self.members if unit.structured()]
        decoration = [
            node for unit in self.members if not unit.structured() for node in unit.all_nodes()
        ]
        if self.node is not None and self.node.annotated:
            members = sorted(readable, key=lambda u: (u.bbox[1], u.bbox[0], u.source_order))
        else:
            members = _order_units(readable)
        return head + [node for unit in members for node in unit.nodes()] + decoration


def _area(bbox: tuple[int, int, int, int]) -> int:
    return max(0, bbox[2] - bbox[0]) * max(0, bbox[3] - bbox[1])


def _encloses(outer: tuple[int, int, int, int], inner: tuple[int, int, int, int]) -> bool:
    # A box may stick out by a little, but never by a large part of its own size.
    across = min(CONTAIN_TOLERANCE_EMU, (inner[2] - inner[0]) // 4)
    down = min(CONTAIN_TOLERANCE_EMU, (inner[3] - inner[1]) // 4)
    return (
        inner[0] >= outer[0] - across
        and inner[1] >= outer[1] - down
        and inner[2] <= outer[2] + across
        and inner[3] <= outer[3] + down
        and _area(inner) < _area(outer)
    )


def _nest(units: list[_Unit[T]]) -> list[_Unit[T]]:
    """Move units into the smallest container among them that encloses them."""

    containers = [unit for unit in units if unit.node is not None and unit.node.container]
    parent: dict[int, _Unit[T]] = {}
    for unit in units:
        enclosing = [
            container
            for container in containers
            if container is not unit and _encloses(container.bbox, unit.bbox)
        ]
        if enclosing:
            parent[id(unit)] = min(enclosing, key=lambda c: (_area(c.bbox), c.source_order))
    top: list[_Unit[T]] = []
    for unit in units:
        holder = parent.get(id(unit))
        if holder is None:
            top.append(unit)
        else:
            holder.members = [*(holder.members or []), unit]
    return top


def _units(
    nodes: list[LayoutNode[T]], spare: list[LayoutNode[T]], depth: int = 0
) -> list[_Unit[T]]:
    """Units at one group depth: loose nodes, and one unit per group.

    A unit with nothing to read (a frame enclosing nothing, a group of decorative
    shapes) is decoration again: it goes to ``spare`` so it cannot seal a gap.
    """

    loose: list[_Unit[T]] = []
    grouped: dict[int, list[LayoutNode[T]]] = {}
    for node in nodes:
        if len(node.groups) > depth:
            grouped.setdefault(node.groups[depth], []).append(node)
        else:
            assert node.bbox is not None
            loose.append(_Unit(node.bbox, node.source_order, node=node))
    for members in grouped.values():
        inner = _units(members, spare, depth + 1)
        if not inner:
            continue
        boxes = [unit.bbox for unit in inner]
        bbox = (
            min(b[0] for b in boxes),
            min(b[1] for b in boxes),
            max(b[2] for b in boxes),
            max(b[3] for b in boxes),
        )
        loose.append(_Unit(bbox, min(unit.source_order for unit in inner), members=inner))
    kept: list[_Unit[T]] = []
    for unit in _nest(loose):
        if unit.structured():
            kept.append(unit)
        else:
            spare.extend(unit.all_nodes())
    return kept


def _order_units(units: list[_Unit[T]]) -> list[_Unit[T]]:
    nodes = [LayoutNode(unit, unit.source_order, unit.bbox) for unit in units]
    return [node.payload for node in _xycut(nodes)]


def geometry_reading_order(nodes: Iterable[LayoutNode[T]]) -> list[T]:
    """Order positioned shapes by containers and XY-cut, then the rest.

    Unpositioned nodes follow in source order, and nodes that do not structure the
    layout (backgrounds, footers) come last, unless they are containers: a frame
    holding text is read where its contents are.
    """

    materialized = list(nodes)
    placed = [
        node
        for node in materialized
        if node.bbox is not None and (node.structure or node.container)
    ]
    unpositioned = [node for node in materialized if node.bbox is None and node.structure]
    background = [
        node
        for node in materialized
        if not node.structure and not (node.container and node.bbox is not None)
    ]
    ordered = [node for unit in _order_units(_units(placed, background)) for node in unit.nodes()]
    ordered.extend(sorted(unpositioned, key=lambda node: node.source_order))
    ordered.extend(sorted(background, key=lambda node: node.source_order))
    return [node.payload for node in ordered]
