"""Deterministic geometry-only reading order."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Generic, Iterable, TypeVar

T = TypeVar("T")


@dataclass(frozen=True)
class LayoutNode(Generic[T]):
    """A payload with an optional EMU bounding box and source order."""

    payload: T
    source_order: int
    bbox: tuple[int, int, int, int] | None


def _axis_groups(
    nodes: list[LayoutNode[T]],
    *,
    axis: int,
) -> list[list[LayoutNode[T]]]:
    """Partition nodes at every positive gap along one axis."""

    intervals = sorted(
        nodes,
        key=lambda node: (
            node.bbox[axis] if node.bbox is not None else 0,
            node.bbox[axis + 2] if node.bbox is not None else 0,
            node.source_order,
        ),
    )
    if len(intervals) < 2:
        return [intervals]

    groups: list[list[LayoutNode[T]]] = []
    current: list[LayoutNode[T]] = []
    running_end: int | None = None
    for node in intervals:
        assert node.bbox is not None
        start = node.bbox[axis]
        end = node.bbox[axis + 2]
        if current and running_end is not None and start > running_end:
            groups.append(current)
            current = []
            running_end = None
        current.append(node)
        running_end = end if running_end is None else max(running_end, end)
    if current:
        groups.append(current)
    return groups


def _recursive_order(nodes: list[LayoutNode[T]], *, prefer_vertical: bool) -> list[LayoutNode[T]]:
    if len(nodes) < 2:
        return nodes

    axes = (1, 0) if prefer_vertical else (0, 1)
    for axis in axes:
        groups = _axis_groups(nodes, axis=axis)
        if len(groups) <= 1:
            continue
        ordered: list[LayoutNode[T]] = []
        for group in groups:
            ordered.extend(_recursive_order(group, prefer_vertical=(axis == 0)))
        return ordered

    return sorted(
        nodes,
        key=lambda node: (
            node.bbox[1] if node.bbox is not None else 0,
            node.bbox[0] if node.bbox is not None else 0,
            node.source_order,
        ),
    )


def geometry_reading_order(nodes: Iterable[LayoutNode[T]]) -> list[T]:
    """Order positioned shapes geometrically and append unpositioned shapes."""

    materialized = list(nodes)
    positioned = [node for node in materialized if node.bbox is not None]
    unpositioned = [node for node in materialized if node.bbox is None]
    ordered = _recursive_order(positioned, prefer_vertical=True)
    ordered.extend(sorted(unpositioned, key=lambda node: node.source_order))
    return [node.payload for node in ordered]
