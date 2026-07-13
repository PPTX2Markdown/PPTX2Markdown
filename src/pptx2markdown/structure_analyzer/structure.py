from __future__ import annotations

import re
import statistics
from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence, Tuple

from .constants import LARGE_INT, TITLE_TYPES
from .text_rules import (
    is_numbered_heading_text,
    numbered_heading_kind,
    numbered_suggested_depth,
)

EMU_PER_INCH = 914400
SCREEN_DPI = 96
XYCUT_MAX_OVERLAP_EMU = EMU_PER_INCH // SCREEN_DPI
XYCUT_VERTICAL_FLOW_GAP_RATIO = 0.5


@dataclass
class SlideObject:
    shape_id: str
    xml_index: int
    tag: str
    name: str
    ph_type: Optional[str]
    ph_idx: Optional[str]
    x: int
    y: int
    coord_source: str
    text: str
    normalized: str
    is_footer: bool
    is_decorative: bool
    is_heading: bool
    is_title_placeholder: bool
    font_pt: Optional[float]
    bbox: Optional[Tuple[int, int, int, int]] = None
    group_path: Tuple[str, ...] = ()
    z_path: Tuple[int, ...] = ()
    source_part: str = "slide"
    inheritance_kind: str = "direct"
    list_kind: Optional[str] = None
    list_level: Optional[int] = None


@dataclass
class OrderContext:
    slide_left: int
    slide_top: int
    slide_right: int
    slide_bottom: int
    slide_width: int
    slide_height: int
    title_band_bottom: int


def object_left(obj: SlideObject) -> int:
    if obj.bbox is not None:
        return obj.bbox[0]
    return obj.x


def object_top(obj: SlideObject) -> int:
    if obj.bbox is not None:
        return obj.bbox[1]
    return obj.y


def object_right(obj: SlideObject) -> int:
    if obj.bbox is not None:
        return obj.bbox[2]
    width = max(500000, min(2500000, 80000 * max(1, len(obj.normalized))))
    return object_left(obj) + width


def object_height(obj: SlideObject) -> int:
    if obj.bbox is not None:
        return max(1, obj.bbox[3] - obj.bbox[1])
    if obj.tag == "pic":
        return 1800000
    if obj.tag == "graphicFrame":
        return 1200000
    lines = max(1, len(re.findall(r"[.!?]|[\u3002]", obj.text)) + 1)
    return max(250000, min(1800000, 250000 * lines))


def object_bottom(obj: SlideObject) -> int:
    if obj.bbox is not None:
        return obj.bbox[3]
    return object_top(obj) + object_height(obj)


def build_order_context(objects: Sequence[SlideObject]) -> OrderContext:
    candidates = [o for o in objects if not o.is_footer and not o.is_decorative]
    if not candidates:
        candidates = list(objects)

    if candidates:
        slide_left = min(object_left(o) for o in candidates)
        slide_top = min(object_top(o) for o in candidates)
        slide_right = max(object_right(o) for o in candidates)
        slide_bottom = max(object_bottom(o) for o in candidates)
    else:
        slide_left = slide_top = 0
        slide_right = slide_bottom = 1

    slide_width = max(1, slide_right - slide_left)
    slide_height = max(1, slide_bottom - slide_top)
    title_band_bottom = slide_top + max(600000, int(slide_height * 0.22))

    return OrderContext(
        slide_left=slide_left,
        slide_top=slide_top,
        slide_right=slide_right,
        slide_bottom=slide_bottom,
        slide_width=slide_width,
        slide_height=slide_height,
        title_band_bottom=title_band_bottom,
    )


def is_top_title_object(obj: SlideObject, context: OrderContext) -> bool:
    if obj.is_footer or obj.is_decorative:
        return False
    top = object_top(obj)
    if obj.ph_type in TITLE_TYPES:
        return True
    if obj.is_title_placeholder:
        return True
    if (
        obj.coord_source == "layout"
        and obj.ph_type is not None
        and top <= context.title_band_bottom
    ):
        return len(obj.normalized) <= 100
    return False


def _sort_top_left(objects: Sequence[SlideObject]) -> List[SlideObject]:
    return sorted(
        objects,
        key=lambda obj: (
            object_top(obj),
            object_left(obj),
            obj.xml_index,
        ),
    )


def _has_reliable_position(obj: SlideObject) -> bool:
    if obj.bbox is None:
        return False
    left = object_left(obj)
    top = object_top(obj)
    right = object_right(obj)
    bottom = object_bottom(obj)
    return (
        left < LARGE_INT // 2
        and top < LARGE_INT // 2
        and right < LARGE_INT // 2
        and bottom < LARGE_INT // 2
        and right > left
        and bottom > top
    )


def _axis_interval(obj: SlideObject, axis: str) -> Tuple[int, int]:
    """PPTX EMU 좌표에서 한 축이 차지하는 구간을 반환한다."""
    if axis == "x":
        start = object_left(obj)
        end = object_right(obj)
    elif axis == "y":
        start = object_top(obj)
        end = object_bottom(obj)
    else:
        raise ValueError(f"unsupported xycut axis: {axis}")

    if end <= start:
        end = start + 1
    return start, end


def _sort_for_axis(objects: Sequence[SlideObject], axis: str) -> List[SlideObject]:
    """축 projection을 만들기 전에 사용할 순회 순서로 객체를 정렬한다."""
    if axis == "x":
        return sorted(
            objects,
            key=lambda obj: (object_left(obj), object_top(obj), obj.xml_index),
        )
    if axis == "y":
        return sorted(
            objects,
            key=lambda obj: (object_top(obj), object_left(obj), obj.xml_index),
        )
    raise ValueError(f"unsupported xycut axis: {axis}")


def _projection_chunks(
    objects: Sequence[SlideObject],
    axis: str,
    *,
    max_overlap: int,
) -> List[List[SlideObject]]:
    """Split objects at projection gaps without building a dense EMU array."""
    axis_sorted = _sort_for_axis(objects, axis)
    if not axis_sorted:
        return []

    chunks: List[List[SlideObject]] = []
    current_chunk = [axis_sorted[0]]
    _, chunk_end = _axis_interval(axis_sorted[0], axis)

    for obj in axis_sorted[1:]:
        start, end = _axis_interval(obj, axis)
        if start - chunk_end >= -max_overlap:
            chunks.append(current_chunk)
            current_chunk = [obj]
            chunk_end = end
            continue
        current_chunk.append(obj)
        chunk_end = max(chunk_end, end)

    chunks.append(current_chunk)
    return chunks if len(chunks) > 1 else []


def _is_tight_vertical_sequence(y_chunks: Sequence[Sequence[SlideObject]]) -> bool:
    """Return true when geometry forms one closely spaced top-to-bottom stream."""
    if len(y_chunks) <= 1 or any(len(chunk) != 1 for chunk in y_chunks):
        return False

    ordered = [chunk[0] for chunk in y_chunks]
    gaps = [
        max(0, object_top(current) - object_bottom(previous))
        for previous, current in zip(ordered, ordered[1:])
    ]
    heights = [object_height(obj) for obj in ordered]
    return statistics.median(gaps) <= (statistics.median(heights) * XYCUT_VERTICAL_FLOW_GAP_RATIO)


def _recursive_xycut(objects: Sequence[SlideObject], *, max_overlap: int) -> List[SlideObject]:
    """recursive bbox 기반 XY cut으로 객체를 정렬한다.

    촘촘한 단일 Y 흐름은 위에서 아래로 유지한다. 그 외에는 현재 영역을
    관통하는 X축 공백으로 열을 먼저 나누고, X축 절단이 불가능하면 Y축의
    첫 공백으로 위/아래를 나눈다. 의미 정보는 사용하지 않는다.
    """
    if len(objects) <= 1:
        return list(objects)

    x_chunks = _projection_chunks(objects, "x", max_overlap=max_overlap)
    y_chunks = _projection_chunks(objects, "y", max_overlap=max_overlap)

    if x_chunks and _is_tight_vertical_sequence(y_chunks):
        return [obj for chunk in y_chunks for obj in chunk]

    if x_chunks:
        return [
            obj for chunk in x_chunks for obj in _recursive_xycut(chunk, max_overlap=max_overlap)
        ]

    if y_chunks:
        top_chunk = y_chunks[0]
        remaining = [obj for chunk in y_chunks[1:] for obj in chunk]
        return _recursive_xycut(top_chunk, max_overlap=max_overlap) + _recursive_xycut(
            remaining,
            max_overlap=max_overlap,
        )

    return _sort_top_left(objects)


def order_objects(objects: Sequence[SlideObject]) -> List[SlideObject]:
    """Order positioned objects only by XYCut geometry."""
    # A top-level decorative frame is visual chrome and must not seal an
    # otherwise valid column gap. Inside a group, however, that same frame
    # carries the group's visual boundary; retaining it prevents nested rows
    # from being flattened into unrelated outer columns.
    meaningful = [obj for obj in objects if not obj.is_footer and not obj.is_decorative]

    def carries_group_boundary(obj: SlideObject) -> bool:
        return bool(obj.group_path) and any(
            candidate.group_path[: len(obj.group_path)] == obj.group_path
            for candidate in meaningful
        )

    structural_decorators = [
        obj for obj in objects if obj.is_decorative and carries_group_boundary(obj)
    ]
    candidates = meaningful + structural_decorators
    candidate_ids = {id(obj) for obj in candidates}
    excluded = [obj for obj in objects if id(obj) not in candidate_ids]
    positioned = [obj for obj in candidates if _has_reliable_position(obj)]
    unpositioned = [obj for obj in candidates if not _has_reliable_position(obj)]
    ordered = _recursive_xycut(positioned, max_overlap=XYCUT_MAX_OVERLAP_EMU)
    return (
        ordered
        + sorted(unpositioned, key=lambda obj: obj.xml_index)
        + sorted(excluded, key=lambda obj: obj.xml_index)
    )


def compute_heading_depths(
    ordered_objects: Sequence[SlideObject],
    context: OrderContext,
    strict: bool = False,
) -> Dict[str, Optional[int]]:
    depths: Dict[str, Optional[int]] = {}
    seen_title = False
    last_section_depth: Optional[int] = None

    heading_fonts = [
        float(obj.font_pt)
        for obj in ordered_objects
        if obj.is_heading and obj.font_pt is not None and obj.font_pt > 0
    ]

    font_tolerance = 1.0
    if heading_fonts:
        median_font = statistics.median(heading_fonts)
        font_tolerance = max(1.0, float(median_font) * 0.06)

    bands: List[float] = []
    for size in sorted(heading_fonts, reverse=True):
        if not bands or abs(size - bands[-1]) > font_tolerance:
            bands.append(size)

    def depth_from_font(size: Optional[float]) -> Optional[int]:
        if size is None or size <= 0 or not bands:
            return None
        closest_idx = min(range(len(bands)), key=lambda i: abs(size - bands[i]))
        if abs(size - bands[closest_idx]) <= font_tolerance:
            return min(closest_idx + 1, 6)
        return None

    prev_heading_depth: Optional[int] = None
    prev_heading_font: Optional[float] = None
    h1_assigned = False
    seen_numbered_heading = False

    for obj in ordered_objects:
        depth: Optional[int] = None
        if strict:
            if obj.is_heading:
                if obj.ph_type in {"title", "ctrTitle"}:
                    depth = 1
                elif obj.ph_type == "subTitle":
                    depth = 2
            depths[obj.shape_id] = depth
            continue

        kind = numbered_heading_kind(obj.text) if obj.is_heading else None

        if is_top_title_object(obj, context):
            depth = 1 if not h1_assigned else 2
            seen_title = True
            last_section_depth = 1
        elif obj.is_heading:
            font_depth = depth_from_font(obj.font_pt)
            if (
                font_depth is not None
                and prev_heading_depth is not None
                and obj.font_pt is not None
                and prev_heading_font is not None
                and abs(float(obj.font_pt) - float(prev_heading_font)) <= font_tolerance
            ):
                depth = prev_heading_depth
            elif font_depth is not None:
                depth = font_depth
            else:
                if kind == "dotted-multi":
                    depth = 2 if seen_title else 1
                    last_section_depth = depth
                elif kind == "dotted-single":
                    depth = 2 if seen_title else 1
                    last_section_depth = depth
                else:
                    depth = 2 if seen_title else 1
                    last_section_depth = depth

            suggested_depth = numbered_suggested_depth(
                obj.text,
                last_section_depth=last_section_depth,
                seen_title=seen_title,
            )
            if suggested_depth is not None:
                if not seen_numbered_heading:
                    suggested_depth = min(suggested_depth, 2)
                if depth is None:
                    depth = suggested_depth
                else:
                    depth = max(depth, suggested_depth)
                seen_numbered_heading = True

            if h1_assigned and depth == 1:
                depth = 2
            if (
                prev_heading_depth is not None
                and depth is not None
                and depth > (prev_heading_depth + 1)
            ):
                depth = prev_heading_depth + 1

            if depth is not None:
                if depth == 1:
                    h1_assigned = True
                prev_heading_depth = depth
                prev_heading_font = obj.font_pt

        depths[obj.shape_id] = depth

    return depths


def heading_score(obj: SlideObject, strict: bool = False) -> float:
    if strict:
        if not obj.is_heading:
            return 0.0
        if obj.ph_type in {"title", "ctrTitle"}:
            return 0.93
        if obj.ph_type == "subTitle":
            return 0.90
        return 0.0

    if obj.is_heading and is_numbered_heading_text(obj.text):
        return 0.95
    if obj.is_heading:
        return 0.75
    return 0.0


def heading_threshold(strict: bool = False) -> float:
    return 0.88 if strict else 0.7
