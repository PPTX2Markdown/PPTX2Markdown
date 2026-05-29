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
XYCUT_MIN_GAP_EMU = EMU_PER_INCH // SCREEN_DPI


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


def object_width(obj: SlideObject) -> int:
    return max(1, object_right(obj) - object_left(obj))


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


def is_full_width_body(obj: SlideObject, context: OrderContext) -> bool:
    if not obj.normalized:
        return False
    width_ratio = object_width(obj) / max(1, context.slide_width)
    height_ratio = object_height(obj) / max(1, context.slide_height)
    return width_ratio >= 0.58 and (height_ratio >= 0.12 or len(obj.normalized) >= 90)


def is_top_title_object(obj: SlideObject, context: OrderContext) -> bool:
    if obj.is_footer or obj.is_decorative:
        return False
    top = object_top(obj)
    if obj.ph_type in TITLE_TYPES:
        return True
    if obj.is_title_placeholder:
        return True
    if obj.coord_source == "layout" and obj.ph_type is not None and top <= context.title_band_bottom:
        return len(obj.normalized) <= 100
    return False


def reading_top(obj: SlideObject, context: OrderContext) -> int:
    top = object_top(obj)
    if top >= LARGE_INT and is_top_title_object(obj, context):
        return context.slide_top - max(250000, int(context.slide_height * 0.08))
    return top


def is_promotable_numbered_heading(obj: SlideObject, context: OrderContext) -> bool:
    if not is_numbered_heading_text(obj.text):
        return False
    if obj.is_footer or obj.is_decorative:
        return False
    if is_full_width_body(obj, context):
        return False
    top = object_top(obj)
    if top > context.slide_top + int(context.slide_height * 0.45):
        return False
    if object_height(obj) > int(context.slide_height * 0.18):
        return False
    if len(obj.normalized) > 90:
        return False
    return True


def bucket(obj: SlideObject, context: OrderContext) -> int:
    if obj.is_footer:
        return 4
    if obj.is_decorative:
        return 5
    if is_top_title_object(obj, context):
        return 0
    if is_promotable_numbered_heading(obj, context):
        return 1
    return 2


def reason(obj: SlideObject, context: OrderContext) -> str:
    if obj.is_footer:
        return "Footer or slide number placeholder"
    if obj.is_decorative:
        return "Decorative connector/shape, low reading priority"
    if is_top_title_object(obj, context):
        if obj.coord_source == "layout":
            return "Title placeholder with layout-inherited coordinates"
        return "Title/layout placeholder promoted in top title band"
    if is_promotable_numbered_heading(obj, context):
        return "Short numbered heading promoted by position and size"
    if obj.tag == "graphicFrame":
        return "Table/chart frame as a body block"
    if obj.tag == "pic":
        return "Image block (no text)"
    if is_full_width_body(obj, context):
        return "Full-width body block kept in spatial order"
    return "General body block ordered by row clustering"


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


def _projection_segments(
    objects: Sequence[SlideObject],
    axis: str,
    *,
    min_gap: int,
) -> List[Tuple[int, int]]:
    """bbox projection을 점유 구간 단위로 분리한다.

    Sanster/PaddleX 구현은 dense 1D projection 배열을 만든 뒤 projection > 0인
    run을 분리한다. PPTX는 EMU 좌표를 쓰므로 dense 배열을 만들면 너무 커진다.
    min_value == 0 조건에서는 bbox interval을 병합해도 같은 split 결과를 낼 수 있다.
    """
    intervals = sorted(_axis_interval(obj, axis) for obj in objects)
    if not intervals:
        return []

    segments: List[Tuple[int, int]] = []
    segment_start, segment_end = intervals[0]
    for start, end in intervals[1:]:
        if start - segment_end >= min_gap:
            segments.append((segment_start, segment_end))
            segment_start, segment_end = start, end
        else:
            segment_end = max(segment_end, end)
    segments.append((segment_start, segment_end))
    return segments


def _objects_starting_in_segment(
    objects: Sequence[SlideObject],
    axis: str,
    segment: Tuple[int, int],
) -> List[SlideObject]:
    """축 시작 좌표가 projection segment 안에 들어가는 객체를 반환한다."""
    start, end = segment
    return [obj for obj in objects if start <= _axis_interval(obj, axis)[0] < end]


def _recursive_xycut(objects: Sequence[SlideObject], *, min_gap: int) -> List[SlideObject]:
    """recursive bbox 기반 XY cut으로 객체를 정렬한다.

    일반적인 bbox XY-cut 구현 형태를 따른다. 먼저 Y projection으로 나누고,
    각 Y chunk를 다시 X projection으로 나눈다. X chunk가 더 나뉘면 재귀 처리하고,
    더 나뉘지 않으면 해당 chunk의 객체를 top-left 순서로 추가한다.
    """
    if len(objects) <= 1:
        return list(objects)

    y_sorted = _sort_for_axis(objects, "y")
    y_segments = _projection_segments(y_sorted, "y", min_gap=min_gap)
    if not y_segments:
        return _sort_top_left(objects)

    ordered: List[SlideObject] = []
    for y_segment in y_segments:
        y_chunk = _objects_starting_in_segment(y_sorted, "y", y_segment)
        if not y_chunk:
            continue

        x_sorted = _sort_for_axis(y_chunk, "x")
        x_segments = _projection_segments(x_sorted, "x", min_gap=min_gap)
        if not x_segments or len(x_segments) == 1:
            ordered.extend(x_sorted)
            continue

        for x_segment in x_segments:
            x_chunk = _objects_starting_in_segment(x_sorted, "x", x_segment)
            if not x_chunk:
                continue
            ordered.extend(_recursive_xycut(x_chunk, min_gap=min_gap))

    seen = {id(obj) for obj in ordered}
    missing = [obj for obj in objects if id(obj) not in seen]
    if missing:
        ordered.extend(_sort_top_left(missing))
    return ordered


def _order_objects_legacy(objects: Sequence[SlideObject]) -> List[SlideObject]:
    context = build_order_context(objects)
    tail = [o for o in objects if bucket(o, context) >= 4]
    main = [o for o in objects if bucket(o, context) < 4]

    seed = sorted(
        main,
        key=lambda o: (
            reading_top(o, context),
            object_left(o),
            o.xml_index,
        ),
    )

    rows: List[Dict[str, object]] = []
    for obj in seed:
        top = reading_top(obj, context)
        left = object_left(obj)
        height = object_height(obj)
        placed = False

        for row in rows:
            anchor_top = int(row["anchor_top"])
            row_height = int(row["max_height"])
            tolerance = max(160000, int(min(row_height, height) * 0.35))
            if abs(top - anchor_top) <= tolerance:
                row["objects"].append(obj)
                row["tops"].append(top)
                row["lefts"].append(left)
                row["anchor_top"] = min(row["tops"])
                row["max_height"] = max(row_height, height)
                placed = True
                break

        if not placed:
            rows.append(
                {
                    "objects": [obj],
                    "tops": [top],
                    "lefts": [left],
                    "anchor_top": top,
                    "max_height": height,
                }
            )

    def row_key(row: Dict[str, object]) -> Tuple[int, int, int, int]:
        row_objects = row["objects"]
        return (
            int(row["anchor_top"]),
            min(object_left(obj) for obj in row_objects),
            min(bucket(obj, context) for obj in row_objects),
            min(obj.xml_index for obj in row_objects),
        )

    ordered: List[SlideObject] = []
    for row in sorted(rows, key=row_key):
        row_objects = sorted(
            row["objects"],
            key=lambda obj: (
                bucket(obj, context),
                object_left(obj),
                reading_top(obj, context),
                obj.xml_index,
            ),
        )
        ordered.extend(row_objects)

    ordered.extend(
        sorted(
            tail,
            key=lambda o: (
                bucket(o, context),
                reading_top(o, context),
                object_left(o),
                o.xml_index,
            ),
        )
    )
    return ordered


def _order_objects_xycut(objects: Sequence[SlideObject]) -> List[SlideObject]:
    """읽기 대상 body 객체에 XY cut을 적용하고 PPTX용 tail 객체를 뒤에 붙인다."""
    context = build_order_context(objects)
    tail = [o for o in objects if bucket(o, context) >= 4]
    main = [o for o in objects if bucket(o, context) < 4]
    positioned = [o for o in main if _has_reliable_position(o)]
    unpositioned = [o for o in main if not _has_reliable_position(o)]
    if len(positioned) < 2:
        return _order_objects_legacy(objects)

    ordered_main = _recursive_xycut(positioned, min_gap=XYCUT_MIN_GAP_EMU)
    ordered_unknown = sorted(
        unpositioned,
        key=lambda o: (
            bucket(o, context),
            reading_top(o, context),
            object_left(o),
            o.xml_index,
        ),
    )
    ordered_tail = sorted(
        tail,
        key=lambda o: (
            bucket(o, context),
            reading_top(o, context),
            object_left(o),
            o.xml_index,
        ),
    )
    return ordered_main + ordered_unknown + ordered_tail


def order_objects(objects: Sequence[SlideObject], mode: str) -> List[SlideObject]:
    if mode == "xycut":
        return _order_objects_xycut(objects)
    if mode == "xml":
        return _order_objects_legacy(objects)
    raise ValueError(f"unsupported reading order mode: {mode}")


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
            if prev_heading_depth is not None and depth is not None and depth > (prev_heading_depth + 1):
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
