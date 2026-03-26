from __future__ import annotations

import statistics
from typing import Dict, List, Optional, Sequence

from .models import OrderContext, SlideObject
from .ordering import is_top_title_object
from .text_rules import (
    is_numbered_heading_text,
    numbered_heading_kind,
    numbered_suggested_depth,
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
