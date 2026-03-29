from __future__ import annotations

import re
from typing import Dict, List, Sequence, Tuple

from .constants import LARGE_INT, TITLE_TYPES
from .models import OrderContext, SlideObject
from .text_rules import is_numbered_heading_text


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


def order_objects(objects: Sequence[SlideObject], mode: str) -> List[SlideObject]:
    _ = mode
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
