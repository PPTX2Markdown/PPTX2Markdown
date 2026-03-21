#!/usr/bin/env python3
"""
Extract reading order per slide XML and write:
1) reading-order JSON
2) reordered slide XML

Usage:
  python extract_structure_analysis.py [slide1.xml slide2.xml ...]

If no positional args are given, all *.xml in ./target_slides are processed.
Outputs are always written to ./output (created automatically if missing).
"""

from __future__ import annotations

import argparse
import json
import statistics
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple
import xml.etree.ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

REL_NS = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}

SLIDE_LAYOUT_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
)

REORDERABLE = {"sp", "pic", "graphicFrame", "grpSp", "cxnSp"}
FOOTER_TYPES = {"sldNum", "ftr", "dt"}
TITLE_TYPES = {"title", "ctrTitle", "subTitle"}
STRICT_HEADING_PLACEHOLDER_TYPES = {"title", "ctrTitle", "subTitle"}


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


def local_name(tag: str) -> str:
    return tag.split("}", 1)[-1]


def natural_key(path: Path) -> Tuple:
    parts = re.split(r"(\d+)", path.name)
    out: List[object] = []
    for part in parts:
        if part.isdigit():
            out.append(int(part))
        else:
            out.append(part.lower())
    return tuple(out)


def parse_int(v: Optional[str], default: int = 10**18) -> int:
    if v is None:
        return default
    try:
        return int(v)
    except ValueError:
        return default


def normalize_text(s: str) -> str:
    s = s.lower().replace("\n", " ")
    s = re.sub(r"[^0-9a-z\uac00-\ud7a3\.\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def token_set(s: str) -> set:
    return set(re.findall(r"[0-9a-z\uac00-\ud7a3]+", normalize_text(s)))


def is_numbered_heading_text(text: str) -> bool:
    raw = re.sub(r"\s+", " ", (text or "").strip())
    t = normalize_text(text)
    # 1. / 1.1 / 1.1.1
    if re.match(r"^(?:\d+\.\d+(?:\.\d+)*\.?|\d+\.)\s+", t):
        return True
    return False


def strict_heading_semantic_guard(text: str) -> bool:
    # TODO: Add sentence-ending and punctuation-density checks for strict mode.
    _ = text
    return True


def strip_leading_heading_markers(text: str) -> str:
    raw = re.sub(r"\s+", " ", (text or "").strip())
    if not raw:
        return ""
    # Remove common bullet/prefix markers before depth inference.
    return re.sub(r"^(?:[-*•▶√]+\s*)+", "", raw).strip()


def numbered_suggested_depth(
    text: str,
    last_section_depth: Optional[int],
    seen_title: bool,
) -> Optional[int]:
    raw = strip_leading_heading_markers(text)
    if not raw:
        return None
    if re.match(r"^\d+\.\d+(?:\.\d+)*\.?\s+", raw):
        if last_section_depth is not None:
            return min(last_section_depth + 1, 3)
        return 2 if seen_title else 1
    if re.match(r"^\d+\.\s+", raw):
        return 2
    return None


def extract_font_pt(elem: ET.Element) -> Optional[float]:
    sizes: List[float] = []
    for rpr in elem.findall(".//a:rPr", NS):
        sz = rpr.attrib.get("sz")
        if sz is None:
            continue
        val = parse_int(sz, -1)
        if val > 0:
            sizes.append(val / 100.0)
    for rpr in elem.findall(".//a:endParaRPr", NS):
        sz = rpr.attrib.get("sz")
        if sz is None:
            continue
        val = parse_int(sz, -1)
        if val > 0:
            sizes.append(val / 100.0)
    if not sizes:
        return None
    return max(sizes)


def get_nvpr_paths(tag: str) -> Tuple[str, str]:
    if tag == "sp":
        return "./p:nvSpPr/p:cNvPr", "./p:nvSpPr/p:nvPr/p:ph"
    if tag == "pic":
        return "./p:nvPicPr/p:cNvPr", "./p:nvPicPr/p:nvPr/p:ph"
    if tag == "graphicFrame":
        return "./p:nvGraphicFramePr/p:cNvPr", "./p:nvGraphicFramePr/p:nvPr/p:ph"
    if tag == "grpSp":
        return "./p:nvGrpSpPr/p:cNvPr", "./p:nvGrpSpPr/p:nvPr/p:ph"
    if tag == "cxnSp":
        return "./p:nvCxnSpPr/p:cNvPr", "./p:nvCxnSpPr/p:nvPr/p:ph"
    return ".//p:cNvPr", ".//p:ph"


def first_off(elem: ET.Element) -> Optional[ET.Element]:
    for p in (
        "./p:spPr/a:xfrm/a:off",
        "./p:grpSpPr/a:xfrm/a:off",
        "./p:xfrm/a:off",
        ".//a:off",
    ):
        off = elem.find(p, NS)
        if off is not None:
            return off
    return None


def first_ext(elem: ET.Element) -> Optional[ET.Element]:
    for p in (
        "./p:spPr/a:xfrm/a:ext",
        "./p:grpSpPr/a:xfrm/a:ext",
        "./p:xfrm/a:ext",
        ".//a:ext",
    ):
        ext = elem.find(p, NS)
        if ext is not None:
            return ext
    return None


def extract_bbox_emu(elem: ET.Element) -> Optional[Tuple[int, int, int, int]]:
    off = first_off(elem)
    ext = first_ext(elem)
    if off is None or ext is None:
        return None
    x = parse_int(off.attrib.get("x"))
    y = parse_int(off.attrib.get("y"))
    w = parse_int(ext.attrib.get("cx"))
    h = parse_int(ext.attrib.get("cy"))
    if any(v >= 10**18 for v in (x, y, w, h)):
        return None
    if w <= 0 or h <= 0:
        return None
    return (x, y, x + w, y + h)


def resolve_slide_layout(slide_xml: Path) -> Optional[Path]:
    rels = slide_xml.parent / "_rels" / f"{slide_xml.name}.rels"
    if not rels.exists():
        return None

    rel_root = ET.parse(rels).getroot()
    for rel in rel_root.findall("rel:Relationship", REL_NS):
        if rel.attrib.get("Type") != SLIDE_LAYOUT_REL_TYPE:
            continue
        target = rel.attrib.get("Target")
        if not target:
            continue
        resolved = Path(os.path.normpath(str(slide_xml.parent / target)))
        if resolved.exists():
            return resolved
    return None


def parse_layout_placeholders(layout_xml: Optional[Path]) -> Dict[Tuple[str, str], Tuple[int, int]]:
    out: Dict[Tuple[str, str], Tuple[int, int]] = {}
    if layout_xml is None or not layout_xml.exists():
        return out

    root = ET.parse(layout_xml).getroot()
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        return out

    for ch in list(sp_tree):
        tag = local_name(ch.tag)
        if tag not in REORDERABLE:
            continue

        _, ph_path = get_nvpr_paths(tag)
        ph = ch.find(ph_path, NS)
        if ph is None:
            continue
        ph_type = ph.attrib.get("type", "body")
        ph_idx = ph.attrib.get("idx", "")

        off = first_off(ch)
        if off is None:
            continue
        x = parse_int(off.attrib.get("x"))
        y = parse_int(off.attrib.get("y"))

        key = (ph_type, ph_idx)
        if key not in out:
            out[key] = (x, y)
        key_type = (ph_type, "")
        if key_type not in out:
            out[key_type] = (x, y)

    return out


def looks_heading(text: str, ph_type: Optional[str], tag: str, strict: bool = False) -> bool:
    raw = (text or "").strip()
    if raw.startswith(("▶", "-", "*", "√")):
        return False

    t = normalize_text(text)
    if not t:
        return False
    if strict:
        # Strict mode requires placeholder-backed text for heading candidacy.
        if ph_type is None:
            return False
        if ph_type not in STRICT_HEADING_PLACEHOLDER_TYPES:
            return False
        if len(t) < 3 or len(t) > 60:
            return False
        if not strict_heading_semantic_guard(text):
            return False
        return True

    if is_numbered_heading_text(text):
        return True
    if ph_type in TITLE_TYPES and len(t) <= 80:
        return True
    return False


def is_decorative(tag: str, text: str) -> bool:
    t = normalize_text(text)
    if tag == "cxnSp":
        return True
    if not t and tag not in {"pic", "graphicFrame"}:
        return True
    return False


def object_left(obj: SlideObject) -> int:
    if obj.bbox is not None:
        return obj.bbox[0]
    return obj.x


def object_top(obj: SlideObject) -> int:
    if obj.bbox is not None:
        return obj.bbox[1]
    return obj.y


def reading_top(obj: SlideObject, context: OrderContext) -> int:
    top = object_top(obj)
    if top >= 10**18 and is_top_title_object(obj, context):
        return context.slide_top - max(250000, int(context.slide_height * 0.08))
    return top


def object_right(obj: SlideObject) -> int:
    if obj.bbox is not None:
        return obj.bbox[2]
    width = max(500000, min(2500000, 80000 * max(1, len(obj.normalized))))
    return object_left(obj) + width


def object_bottom(obj: SlideObject) -> int:
    if obj.bbox is not None:
        return obj.bbox[3]
    return object_top(obj) + object_height(obj)


def object_width(obj: SlideObject) -> int:
    return max(1, object_right(obj) - object_left(obj))


def object_height(obj: SlideObject) -> int:
    if obj.bbox is not None:
        return max(1, obj.bbox[3] - obj.bbox[1])
    if obj.tag == "pic":
        return 1800000
    if obj.tag == "graphicFrame":
        return 1200000
    lines = max(1, len(re.findall(r"[.!?]|[\u3002]", obj.text)) + 1)
    return max(250000, min(1800000, 250000 * lines))


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


def numbered_heading_kind(text: str) -> Optional[str]:
    t = normalize_text(text)
    if re.match(r"^\d+\.\d+(?:\.\d+)*\.?\s+", t):
        return "dotted-multi"
    if re.match(r"^\d+\.\s+", t):
        return "dotted-single"
    return None


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

    # Build font-size bands (desc). Same/close sizes are considered one local level.
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
            # Not strict: font-size takes precedence for local heading depth.
            font_depth = depth_from_font(obj.font_pt)
            if (
                font_depth is not None
                and prev_heading_depth is not None
                and obj.font_pt is not None
                and prev_heading_font is not None
                and abs(float(obj.font_pt) - float(prev_heading_font)) <= font_tolerance
            ):
                # When sizes are effectively equal, preserve local sequence depth.
                depth = prev_heading_depth
            elif font_depth is not None:
                depth = font_depth
            else:
                # Fallback to legacy sequence heuristics when font signal is unavailable.
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
                # First numbered heading starts at most from H2.
                if not seen_numbered_heading:
                    suggested_depth = min(suggested_depth, 2)
                if depth is None:
                    depth = suggested_depth
                else:
                    # Numbered pattern is a soft hint in not-strict mode.
                    depth = max(depth, suggested_depth)
                seen_numbered_heading = True

            # Keep local heading tree stable: only first top-level heading stays H1.
            if h1_assigned and depth == 1:
                depth = 2
            # Prevent abrupt depth jumps like H1 -> H3.
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
    # Confidence-like score for heading candidacy.
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


def confidence(objs: Sequence[SlideObject]) -> str:
    if not objs:
        return "low"
    unknown = sum(1 for o in objs if o.coord_source == "unknown")
    ratio = unknown / len(objs)
    if ratio >= 0.3:
        return "low"
    if ratio >= 0.1:
        return "medium"
    return "high"


def object_to_dict(
    o: SlideObject,
    context: OrderContext,
    heading_depths: Optional[Dict[str, Optional[int]]] = None,
    strict: bool = False,
) -> Dict[str, object]:
    if heading_depths is None:
        depth = compute_heading_depths([o], context, strict=strict).get(o.shape_id)
    else:
        depth = heading_depths.get(o.shape_id)
    score = heading_score(o, strict=strict)
    is_candidate = depth is not None and score >= heading_threshold(strict=strict)
    return {
        "shape_id": o.shape_id,
        "xml_index": o.xml_index,
        "tag": o.tag,
        "name": o.name,
        "ph_type": o.ph_type,
        "ph_idx": o.ph_idx,
        "x": o.x,
        "y": o.y,
        "coord_source": o.coord_source,
        "text": o.text,
        "font_pt": o.font_pt,
        "is_footer": o.is_footer,
        "is_decorative": o.is_decorative,
        "is_heading": o.is_heading,
        "is_title_placeholder": o.is_title_placeholder,
        "is_heading_candidate": is_candidate,
        "heading_score": round(score, 3),
        "heading_depth_hint": depth,
        "bbox": list(o.bbox) if o.bbox is not None else None,
        "bucket": bucket(o, context),
        "reason": reason(o, context),
    }


def extract_slide_objects_xml(slide_xml: Path, strict: bool = False) -> Tuple[List[SlideObject], Dict[str, object]]:
    root = ET.parse(slide_xml).getroot()
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        return [], {"error": "Missing p:cSld/p:spTree"}

    layout_xml = resolve_slide_layout(slide_xml)
    layout_map = parse_layout_placeholders(layout_xml)

    objects: List[SlideObject] = []
    xml_tables: List[Dict[str, object]] = []
    xml_images: List[Dict[str, object]] = []
    xml_idx = 0
    for ch in list(sp_tree):
        tag = local_name(ch.tag)
        if tag not in REORDERABLE:
            continue
        xml_idx += 1

        c_nv_path, ph_path = get_nvpr_paths(tag)
        c_nv_pr = ch.find(c_nv_path, NS)
        ph = ch.find(ph_path, NS)

        shape_id = c_nv_pr.attrib.get("id", "") if c_nv_pr is not None else ""
        name = c_nv_pr.attrib.get("name", "") if c_nv_pr is not None else ""
        ph_type = ph.attrib.get("type") if ph is not None else None
        ph_idx = ph.attrib.get("idx") if ph is not None else None
        bbox = extract_bbox_emu(ch)

        off = first_off(ch)
        coord_source = "direct"
        if off is not None:
            x = parse_int(off.attrib.get("x"))
            y = parse_int(off.attrib.get("y"))
        elif bbox is not None:
            x = bbox[0]
            y = bbox[1]
        else:
            x = 10**18
            y = 10**18
            if ph_type is not None:
                key = (ph_type, ph_idx or "")
                if key in layout_map:
                    x, y = layout_map[key]
                    coord_source = "layout"
                elif (ph_type, "") in layout_map:
                    x, y = layout_map[(ph_type, "")]
                    coord_source = "layout"
                else:
                    coord_source = "unknown"
            else:
                coord_source = "unknown"

        texts = []
        for t in ch.findall(".//a:t", NS):
            if t.text and t.text.strip():
                texts.append(t.text.strip())
        text = " ".join(texts)
        normalized = normalize_text(text)

        objects.append(
            SlideObject(
                shape_id=shape_id,
                xml_index=xml_idx,
                tag=tag,
                name=name,
                ph_type=ph_type,
                ph_idx=ph_idx,
                x=x,
                y=y,
                coord_source=coord_source,
                text=text,
                normalized=normalized,
                is_footer=(ph_type in FOOTER_TYPES) or bool(re.fullmatch(r"\d+", normalized)),
                is_decorative=is_decorative(tag, text),
                is_heading=looks_heading(text, ph_type, tag, strict=strict),
                is_title_placeholder=ph_type in TITLE_TYPES,
                font_pt=extract_font_pt(ch),
                bbox=bbox,
            )
        )
        if bbox is not None and tag == "pic":
            xml_images.append(
                {
                    "shape_id": shape_id,
                    "name": name,
                    "bbox": list(bbox),
                    "tag": tag,
                }
            )
        if bbox is not None and tag == "graphicFrame" and (ch.find(".//a:tbl", NS) is not None):
            xml_tables.append(
                {
                    "shape_id": shape_id,
                    "name": name,
                    "bbox": list(bbox),
                    "tag": tag,
                }
            )

    meta = {
        "layout_xml": str(layout_xml) if layout_xml else None,
        "layout_placeholder_count": len(layout_map),
        "xml_tables": xml_tables,
        "xml_images": xml_images,
    }
    return objects, meta


def extract_slide_objects(
    slide_xml: Path,
    mode: str,
    strict: bool = False,
) -> Tuple[List[SlideObject], Dict[str, object]]:
    objects, meta = extract_slide_objects_xml(slide_xml, strict=strict)
    meta["mode"] = mode
    meta["strict"] = strict
    return objects, meta


def order_objects(objects: Sequence[SlideObject], mode: str) -> List[SlideObject]:
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


def reorder_tree_by_indexes(tree: ET.ElementTree, ordered_xml_indexes: Sequence[int]) -> None:
    root = tree.getroot()
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        return

    children = list(sp_tree)
    reorderables = [ch for ch in children if local_name(ch.tag) in REORDERABLE]
    if not reorderables:
        return

    idx_to_elem = {i + 1: elem for i, elem in enumerate(reorderables)}
    reordered_elems = [idx_to_elem[i] for i in ordered_xml_indexes if i in idx_to_elem]
    used = set(ordered_xml_indexes)
    for i, elem in idx_to_elem.items():
        if i not in used:
            reordered_elems.append(elem)

    new_children: List[ET.Element] = []
    inserted = False
    for ch in children:
        if local_name(ch.tag) in REORDERABLE:
            if not inserted:
                new_children.extend(reordered_elems)
                inserted = True
            continue
        new_children.append(ch)
    sp_tree[:] = new_children


def gather_input_files(target_dir: Path, raw_inputs: Sequence[str]) -> List[Path]:
    files: List[Path] = []

    if raw_inputs:
        for item in raw_inputs:
            p = Path(item)
            candidates = [p, target_dir / item]
            if p.suffix == "":
                candidates.append(target_dir / f"{item}.xml")
            selected: Optional[Path] = None
            for c in candidates:
                if c.exists() and c.is_file():
                    selected = c
                    break
            if selected is None:
                continue
            files.append(selected.resolve())
    else:
        if target_dir.exists() and not target_dir.is_dir():
            raise NotADirectoryError(f"target_slides path exists but is not a directory: {target_dir}")
        target_dir.mkdir(parents=True, exist_ok=True)
        files = sorted(target_dir.glob("*.xml"), key=natural_key)
        files = [f.resolve() for f in files if f.is_file()]

    unique: Dict[str, Path] = {}
    for f in files:
        unique[str(f)] = f
    return list(unique.values())


def write_outputs(
    slide_xml: Path,
    output_dir: Path,
    mode: str,
    strict: bool = False,
) -> Dict[str, object]:
    objects, meta = extract_slide_objects(slide_xml, mode=mode, strict=strict)
    context = build_order_context(objects)
    ordered = order_objects(objects, mode=mode)
    ordered_heading_depths = compute_heading_depths(ordered, context, strict=strict)
    raw_heading_depths = compute_heading_depths(sorted(objects, key=lambda x: x.xml_index), context, strict=strict)

    ordered_indexes = [obj.xml_index for obj in ordered]

    tree = ET.parse(slide_xml)
    reorder_tree_by_indexes(tree, ordered_indexes)

    stem = slide_xml.stem
    json_path = output_dir / f"{stem}.structure_analysis.json"
    xml_path = output_dir / f"{stem}.reordered.xml"

    report: Dict[str, object] = {
        "input_xml": str(slide_xml),
        "mode": mode,
        "strict": strict,
        "layout_xml": meta.get("layout_xml"),
        "confidence": confidence(objects),
        "counts": {
            "total": len(objects),
            "text": sum(1 for o in objects if bool(o.normalized)),
            "graphicFrame": sum(1 for o in objects if o.tag == "graphicFrame"),
            "pic": sum(1 for o in objects if o.tag == "pic"),
            "footer": sum(1 for o in objects if o.is_footer),
            "decorative": sum(1 for o in objects if o.is_decorative),
            "layout_coord_used": sum(1 for o in objects if o.coord_source == "layout"),
            "xml_tables": len(meta.get("xml_tables", [])),
            "xml_images": len(meta.get("xml_images", [])),
        },
        "xml_tables": meta.get("xml_tables", []),
        "xml_images": meta.get("xml_images", []),
        "structure_order": [object_to_dict(o, context, ordered_heading_depths, strict=strict) for o in ordered],
        "raw_xml_order": [
            object_to_dict(o, context, raw_heading_depths, strict=strict)
            for o in sorted(objects, key=lambda x: x.xml_index)
        ],
        "ordered_xml_indexes": ordered_indexes,
        "output_structure_xml": str(xml_path),
    }

    json_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    # Keep reordered XML human-readable for debugging and diffs.
    try:
        ET.indent(tree, space="  ")
    except AttributeError:
        pass
    tree.write(xml_path, encoding="utf-8", xml_declaration=True)

    return {
        "input_xml": str(slide_xml),
        "output_json": str(json_path),
        "output_xml": str(xml_path),
        "confidence": report["confidence"],
        "object_count": len(objects),
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract reading order from slide XML and write JSON + reordered XML."
    )
    parser.add_argument(
        "inputs",
        nargs="*",
        help="Input slide xml file(s). If omitted, process all *.xml from ./target_slides.",
    )
    parser.add_argument(
        "--mode",
        choices=("xml",),
        default="xml",
        help="Reading order mode. Only xml mode is supported.",
    )
    parser.add_argument(
        "--output-dir",
        default="./output",
        help="Output directory for reading-order JSON/XML artifacts.",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Use stricter xml heading detection and candidacy thresholds.",
    )
    args = parser.parse_args()

    target_dir = Path("./target_slides")
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    ET.register_namespace("a", NS["a"])
    ET.register_namespace("p", NS["p"])
    ET.register_namespace("r", NS["r"])

    inputs = gather_input_files(target_dir, args.inputs)
    if not inputs:
        print("No input XML files found.")
        print(f"Checked default directory: {target_dir.resolve()}")
        return

    manifest = {"processed": [], "failed": []}

    for slide_xml in sorted(inputs, key=natural_key):
        try:
            row = write_outputs(slide_xml, output_dir, mode=args.mode, strict=args.strict)
            manifest["processed"].append(row)
            print(f"Processed: {slide_xml.name}")
        except Exception as e:  # noqa: BLE001
            manifest["failed"].append({"input_xml": str(slide_xml), "error": str(e)})
            print(f"Failed: {slide_xml.name} -> {e}")

    manifest_path = output_dir / "structure_analysis_manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Wrote manifest: {manifest_path.resolve()}")
    print(f"Output directory: {output_dir.resolve()}")
    print(f"Processed: {len(manifest['processed'])}, Failed: {len(manifest['failed'])}")


if __name__ == "__main__":
    main()
