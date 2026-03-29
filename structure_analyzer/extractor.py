from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import xml.etree.ElementTree as ET

from .constants import (
    FOOTER_TYPES,
    LARGE_INT,
    NS,
    REORDERABLE,
    REL_NS,
    SLIDE_LAYOUT_REL_TYPE,
    STRICT_HEADING_PLACEHOLDER_TYPES,
    TITLE_TYPES,
)
from .models import SlideObject
from .text_rules import is_numbered_heading_text, normalize_text
from .xml_primitives import (
    extract_bbox_emu,
    first_off,
    get_nvpr_paths,
    local_name,
    parse_int,
)


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

    for child in list(sp_tree):
        tag = local_name(child.tag)
        if tag not in REORDERABLE:
            continue

        _, ph_path = get_nvpr_paths(tag)
        ph = child.find(ph_path, NS)
        if ph is None:
            continue
        ph_type = ph.attrib.get("type", "body")
        ph_idx = ph.attrib.get("idx", "")

        off = first_off(child)
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


def strict_heading_semantic_guard(text: str) -> bool:
    # TODO: Add sentence-ending and punctuation-density checks for strict mode.
    _ = text
    return True


def looks_heading(text: str, ph_type: Optional[str], strict: bool = False) -> bool:
    raw = (text or "").strip()
    if raw.startswith(("▶", "-", "*", "√")):
        return False

    normalized = normalize_text(text)
    if not normalized:
        return False

    if strict:
        if ph_type is None:
            return False
        if ph_type not in STRICT_HEADING_PLACEHOLDER_TYPES:
            return False
        if len(normalized) < 3 or len(normalized) > 60:
            return False
        if not strict_heading_semantic_guard(text):
            return False
        return True

    if is_numbered_heading_text(text):
        return True
    if ph_type in TITLE_TYPES and len(normalized) <= 80:
        return True
    return False


def is_decorative(tag: str, text: str) -> bool:
    normalized = normalize_text(text)
    if tag == "cxnSp":
        return True
    if not normalized and tag not in {"pic", "graphicFrame"}:
        return True
    return False


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

    for child in list(sp_tree):
        tag = local_name(child.tag)
        if tag not in REORDERABLE:
            continue
        xml_idx += 1

        c_nv_path, ph_path = get_nvpr_paths(tag)
        c_nv_pr = child.find(c_nv_path, NS)
        ph = child.find(ph_path, NS)

        shape_id = c_nv_pr.attrib.get("id", "") if c_nv_pr is not None else ""
        name = c_nv_pr.attrib.get("name", "") if c_nv_pr is not None else ""
        ph_type = ph.attrib.get("type") if ph is not None else None
        ph_idx = ph.attrib.get("idx") if ph is not None else None
        bbox = extract_bbox_emu(child)

        off = first_off(child)
        coord_source = "direct"
        if off is not None:
            x = parse_int(off.attrib.get("x"))
            y = parse_int(off.attrib.get("y"))
        elif bbox is not None:
            x = bbox[0]
            y = bbox[1]
        else:
            x = LARGE_INT
            y = LARGE_INT
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

        texts: List[str] = []
        for t in child.findall(".//a:t", NS):
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
                is_heading=looks_heading(text, ph_type, strict=strict),
                is_title_placeholder=ph_type in TITLE_TYPES,
                font_pt=extract_font_pt(child),
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

        if bbox is not None and tag == "graphicFrame" and (child.find(".//a:tbl", NS) is not None):
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
