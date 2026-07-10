from __future__ import annotations

import os
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

from pptx2markdown.pptx_inheritance.resolver import (
    EffectiveShape,
    normalize_inherited_shapes_mode,
    normalize_pptx_inheritance_mode,
    resolve_effective_slide,
)

from .constants import (
    FOOTER_TYPES,
    LARGE_INT,
    NS,
    REL_NS,
    REORDERABLE,
    SLIDE_LAYOUT_REL_TYPE,
    STRICT_HEADING_PLACEHOLDER_TYPES,
    TITLE_TYPES,
)
from .structure import SlideObject
from .text_rules import is_numbered_heading_text, normalize_text
from .xml_primitives import (
    contains_math,
    first_off,
    get_nvpr_paths,
    local_name,
    parse_int,
)

LEAF_DRAWABLE_TAGS = {"sp", "pic", "graphicFrame", "cxnSp"}


@dataclass(frozen=True)
class _GroupTransform:
    sx: float = 1.0
    sy: float = 1.0
    tx: float = 0.0
    ty: float = 0.0

    def apply_bbox(self, bbox: Tuple[int, int, int, int]) -> Tuple[int, int, int, int]:
        x1, y1, x2, y2 = bbox
        return (
            int(round(self.sx * x1 + self.tx)),
            int(round(self.sy * y1 + self.ty)),
            int(round(self.sx * x2 + self.tx)),
            int(round(self.sy * y2 + self.ty)),
        )

    def compose(self, inner: "_GroupTransform") -> "_GroupTransform":
        return _GroupTransform(
            sx=self.sx * inner.sx,
            sy=self.sy * inner.sy,
            tx=self.sx * inner.tx + self.tx,
            ty=self.sy * inner.ty + self.ty,
        )


@dataclass(frozen=True)
class _FlattenedShape:
    elem: ET.Element
    tag: str
    shape_id: str
    name: str
    group_path: Tuple[str, ...]
    z_path: Tuple[int, ...]
    bbox: Optional[Tuple[int, int, int, int]]


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


def parse_layout_placeholders(
    layout_xml: Optional[Path],
) -> Dict[Tuple[str, str], Tuple[int, int]]:
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
        return True

    if is_numbered_heading_text(text):
        return True
    if ph_type in TITLE_TYPES and len(normalized) <= 80:
        return True
    return False


def is_decorative(tag: str, text: str, elem: Optional[ET.Element] = None) -> bool:
    normalized = normalize_text(text)
    if tag == "cxnSp":
        return True
    if not normalized and not contains_math(elem) and tag not in {"pic", "graphicFrame"}:
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


def _first(elem: ET.Element, paths: Tuple[str, ...]) -> Optional[ET.Element]:
    for path in paths:
        found = elem.find(path, NS)
        if found is not None:
            return found
    return None


def _point_attrs(
    node: Optional[ET.Element], x_name: str, y_name: str
) -> Optional[Tuple[int, int]]:
    if node is None:
        return None
    x = parse_int(node.attrib.get(x_name))
    y = parse_int(node.attrib.get(y_name))
    if x >= LARGE_INT or y >= LARGE_INT:
        return None
    return x, y


def _shape_id_name(elem: ET.Element, tag: str) -> Tuple[str, str]:
    c_nv_path, _ = get_nvpr_paths(tag)
    c_nv_pr = elem.find(c_nv_path, NS)
    if c_nv_pr is None:
        return "", ""
    return c_nv_pr.attrib.get("id", ""), c_nv_pr.attrib.get("name", "")


def _extract_local_bbox_emu(elem: ET.Element) -> Optional[Tuple[int, int, int, int]]:
    off = _first(
        elem,
        (
            "./p:spPr/a:xfrm/a:off",
            "./p:grpSpPr/a:xfrm/a:off",
            "./p:xfrm/a:off",
        ),
    )
    ext = _first(
        elem,
        (
            "./p:spPr/a:xfrm/a:ext",
            "./p:grpSpPr/a:xfrm/a:ext",
            "./p:xfrm/a:ext",
        ),
    )
    origin = _point_attrs(off, "x", "y")
    size = _point_attrs(ext, "cx", "cy")
    if origin is None or size is None:
        return None
    x, y = origin
    w, h = size
    if w <= 0 or h <= 0:
        return None
    return (x, y, x + w, y + h)


def _group_child_transform(group: ET.Element) -> _GroupTransform:
    xfrm = group.find("./p:grpSpPr/a:xfrm", NS)
    if xfrm is None:
        return _GroupTransform()

    off = _point_attrs(xfrm.find("./a:off", NS), "x", "y")
    ext = _point_attrs(xfrm.find("./a:ext", NS), "cx", "cy")
    ch_off = _point_attrs(xfrm.find("./a:chOff", NS), "x", "y")
    ch_ext = _point_attrs(xfrm.find("./a:chExt", NS), "cx", "cy")
    if off is None or ext is None or ch_off is None or ch_ext is None:
        return _GroupTransform()

    off_x, off_y = off
    ext_x, ext_y = ext
    ch_off_x, ch_off_y = ch_off
    ch_ext_x, ch_ext_y = ch_ext
    if ch_ext_x <= 0 or ch_ext_y <= 0:
        return _GroupTransform()

    sx = ext_x / ch_ext_x
    sy = ext_y / ch_ext_y
    return _GroupTransform(
        sx=sx,
        sy=sy,
        tx=off_x - sx * ch_off_x,
        ty=off_y - sy * ch_off_y,
    )


def _expanded_children(elem: ET.Element) -> Iterable[ET.Element]:
    for child in list(elem):
        tag = local_name(child.tag)
        if tag != "AlternateContent":
            yield child
            continue

        selected = child.find("./{*}Choice")
        if selected is None:
            selected = child.find("./{*}Fallback")
        if selected is None:
            continue
        yield from list(selected)


def _iter_flattened_shapes(container: ET.Element) -> Iterable[_FlattenedShape]:
    def walk(
        elem: ET.Element,
        transform: _GroupTransform,
        group_path: Tuple[str, ...],
        z_prefix: Tuple[int, ...],
    ) -> Iterable[_FlattenedShape]:
        ordinal = 0
        for child in _expanded_children(elem):
            tag = local_name(child.tag)
            if tag not in REORDERABLE:
                continue

            ordinal += 1
            z_path = z_prefix + (ordinal,)
            shape_id, name = _shape_id_name(child, tag)
            local_bbox = _extract_local_bbox_emu(child)
            bbox = transform.apply_bbox(local_bbox) if local_bbox is not None else None

            if tag == "grpSp":
                group_key = shape_id or ".".join(str(part) for part in z_path)
                group_transform = _group_child_transform(child)
                yield from walk(
                    child,
                    transform.compose(group_transform),
                    group_path + (group_key,),
                    z_path,
                )
                continue

            if tag in LEAF_DRAWABLE_TAGS:
                yield _FlattenedShape(
                    elem=child,
                    tag=tag,
                    shape_id=shape_id,
                    name=name,
                    group_path=group_path,
                    z_path=z_path,
                    bbox=bbox,
                )

    yield from walk(container, _GroupTransform(), (), ())


def _slide_object_from_effective(shape: EffectiveShape) -> SlideObject:
    return SlideObject(
        shape_id=shape.shape_id,
        xml_index=shape.xml_index,
        tag=shape.tag,
        name=shape.name,
        ph_type=shape.ph_type,
        ph_idx=shape.ph_idx,
        x=shape.x,
        y=shape.y,
        coord_source=shape.coord_source,
        text=shape.text,
        normalized=shape.normalized,
        is_footer=shape.is_footer,
        is_decorative=shape.is_decorative,
        is_heading=shape.is_heading,
        is_title_placeholder=shape.is_title_placeholder,
        font_pt=shape.font_pt,
        list_kind=shape.list_kind,
        list_level=shape.list_level,
        bbox=shape.bbox,
        source_part=shape.source_part,
        inheritance_kind=shape.inheritance_kind,
    )


def extract_slide_objects_xml(
    slide_xml: Path,
    strict: bool = False,
    pptx_inheritance: str = "style",
    inherited_shapes: str = "visible",
) -> Tuple[List[SlideObject], Dict[str, object]]:
    pptx_inheritance = normalize_pptx_inheritance_mode(pptx_inheritance)
    inherited_shapes = normalize_inherited_shapes_mode(inherited_shapes)

    root = ET.parse(slide_xml).getroot()
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        return [], {"error": "Missing p:cSld/p:spTree"}

    effective_slide = resolve_effective_slide(
        slide_xml,
        strict=strict,
        pptx_inheritance=pptx_inheritance,
        inherited_shapes=inherited_shapes,
    )
    effective_by_shape_id = {
        shape.shape_id: shape for shape in effective_slide.shapes if shape.source_part == "slide"
    }

    layout_xml = resolve_slide_layout(slide_xml)
    layout_map = parse_layout_placeholders(layout_xml)

    objects: List[SlideObject] = []
    xml_tables: List[Dict[str, object]] = []
    xml_images: List[Dict[str, object]] = []
    xml_idx = 0
    group_count = len(sp_tree.findall(".//p:grpSp", NS))

    for item in _iter_flattened_shapes(sp_tree):
        child = item.elem
        tag = item.tag
        xml_idx += 1

        _, ph_path = get_nvpr_paths(tag)
        ph = child.find(ph_path, NS)

        shape_id = item.shape_id
        name = item.name
        ph_type = ph.attrib.get("type") if ph is not None else None
        ph_idx = ph.attrib.get("idx") if ph is not None else None
        bbox = item.bbox
        effective = effective_by_shape_id.get(shape_id)
        if effective is not None:
            ph_type = effective.ph_type
            ph_idx = effective.ph_idx
            if bbox is None:
                bbox = effective.bbox

        off = first_off(child)
        coord_source = "direct"
        if bbox is not None:
            x = bbox[0]
            y = bbox[1]
            if effective is not None and item.bbox is None:
                coord_source = effective.coord_source
        elif off is not None:
            x = parse_int(off.attrib.get("x"))
            y = parse_int(off.attrib.get("y"))
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

        text = normalize_text("".join(t.text or "" for t in child.findall(".//a:t", NS)))
        normalized = text
        font_pt = extract_font_pt(child)
        list_kind = None
        list_level = None
        is_decorative_value = is_decorative(tag, text, child)
        is_heading_value = looks_heading(text, ph_type, strict=strict)
        source_part = "slide"
        inheritance_kind = "placeholder" if coord_source in {"layout", "master"} else "direct"
        if effective is not None:
            text = effective.text
            normalized = effective.normalized
            font_pt = effective.font_pt
            list_kind = effective.list_kind
            list_level = effective.list_level
            is_decorative_value = effective.is_decorative
            is_heading_value = effective.is_heading
            source_part = effective.source_part
            inheritance_kind = effective.inheritance_kind

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
                is_decorative=is_decorative_value,
                is_heading=is_heading_value,
                is_title_placeholder=ph_type in TITLE_TYPES,
                font_pt=font_pt,
                list_kind=list_kind,
                list_level=list_level,
                bbox=bbox,
                group_path=item.group_path,
                z_path=item.z_path,
                source_part=source_part,
                inheritance_kind=inheritance_kind,
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

    for shape in effective_slide.shapes:
        if shape.inheritance_kind == "materialized":
            objects.append(_slide_object_from_effective(shape))

    meta = effective_slide.meta()
    meta["xml_tables"] = xml_tables
    meta["xml_images"] = xml_images
    meta["group_count"] = group_count
    meta["flattened_groups"] = group_count > 0
    return objects, meta
