from __future__ import annotations

import os
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from pptx2markdown.structure_analyzer.constants import (
    FOOTER_TYPES,
    LARGE_INT,
    NS,
    REL_NS,
    REORDERABLE,
    SLIDE_LAYOUT_REL_TYPE,
    SLIDE_MASTER_REL_TYPE,
    STRICT_HEADING_PLACEHOLDER_TYPES,
    TITLE_TYPES,
)
from pptx2markdown.structure_analyzer.text_rules import is_numbered_heading_text, normalize_text
from pptx2markdown.structure_analyzer.xml_primitives import (
    contains_math,
    extract_bbox_emu,
    first_off,
    get_nvpr_paths,
    local_name,
    parse_int,
)

PLACEHOLDER_DEFAULT_TYPE = "obj"
PLACEHOLDER_DEFAULT_IDX = "0"
PLACEHOLDER_SPECIAL_IDX = str(0xFFFFFFFF)
PPTX_INHERITANCE_MODES = {"none", "geometry", "style"}
INHERITED_SHAPE_MODES = {"none", "visible", "all"}
MASTER_TITLE_TYPES = {"title", "ctrTitle", "subTitle"}
MASTER_BODY_TYPES = {"body"}


def normalize_pptx_inheritance_mode(value: str) -> str:
    mode = (value or "style").strip()
    if mode not in PPTX_INHERITANCE_MODES:
        raise ValueError(f"invalid placeholder inheritance mode: {value}")
    return mode


def normalize_inherited_shapes_mode(value: str) -> str:
    mode = (value or "visible").strip()
    if mode not in INHERITED_SHAPE_MODES:
        raise ValueError(f"invalid inherited shapes mode: {value}")
    return mode


@dataclass(frozen=True)
class EffectiveShape:
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
    list_kind: Optional[str] = None
    list_level: Optional[int] = None
    bbox: Optional[Tuple[int, int, int, int]] = None
    source_part: str = "slide"
    inheritance_kind: str = "direct"

    @property
    def has_list_semantics(self) -> bool:
        return self.list_kind in {"ul", "ol"}


@dataclass(frozen=True)
class EffectiveSlide:
    slide_xml: Path
    layout_xml: Optional[Path]
    master_xml: Optional[Path]
    pptx_inheritance: str
    inherited_shapes: str
    shapes: Tuple[EffectiveShape, ...]
    xml_tables: Tuple[Dict[str, object], ...] = field(default_factory=tuple)
    xml_images: Tuple[Dict[str, object], ...] = field(default_factory=tuple)
    layout_placeholder_count: int = 0
    layout_placeholder_ambiguous: Tuple[str, ...] = field(default_factory=tuple)
    master_placeholder_ambiguous: Tuple[str, ...] = field(default_factory=tuple)
    error: Optional[str] = None

    def meta(self) -> Dict[str, object]:
        out: Dict[str, object] = {
            "pptx_inheritance": self.pptx_inheritance,
            "inherited_shapes": self.inherited_shapes,
            "layout_xml": str(self.layout_xml) if self.layout_xml else None,
            "master_xml": str(self.master_xml) if self.master_xml else None,
            "layout_placeholder_count": self.layout_placeholder_count,
            "layout_placeholder_ambiguous": list(self.layout_placeholder_ambiguous),
            "master_placeholder_ambiguous": list(self.master_placeholder_ambiguous),
            "xml_tables": list(self.xml_tables),
            "xml_images": list(self.xml_images),
        }
        if self.error:
            out["error"] = self.error
        return out


@dataclass(frozen=True)
class PlaceholderChain:
    slide_ph: Optional[ET.Element]
    layout_ph: Optional[ET.Element]
    master_ph: Optional[ET.Element]
    layout_elem: Optional[ET.Element]
    master_elem: Optional[ET.Element]
    idx: Optional[str]


def _resolve_related_part(source_xml: Path, rel_target: str) -> Path:
    if rel_target.startswith("/"):
        parts = source_xml.resolve().parts
        try:
            ppt_idx = parts.index("ppt")
        except ValueError:
            return Path(rel_target).resolve()
        package_root = Path(*parts[:ppt_idx])
        return (package_root / rel_target.lstrip("/")).resolve()
    return Path(os.path.normpath(str(source_xml.parent / rel_target))).resolve()


def _relationship_target(source_xml: Path, rel_type: str) -> Optional[Path]:
    rels = source_xml.parent / "_rels" / f"{source_xml.name}.rels"
    if not rels.exists():
        return None

    rel_root = ET.parse(rels).getroot()
    for rel in rel_root.findall("rel:Relationship", REL_NS):
        current_type = str(rel.attrib.get("Type", ""))
        expected_suffix = "/" + rel_type.rsplit("/", 1)[-1]
        if current_type != rel_type and not current_type.endswith(expected_suffix):
            continue
        target = rel.attrib.get("Target")
        if not target:
            continue
        resolved = _resolve_related_part(source_xml, target)
        if resolved.exists():
            return resolved
    return None


def resolve_slide_layout(slide_xml: Path) -> Optional[Path]:
    return _relationship_target(slide_xml, SLIDE_LAYOUT_REL_TYPE)


def resolve_slide_master(layout_xml: Optional[Path]) -> Optional[Path]:
    if layout_xml is None:
        return None
    return _relationship_target(layout_xml, SLIDE_MASTER_REL_TYPE)


def _normalize_placeholder_idx(value: Optional[str]) -> str:
    if value is None or str(value).strip() == "":
        return PLACEHOLDER_DEFAULT_IDX
    raw = str(value).strip()
    try:
        return str(int(raw, 0))
    except ValueError:
        return raw


def _is_special_placeholder_idx(idx: Optional[str]) -> bool:
    return idx is not None and _normalize_placeholder_idx(idx) == PLACEHOLDER_SPECIAL_IDX


def _placeholder_idx(ph: Optional[ET.Element]) -> Optional[str]:
    if ph is None:
        return None
    return _normalize_placeholder_idx(ph.attrib.get("idx"))


def _effective_placeholder_attr(
    name: str,
    default: str,
    chain: PlaceholderChain,
) -> Optional[str]:
    for ph in (chain.slide_ph, chain.layout_ph, chain.master_ph):
        if ph is not None and name in ph.attrib:
            return ph.attrib[name]
    if chain.slide_ph is None:
        return None
    return default


def _effective_placeholder_type(chain: PlaceholderChain) -> Optional[str]:
    return _effective_placeholder_attr("type", PLACEHOLDER_DEFAULT_TYPE, chain)


def _build_placeholder_index(part_xml: Optional[Path]) -> Tuple[Dict[str, ET.Element], set[str]]:
    out: Dict[str, ET.Element] = {}
    ambiguous: set[str] = set()
    if part_xml is None or not part_xml.exists():
        return out, ambiguous

    root = ET.parse(part_xml).getroot()
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        return out, ambiguous

    for child in list(sp_tree):
        tag = local_name(child.tag)
        if tag not in REORDERABLE:
            continue
        _, ph_path = get_nvpr_paths(tag)
        ph = child.find(ph_path, NS)
        if ph is None:
            continue
        idx = _placeholder_idx(ph)
        if idx is None:
            continue
        if idx in out:
            ambiguous.add(idx)
            continue
        out[idx] = child
    return out, ambiguous


def _placeholder_for_idx(
    index: Dict[str, ET.Element],
    ambiguous: set[str],
    idx: Optional[str],
) -> Optional[ET.Element]:
    if idx is None or _is_special_placeholder_idx(idx) or idx in ambiguous:
        return None
    return index.get(idx)


def _placeholder_element(elem: Optional[ET.Element]) -> Optional[ET.Element]:
    if elem is None:
        return None
    _, ph_path = get_nvpr_paths(local_name(elem.tag))
    return elem.find(ph_path, NS)


def _placeholder_chain(
    slide_ph: Optional[ET.Element],
    layout_index: Dict[str, ET.Element],
    layout_ambiguous: set[str],
    master_index: Dict[str, ET.Element],
    master_ambiguous: set[str],
) -> PlaceholderChain:
    idx = _placeholder_idx(slide_ph)
    layout_elem = _placeholder_for_idx(layout_index, layout_ambiguous, idx)
    layout_ph = _placeholder_element(layout_elem)

    master_elem = _placeholder_for_idx(master_index, master_ambiguous, idx)
    master_ph = _placeholder_element(master_elem)

    return PlaceholderChain(
        slide_ph=slide_ph,
        layout_ph=layout_ph,
        master_ph=master_ph,
        layout_elem=layout_elem,
        master_elem=master_elem,
        idx=idx,
    )


def _parse_optional_int(value: Optional[str]) -> Optional[int]:
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def _is_false(value: Optional[str]) -> bool:
    return str(value or "").strip().lower() in {"0", "false"}


def _show_master_shapes(slide_xml: Path, layout_xml: Optional[Path]) -> bool:
    for part_xml in (slide_xml, layout_xml):
        if part_xml is None or not part_xml.exists():
            continue
        root = ET.parse(part_xml).getroot()
        if _is_false(root.attrib.get("showMasterSp")):
            return False
    return True


def _xfrm_components(elem: Optional[ET.Element]) -> Dict[str, Optional[int]]:
    if elem is None:
        return {"x": None, "y": None, "cx": None, "cy": None}
    off = first_off(elem)
    ext = None
    for path in (
        "./p:spPr/a:xfrm/a:ext",
        "./p:grpSpPr/a:xfrm/a:ext",
        "./p:xfrm/a:ext",
        ".//a:ext",
    ):
        ext = elem.find(path, NS)
        if ext is not None:
            break
    cx = _parse_optional_int(ext.attrib.get("cx")) if ext is not None else None
    cy = _parse_optional_int(ext.attrib.get("cy")) if ext is not None else None
    return {
        "x": _parse_optional_int(off.attrib.get("x")) if off is not None else None,
        "y": _parse_optional_int(off.attrib.get("y")) if off is not None else None,
        "cx": cx if cx is not None and cx > 0 else None,
        "cy": cy if cy is not None and cy > 0 else None,
    }


def _resolve_effective_bbox(
    slide_elem: ET.Element,
    chain: PlaceholderChain,
) -> Tuple[Optional[Tuple[int, int, int, int]], Optional[str]]:
    sources = [
        ("direct", _xfrm_components(slide_elem)),
        ("layout", _xfrm_components(chain.layout_elem)),
        ("master", _xfrm_components(chain.master_elem)),
    ]
    values: Dict[str, int] = {}
    used_sources: set[str] = set()
    for key in ("x", "y", "cx", "cy"):
        for source, comps in sources:
            value = comps.get(key)
            if value is None:
                continue
            values[key] = value
            used_sources.add(source)
            break

    if set(values) != {"x", "y", "cx", "cy"}:
        return None, None
    if values["cx"] <= 0 or values["cy"] <= 0:
        return None, None

    bbox = (
        values["x"],
        values["y"],
        values["x"] + values["cx"],
        values["y"] + values["cy"],
    )
    if used_sources == {"direct"}:
        return bbox, None
    if "master" in used_sources:
        return bbox, "master"
    if "layout" in used_sources:
        return bbox, "layout"
    return bbox, None


def extract_font_pt(elem: Optional[ET.Element]) -> Optional[float]:
    if elem is None:
        return None
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


def _extract_font_pt_from_chain(
    slide_elem: ET.Element,
    chain: PlaceholderChain,
) -> Optional[float]:
    for elem in (slide_elem, chain.layout_elem, chain.master_elem):
        font_pt = extract_font_pt(elem)
        if font_pt is not None:
            return font_pt
    return None


def _paragraph_level(elem: ET.Element) -> int:
    for p_pr in elem.findall(".//a:pPr", NS):
        raw = p_pr.attrib.get("lvl")
        level = _parse_optional_int(raw)
        if level is not None:
            return max(0, min(level, 8))
    return 0


def _list_semantics_from_ppr(p_pr: Optional[ET.Element]) -> Tuple[Optional[str], Optional[int]]:
    if p_pr is None:
        return None, None
    raw_level = _parse_optional_int(p_pr.attrib.get("lvl"))
    level = max(0, min(raw_level, 8)) if raw_level is not None else None
    if p_pr.find("a:buNone", NS) is not None:
        return None, level
    if p_pr.find("a:buAutoNum", NS) is not None:
        return "ol", level
    if p_pr.find("a:buChar", NS) is not None:
        return "ul", level
    return None, None


def _extract_list_semantics_from_shape(
    elem: Optional[ET.Element],
) -> Tuple[Optional[str], Optional[int]]:
    if elem is None:
        return None, None
    for p_pr in elem.findall(".//a:pPr", NS):
        kind, level = _list_semantics_from_ppr(p_pr)
        if kind is not None:
            return kind, 0 if level is None else level
    return None, None


def _tx_style_name(ph_type: Optional[str]) -> str:
    if ph_type in MASTER_TITLE_TYPES:
        return "titleStyle"
    if ph_type in MASTER_BODY_TYPES:
        return "bodyStyle"
    return "otherStyle"


def _extract_master_tx_style_font_pt(
    master_xml: Optional[Path],
    ph_type: Optional[str],
    paragraph_level: int,
) -> Optional[float]:
    if master_xml is None or not master_xml.exists():
        return None
    root = ET.parse(master_xml).getroot()
    style = root.find(f"p:txStyles/p:{_tx_style_name(ph_type)}", NS)
    if style is None:
        return None
    level = max(0, min(paragraph_level, 8)) + 1
    candidates = [
        style.find(f"a:lvl{level}pPr/a:defRPr", NS),
        style.find("a:defPPr/a:defRPr", NS),
    ]
    for def_rpr in candidates:
        if def_rpr is None:
            continue
        sz = def_rpr.attrib.get("sz")
        if sz is None:
            continue
        value = _parse_optional_int(sz)
        if value is not None and value > 0:
            return value / 100.0
    return None


def _master_tx_style_level_ppr(
    master_xml: Optional[Path],
    ph_type: Optional[str],
    paragraph_level: int,
) -> Optional[ET.Element]:
    if master_xml is None or not master_xml.exists():
        return None
    root = ET.parse(master_xml).getroot()
    style = root.find(f"p:txStyles/p:{_tx_style_name(ph_type)}", NS)
    if style is None:
        return None
    level = max(0, min(paragraph_level, 8)) + 1
    return style.find(f"a:lvl{level}pPr", NS)


def _extract_semantic_font_pt(
    slide_elem: ET.Element,
    chain: PlaceholderChain,
    ph_type: Optional[str],
    master_xml: Optional[Path],
) -> Optional[float]:
    font_pt = _extract_font_pt_from_chain(slide_elem, chain)
    if font_pt is not None:
        return font_pt
    return _extract_master_tx_style_font_pt(master_xml, ph_type, _paragraph_level(slide_elem))


def _extract_semantic_list_semantics(
    slide_elem: ET.Element,
    chain: PlaceholderChain,
    ph_type: Optional[str],
    master_xml: Optional[Path],
) -> Tuple[Optional[str], Optional[int]]:
    for elem in (slide_elem, chain.layout_elem, chain.master_elem):
        kind, level = _extract_list_semantics_from_shape(elem)
        if kind is not None:
            return kind, 0 if level is None else level
    paragraph_level = _paragraph_level(slide_elem)
    kind, level = _list_semantics_from_ppr(
        _master_tx_style_level_ppr(master_xml, ph_type, paragraph_level)
    )
    if kind is not None:
        return kind, 0 if level is None else level
    return None, None


def _shape_text(elem: ET.Element) -> Tuple[str, str]:
    texts: List[str] = []
    for t in elem.findall(".//a:t", NS):
        if t.text and t.text.strip():
            texts.append(t.text.strip())
    text = " ".join(texts)
    return text, normalize_text(text)


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


def _has_embedded_image(elem: ET.Element) -> bool:
    blip = elem.find(".//a:blip", NS)
    return blip is not None and bool(blip.attrib.get(f"{{{NS['r']}}}embed"))


def is_decorative(
    tag: str,
    text: str,
    elem: Optional[ET.Element] = None,
    ph_type: Optional[str] = None,
) -> bool:
    if tag == "cxnSp":
        return True
    if tag == "pic" and ph_type is not None and (elem is None or not _has_embedded_image(elem)):
        return True
    if not (text or "").strip() and not contains_math(elem) and tag not in {"pic", "graphicFrame"}:
        return True
    return False


def _is_visible_materialized_shape(elem: ET.Element, tag: str, ph_type: Optional[str]) -> bool:
    text, _ = _shape_text(elem)
    if ph_type is not None:
        return False
    if text.strip():
        return True
    if tag == "pic":
        bbox = extract_bbox_emu(elem)
        if bbox is None:
            return True
        width = bbox[2] - bbox[0]
        height = bbox[3] - bbox[1]
        longer_side = max(width, height)
        shorter_side = min(width, height)
        if longer_side > 0 and shorter_side / longer_side < 0.03:
            return False
        return True
    if tag == "graphicFrame" and elem.find(".//a:tbl", NS) is not None:
        return True
    return False


def _inheritance_kind(source_part: str, coord_source: str) -> str:
    if source_part != "slide":
        return "materialized"
    if coord_source in {"layout", "master"}:
        return "placeholder"
    return "direct"


def _shape_from_part(
    child: ET.Element,
    *,
    source_part: str,
    xml_index: int,
    ph_type: Optional[str],
    ph_idx: Optional[str],
    bbox: Optional[Tuple[int, int, int, int]],
    coord_source: str,
    font_pt: Optional[float],
    list_kind: Optional[str],
    list_level: Optional[int],
    strict: bool,
    text_override: Optional[str] = None,
) -> EffectiveShape:
    tag = local_name(child.tag)
    c_nv_path, _ = get_nvpr_paths(tag)
    c_nv_pr = child.find(c_nv_path, NS)
    if bbox is None:
        x = LARGE_INT
        y = LARGE_INT
    else:
        x = bbox[0]
        y = bbox[1]
    text, normalized = _shape_text(child)
    if text_override is not None:
        text = text_override
        normalized = normalize_text(text_override)
    shape_id = c_nv_pr.attrib.get("id", str(xml_index)) if c_nv_pr is not None else str(xml_index)
    name = c_nv_pr.attrib.get("name", "") if c_nv_pr is not None else ""
    return EffectiveShape(
        shape_id=(shape_id if source_part == "slide" else f"{source_part}:{shape_id}"),
        xml_index=xml_index,
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
        is_decorative=is_decorative(tag, text, child, ph_type),
        is_heading=looks_heading(text, ph_type, strict=strict),
        is_title_placeholder=ph_type in TITLE_TYPES,
        font_pt=font_pt,
        list_kind=list_kind,
        list_level=list_level,
        bbox=bbox,
        source_part=source_part,
        inheritance_kind=_inheritance_kind(source_part, coord_source),
    )


def _materialize_part_shapes(
    part_xml: Optional[Path],
    source_part: str,
    start_xml_index: int,
    used_placeholder_indexes: set[str],
    inherited_shapes: str,
    strict: bool,
) -> List[EffectiveShape]:
    if part_xml is None or not part_xml.exists() or inherited_shapes == "none":
        return []
    root = ET.parse(part_xml).getroot()
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        return []
    shapes: List[EffectiveShape] = []
    local_index = 0
    for child in list(sp_tree):
        tag = local_name(child.tag)
        if tag not in REORDERABLE:
            continue
        local_index += 1
        _, ph_path = get_nvpr_paths(tag)
        ph = child.find(ph_path, NS)
        ph_idx = _placeholder_idx(ph)
        if ph_idx is not None and ph_idx in used_placeholder_indexes:
            continue
        ph_type = ph.attrib.get("type", PLACEHOLDER_DEFAULT_TYPE) if ph is not None else None
        if inherited_shapes == "visible" and not _is_visible_materialized_shape(
            child,
            tag,
            ph_type,
        ):
            continue
        bbox = extract_bbox_emu(child)
        list_kind, list_level = _extract_list_semantics_from_shape(child)
        shapes.append(
            _shape_from_part(
                child,
                source_part=source_part,
                xml_index=start_xml_index + local_index,
                ph_type=ph_type,
                ph_idx=ph_idx,
                bbox=bbox,
                coord_source=source_part,
                font_pt=extract_font_pt(child),
                list_kind=list_kind,
                list_level=list_level,
                strict=strict,
            )
        )
    return shapes


def resolve_effective_slide(
    slide_xml: Path,
    *,
    strict: bool = False,
    pptx_inheritance: str = "style",
    inherited_shapes: str = "visible",
) -> EffectiveSlide:
    pptx_inheritance = normalize_pptx_inheritance_mode(pptx_inheritance)
    inherited_shapes = normalize_inherited_shapes_mode(inherited_shapes)

    root = ET.parse(slide_xml).getroot()
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        return EffectiveSlide(
            slide_xml=slide_xml,
            layout_xml=None,
            master_xml=None,
            pptx_inheritance=pptx_inheritance,
            inherited_shapes=inherited_shapes,
            shapes=tuple(),
            error="Missing p:cSld/p:spTree",
        )

    use_placeholder_inheritance = pptx_inheritance in {"geometry", "style"}
    use_semantic_inheritance = pptx_inheritance == "style"
    layout_xml = resolve_slide_layout(slide_xml) if use_placeholder_inheritance else None
    master_xml = resolve_slide_master(layout_xml) if use_placeholder_inheritance else None
    layout_index, layout_ambiguous = _build_placeholder_index(layout_xml)
    master_index, master_ambiguous = _build_placeholder_index(master_xml)

    shapes: List[EffectiveShape] = []
    xml_tables: List[Dict[str, object]] = []
    xml_images: List[Dict[str, object]] = []
    used_placeholder_indexes: set[str] = set()
    xml_idx = 0

    for child in list(sp_tree):
        tag = local_name(child.tag)
        if tag not in REORDERABLE:
            continue
        xml_idx += 1
        _, ph_path = get_nvpr_paths(tag)
        ph = child.find(ph_path, NS)

        if use_placeholder_inheritance:
            chain = _placeholder_chain(
                ph,
                layout_index,
                layout_ambiguous,
                master_index,
                master_ambiguous,
            )
            ph_type = _effective_placeholder_type(chain)
            ph_idx = chain.idx
            bbox, inherited_coord_source = _resolve_effective_bbox(child, chain)
        else:
            chain = PlaceholderChain(ph, None, None, None, None, _placeholder_idx(ph))
            ph_type = ph.attrib.get("type", PLACEHOLDER_DEFAULT_TYPE) if ph is not None else None
            ph_idx = chain.idx
            bbox = extract_bbox_emu(child)
            inherited_coord_source = None
        if ph_idx is not None:
            used_placeholder_indexes.add(ph_idx)

        coord_source = inherited_coord_source or "direct"
        if use_semantic_inheritance:
            font_pt = _extract_semantic_font_pt(child, chain, ph_type, master_xml)
            list_kind, list_level = _extract_semantic_list_semantics(
                child,
                chain,
                ph_type,
                master_xml,
            )
        else:
            font_pt = extract_font_pt(child)
            list_kind, list_level = _extract_list_semantics_from_shape(child)
        shape = _shape_from_part(
            child,
            source_part="slide",
            xml_index=xml_idx,
            ph_type=ph_type,
            ph_idx=ph_idx,
            bbox=bbox,
            coord_source=coord_source,
            font_pt=font_pt,
            list_kind=list_kind,
            list_level=list_level,
            strict=strict,
        )
        shapes.append(shape)

        if bbox is not None and tag == "pic":
            xml_images.append(
                {
                    "shape_id": shape.shape_id,
                    "name": shape.name,
                    "bbox": list(bbox),
                    "tag": tag,
                }
            )
        if bbox is not None and tag == "graphicFrame" and (child.find(".//a:tbl", NS) is not None):
            xml_tables.append(
                {
                    "shape_id": shape.shape_id,
                    "name": shape.name,
                    "bbox": list(bbox),
                    "tag": tag,
                }
            )

    if use_placeholder_inheritance and inherited_shapes != "none":
        layout_shapes = _materialize_part_shapes(
            layout_xml,
            "layout",
            1_000_000,
            used_placeholder_indexes,
            inherited_shapes,
            strict,
        )
        shapes.extend(layout_shapes)
        used_placeholder_indexes.update(
            shape.ph_idx for shape in layout_shapes if shape.ph_idx is not None
        )
        if _show_master_shapes(slide_xml, layout_xml):
            shapes.extend(
                _materialize_part_shapes(
                    master_xml,
                    "master",
                    2_000_000,
                    used_placeholder_indexes,
                    inherited_shapes,
                    strict,
                )
            )

    return EffectiveSlide(
        slide_xml=slide_xml,
        layout_xml=layout_xml,
        master_xml=master_xml,
        pptx_inheritance=pptx_inheritance,
        inherited_shapes=inherited_shapes,
        shapes=tuple(shapes),
        xml_tables=tuple(xml_tables),
        xml_images=tuple(xml_images),
        layout_placeholder_count=len(layout_index),
        layout_placeholder_ambiguous=tuple(sorted(layout_ambiguous)),
        master_placeholder_ambiguous=tuple(sorted(master_ambiguous)),
    )
