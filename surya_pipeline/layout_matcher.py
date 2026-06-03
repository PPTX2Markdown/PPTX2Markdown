from __future__ import annotations

import json
import math
import re
import zipfile
from copy import deepcopy
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as ET


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"p": P_NS, "a": A_NS, "r": R_NS}
EMU_PER_INCH = 914400
SURYA_HEADING_TOP_K_THRESHOLD = 0.25


@dataclass(frozen=True)
class SlideSize:
    cx: int
    cy: int


@dataclass(frozen=True)
class BBox:
    x: float
    y: float
    w: float
    h: float

    @property
    def x2(self) -> float:
        return self.x + self.w

    @property
    def y2(self) -> float:
        return self.y + self.h

    @property
    def area(self) -> float:
        return max(0.0, self.w) * max(0.0, self.h)

    @property
    def center(self) -> tuple[float, float]:
        return (self.x + self.w / 2.0, self.y + self.h / 2.0)

    def to_list(self) -> list[float]:
        return [self.x, self.y, self.w, self.h]


@dataclass
class XmlObject:
    xml_index: int
    shape_id: str
    tag: str
    name: str
    ph_type: str | None
    ph_idx: str | None
    bbox: BBox | None
    bbox_source: str
    bbox_source_part: str | None
    text: str
    font_pt: float | None

    @property
    def is_textual(self) -> bool:
        return self.tag in {"sp", "grpSp"} and bool(self.text.strip())

    @property
    def is_title_placeholder(self) -> bool:
        return (self.ph_type or "").lower() in {"title", "ctrtitle"}

    @property
    def is_subtitle_placeholder(self) -> bool:
        return (self.ph_type or "").lower() == "subtitle"

    @property
    def is_footer(self) -> bool:
        ph = (self.ph_type or "").lower()
        name = self.name.lower()
        return ph in {"ftr", "sldnum", "dt"} or "footer" in name

    @property
    def is_decorative(self) -> bool:
        if self.text.strip():
            return False
        if self.tag == "cxnSp":
            return True
        if self.bbox is not None and (self.bbox.w <= 0 or self.bbox.h <= 0):
            return True
        return False


@dataclass
class LayoutBlock:
    block_id: str
    page: int
    model_position: int
    label: str
    confidence: float | None
    top_k: dict[str, float]
    bbox_px: BBox
    bbox_emu: BBox


@dataclass
class MatchCandidate:
    shape_id: str
    score: float
    overlap_min: float
    layout_coverage: float
    object_coverage: float
    center_score: float
    compatibility: float
    reason: str


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def load_presentation_root(*, ppt_root: Path | None = None, pptx_path: Path | None = None) -> tuple[Path, Any]:
    if ppt_root is not None:
        return ppt_root, None
    if pptx_path is None:
        raise ValueError("Either --ppt-root or --pptx-path is required.")
    extract_root = pptx_path.with_suffix("")
    if not extract_root.exists():
        with zipfile.ZipFile(pptx_path) as zf:
            zf.extractall(extract_root)
    return extract_root, None


def read_slide_size(ppt_root: Path) -> SlideSize:
    pres_xml = ppt_root / "ppt" / "presentation.xml"
    root = ET.parse(pres_xml).getroot()
    sld_sz = root.find("p:sldSz", NS)
    if sld_sz is None:
        return SlideSize(cx=10 * EMU_PER_INCH, cy=int(7.5 * EMU_PER_INCH))
    return SlideSize(cx=int(sld_sz.get("cx", "0")), cy=int(sld_sz.get("cy", "0")))


def slide_xml_paths(ppt_root: Path) -> list[Path]:
    return sorted((ppt_root / "ppt" / "slides").glob("slide*.xml"), key=_slide_number)


def _slide_number(path: Path) -> int:
    digits = "".join(ch for ch in path.stem if ch.isdigit())
    return int(digits or 0)


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _first_text(node: ET.Element) -> str:
    return " ".join(t.text.strip() for t in node.findall(".//a:t", NS) if t.text and t.text.strip())


def _shape_id(node: ET.Element) -> str:
    cnv = node.find(".//p:cNvPr", NS)
    return (cnv.get("id") if cnv is not None else "") or ""


def _shape_name(node: ET.Element) -> str:
    cnv = node.find(".//p:cNvPr", NS)
    return (cnv.get("name") if cnv is not None else "") or ""


def _placeholder(node: ET.Element) -> tuple[str | None, str | None]:
    ph = node.find(".//p:ph", NS)
    if ph is None:
        return None, None
    return ph.get("type"), ph.get("idx")


def _bbox_from_node(node: ET.Element) -> BBox | None:
    xfrm = None
    for path in (
        "./p:xfrm",
        "./p:spPr/a:xfrm",
        "./p:grpSpPr/a:xfrm",
        "./a:xfrm",
        ".//p:xfrm",
        ".//a:xfrm",
    ):
        xfrm = node.find(path, NS)
        if xfrm is not None:
            break
    if xfrm is None:
        return None
    off = xfrm.find("a:off", NS)
    ext = xfrm.find("a:ext", NS)
    if off is None or ext is None:
        return None
    try:
        return BBox(
            x=float(off.get("x", "0")),
            y=float(off.get("y", "0")),
            w=float(ext.get("cx", "0")),
            h=float(ext.get("cy", "0")),
        )
    except ValueError:
        return None


def _font_pt(node: ET.Element) -> float | None:
    values: list[float] = []
    for rpr in node.findall(".//a:rPr", NS):
        sz = rpr.get("sz")
        if sz:
            try:
                values.append(float(sz) / 100.0)
            except ValueError:
                pass
    if values:
        return max(values)
    return None


def parse_slide_xml_objects(slide_xml: Path) -> list[XmlObject]:
    root = ET.parse(slide_xml).getroot()
    sp_tree = root.find(".//p:spTree", NS)
    if sp_tree is None:
        return []
    placeholder_boxes = _placeholder_bbox_lookup(slide_xml)
    objects: list[XmlObject] = []
    for idx, child in enumerate(list(sp_tree)):
        tag = _local_name(child.tag)
        if tag not in {"sp", "pic", "graphicFrame", "grpSp", "cxnSp"}:
            continue
        shape_id = _shape_id(child)
        if not shape_id:
            continue
        ph_type, ph_idx = _placeholder(child)
        bbox = _bbox_from_node(child)
        bbox_source = "slide_xml" if bbox is not None else "missing"
        bbox_source_part: str | None = str(slide_xml) if bbox is not None else None
        if bbox is None:
            inherited = _lookup_placeholder_bbox(placeholder_boxes, ph_type, ph_idx)
            if inherited is not None:
                bbox, bbox_source_part = inherited
                bbox_source = "placeholder_inheritance"
        objects.append(
            XmlObject(
                xml_index=idx,
                shape_id=shape_id,
                tag=tag,
                name=_shape_name(child),
                ph_type=ph_type,
                ph_idx=ph_idx,
                bbox=bbox,
                bbox_source=bbox_source,
                bbox_source_part=bbox_source_part,
                text=_first_text(child),
                font_pt=_font_pt(child),
            )
        )
    return objects


def _placeholder_bbox_lookup(slide_xml: Path) -> dict[tuple[str | None, str | None], tuple[BBox, str]]:
    lookup: dict[tuple[str | None, str | None], tuple[BBox, str]] = {}
    layout_xml = _related_part(slide_xml, "slideLayout")
    if layout_xml is not None:
        lookup.update(_placeholder_boxes_in_part(layout_xml))
        master_xml = _related_part(layout_xml, "slideMaster")
        if master_xml is not None:
            master_lookup = _placeholder_boxes_in_part(master_xml)
            master_lookup.update(lookup)
            lookup = master_lookup
    return lookup


def _placeholder_boxes_in_part(xml_path: Path) -> dict[tuple[str | None, str | None], tuple[BBox, str]]:
    if not xml_path.exists():
        return {}
    root = ET.parse(xml_path).getroot()
    out: dict[tuple[str | None, str | None], tuple[BBox, str]] = {}
    for shape in root.findall(".//p:sp", NS):
        ph_type, ph_idx = _placeholder(shape)
        if ph_type is None and ph_idx is None:
            continue
        bbox = _bbox_from_node(shape)
        if bbox is None:
            continue
        value = (bbox, str(xml_path))
        out[(ph_type, ph_idx)] = value
        out.setdefault((ph_type, None), value)
        out.setdefault((None, ph_idx), value)
    return out


def _lookup_placeholder_bbox(
    lookup: dict[tuple[str | None, str | None], tuple[BBox, str]],
    ph_type: str | None,
    ph_idx: str | None,
) -> tuple[BBox, str] | None:
    for key in ((ph_type, ph_idx), (ph_type, None), (None, ph_idx)):
        if key in lookup:
            return lookup[key]
    return None


def _related_part(xml_path: Path, relationship_name: str) -> Path | None:
    rels = xml_path.parent / "_rels" / f"{xml_path.name}.rels"
    if not rels.exists():
        return None
    root = ET.parse(rels).getroot()
    for rel in root:
        rel_type = rel.get("Type", "")
        if not rel_type.endswith(f"/{relationship_name}"):
            continue
        target = rel.get("Target")
        if not target:
            continue
        return (xml_path.parent / target).resolve()
    return None


def parse_surya_blocks(layout_json: dict[str, Any], slide_size: SlideSize) -> list[LayoutBlock]:
    pages = _layout_pages(layout_json)
    blocks: list[LayoutBlock] = []
    for page_index, page in enumerate(pages, start=1):
        image_bbox = _bbox_from_raw(page.get("image_bbox")) or _fallback_image_bbox(page, slide_size)
        for pos, raw in enumerate(_raw_blocks(page), start=1):
            bbox_px = _bbox_from_raw(raw.get("bbox"))
            if bbox_px is None:
                continue
            bbox_emu = _scale_bbox(bbox_px, image_bbox, slide_size)
            blocks.append(
                LayoutBlock(
                    block_id=f"p{page_index}_b{pos}",
                    page=page_index,
                    model_position=pos,
                    label=str(raw.get("label") or raw.get("type") or ""),
                    confidence=_safe_float(raw.get("confidence") or raw.get("score")),
                    top_k=_top_k_from_raw(raw.get("top_k")),
                    bbox_px=bbox_px,
                    bbox_emu=bbox_emu,
                )
            )
    return blocks


def _layout_pages(payload: dict[str, Any]) -> list[dict[str, Any]]:
    if isinstance(payload.get("pages"), list):
        return [p for p in payload["pages"] if isinstance(p, dict)]
    if isinstance(payload.get("layout"), list):
        return [p for p in payload["layout"] if isinstance(p, dict)]
    children = payload.get("children")
    if isinstance(children, list):
        return [p for p in children if isinstance(p, dict)]
    if len(payload) == 1:
        only_value = next(iter(payload.values()))
        if isinstance(only_value, list):
            return [p for p in only_value if isinstance(p, dict)]
    return [payload]


def _raw_blocks(page: dict[str, Any]) -> list[dict[str, Any]]:
    for key in ("layout", "bboxes", "blocks", "children"):
        value = page.get(key)
        if isinstance(value, list):
            return [item for item in value if isinstance(item, dict)]
    return []


def _fallback_image_bbox(page: dict[str, Any], slide_size: SlideSize) -> BBox:
    width = _safe_float(page.get("width")) or float(slide_size.cx)
    height = _safe_float(page.get("height")) or float(slide_size.cy)
    return BBox(0.0, 0.0, width, height)


def _bbox_from_raw(raw: Any) -> BBox | None:
    if isinstance(raw, dict):
        vals = [raw.get("x"), raw.get("y"), raw.get("w") or raw.get("width"), raw.get("h") or raw.get("height")]
        if all(v is not None for v in vals):
            return BBox(float(vals[0]), float(vals[1]), float(vals[2]), float(vals[3]))
        vals = [raw.get("x1"), raw.get("y1"), raw.get("x2"), raw.get("y2")]
        if all(v is not None for v in vals):
            return BBox(float(vals[0]), float(vals[1]), float(vals[2]) - float(vals[0]), float(vals[3]) - float(vals[1]))
    if isinstance(raw, list) and len(raw) >= 4:
        x1, y1, x2, y2 = [float(v) for v in raw[:4]]
        return BBox(x1, y1, x2 - x1, y2 - y1)
    return None


def _bbox_from_xywh(raw: Any) -> BBox | None:
    if isinstance(raw, list) and len(raw) >= 4:
        return BBox(float(raw[0]), float(raw[1]), float(raw[2]), float(raw[3]))
    if isinstance(raw, dict):
        vals = [raw.get("x"), raw.get("y"), raw.get("w") or raw.get("width"), raw.get("h") or raw.get("height")]
        if all(v is not None for v in vals):
            return BBox(float(vals[0]), float(vals[1]), float(vals[2]), float(vals[3]))
    return None


def _scale_bbox(bbox: BBox, image_bbox: BBox, slide_size: SlideSize) -> BBox:
    sx = slide_size.cx / image_bbox.w if image_bbox.w else 1.0
    sy = slide_size.cy / image_bbox.h if image_bbox.h else 1.0
    return BBox(
        x=(bbox.x - image_bbox.x) * sx,
        y=(bbox.y - image_bbox.y) * sy,
        w=bbox.w * sx,
        h=bbox.h * sy,
    )


def _safe_float(value: Any) -> float | None:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _top_k_from_raw(value: Any) -> dict[str, float]:
    if not isinstance(value, dict):
        return {}
    out: dict[str, float] = {}
    for key, raw_score in value.items():
        score = _safe_float(raw_score)
        if score is not None:
            out[str(key)] = score
    return out


def build_normalized_pages(layout_json: dict[str, Any], ppt_root: Path) -> list[dict[str, Any]]:
    slide_size = read_slide_size(ppt_root)
    blocks = parse_surya_blocks(layout_json, slide_size)
    blocks_by_page: dict[int, list[LayoutBlock]] = {}
    for block in blocks:
        blocks_by_page.setdefault(block.page, []).append(block)

    pages: list[dict[str, Any]] = []
    for page_num, slide_xml in enumerate(slide_xml_paths(ppt_root), start=1):
        objects = parse_slide_xml_objects(slide_xml)
        page_blocks = blocks_by_page.get(page_num, [])
        reading_order, unmatched_blocks, decorative_objects = match_page(page_num, page_blocks, objects, slide_size)
        match_candidates = build_match_candidates(page_num, page_blocks, objects)
        pages.append(
            {
                "page": page_num,
                "slide_xml": str(slide_xml),
                "slide_size_emu": asdict(slide_size),
                "layout_blocks": [_block_dict(block) for block in page_blocks],
                "xml_objects": [_object_dict(obj) for obj in objects],
                "match_candidates": match_candidates,
                "reading_order": reading_order,
                "unmatched_layout_blocks": unmatched_blocks,
                "decorative_objects": [_object_dict(obj) for obj in decorative_objects],
            }
        )
    return pages


def build_xml_bbox_pages(ppt_root: Path) -> list[dict[str, Any]]:
    slide_size = read_slide_size(ppt_root)
    pages: list[dict[str, Any]] = []
    for page_num, slide_xml in enumerate(slide_xml_paths(ppt_root), start=1):
        objects = parse_slide_xml_objects(slide_xml)
        pages.append(
            {
                "page": page_num,
                "slide_xml": str(slide_xml),
                "slide_size_emu": asdict(slide_size),
                "objects": [_object_dict(obj) for obj in objects],
            }
        )
    return pages


def build_surya_bbox_pages(layout_json: dict[str, Any], ppt_root: Path) -> list[dict[str, Any]]:
    slide_size = read_slide_size(ppt_root)
    blocks = parse_surya_blocks(layout_json, slide_size)
    blocks_by_page: dict[int, list[LayoutBlock]] = {}
    for block in blocks:
        blocks_by_page.setdefault(block.page, []).append(block)
    page_count = max(len(slide_xml_paths(ppt_root)), max(blocks_by_page, default=0))
    return [
        {
            "page": page_num,
            "slide_size_emu": asdict(slide_size),
            "blocks": [_block_dict(block) for block in blocks_by_page.get(page_num, [])],
        }
        for page_num in range(1, page_count + 1)
    ]


def build_match_candidate_pages(
    xml_bbox_pages: list[dict[str, Any]],
    surya_bbox_pages: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    pages: list[dict[str, Any]] = []
    xml_by_page = {int(page["page"]): page for page in xml_bbox_pages}
    surya_by_page = {int(page["page"]): page for page in surya_bbox_pages}
    for page_num in sorted(set(xml_by_page) | set(surya_by_page)):
        objects = [_xml_object_from_dict(item) for item in xml_by_page.get(page_num, {}).get("objects", [])]
        blocks = [_layout_block_from_dict(item) for item in surya_by_page.get(page_num, {}).get("blocks", [])]
        pages.append(
            {
                "page": page_num,
                "candidates": build_match_candidates(page_num, blocks, objects),
            }
        )
    return pages


def build_match_candidates(page_num: int, blocks: list[LayoutBlock], objects: list[XmlObject]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    matchable_objects = [obj for obj in objects if not obj.is_decorative]
    for block in sorted(blocks, key=lambda b: b.model_position):
        candidates = sorted(
            (_candidate(block, obj) for obj in matchable_objects if obj.bbox is not None),
            key=lambda c: c.score,
            reverse=True,
        )
        rows.append(
            {
                "page": page_num,
                "layout_block": _block_dict(block),
                "candidates": [
                    _candidate_dict(candidate, _find_object(matchable_objects, candidate.shape_id))
                    for candidate in candidates
                ],
            }
        )
    return rows


def match_page(
    page_num: int,
    blocks: list[LayoutBlock],
    objects: list[XmlObject],
    slide_size: SlideSize | None = None,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[XmlObject]]:
    assigned: set[str] = set()
    rows: list[dict[str, Any]] = []
    unmatched_blocks: list[dict[str, Any]] = []
    decorative_objects = [obj for obj in objects if obj.is_decorative]
    matchable_objects = [obj for obj in objects if not obj.is_decorative]

    for block in sorted(blocks, key=lambda b: b.model_position):
        role = _label_role(block.label, block.top_k)
        candidates = sorted(
            (_candidate(block, obj) for obj in matchable_objects if obj.shape_id not in assigned and obj.bbox is not None),
            key=lambda c: c.score,
            reverse=True,
        )
        selected: list[MatchCandidate] = []
        if role in {"figure", "picture"}:
            selected = [c for c in candidates if c.score >= 0.35 and c.object_coverage >= 0.35]
            if not selected and candidates and candidates[0].score >= 0.50:
                selected = [candidates[0]]
        elif candidates:
            threshold = 0.62 if role in {"table", "footer"} else 0.48
            if candidates[0].score >= threshold:
                selected = [candidates[0]]

        if not selected:
            unmatched_blocks.append(_block_dict(block))
            continue

        selected.sort(key=lambda c: _object_sort_key(_find_object(matchable_objects, c.shape_id)))
        for candidate in selected:
            obj = _find_object(matchable_objects, candidate.shape_id)
            if obj is None or obj.shape_id in assigned:
                continue
            assigned.add(obj.shape_id)
            rows.append(_reading_row(page_num, block, obj, candidate, "surya_region" if len(selected) > 1 else "surya_match"))

    fallback_objects = [obj for obj in matchable_objects if obj.shape_id not in assigned]
    for obj in sorted(fallback_objects, key=_object_sort_key):
        rows.append(_reading_row(page_num, None, obj, None, "xml_append"))

    for order_index, row in enumerate(rows, start=1):
        row["order_index"] = order_index
    _assign_surya_heading_depths(rows, blocks, slide_size)
    return rows, unmatched_blocks, decorative_objects


def _candidate(block: LayoutBlock, obj: XmlObject) -> MatchCandidate:
    assert obj.bbox is not None
    overlap_min, layout_coverage, object_coverage = _overlap_features(block.bbox_emu, obj.bbox)
    center = _center_score(block.bbox_emu, obj.bbox)
    compatibility, reason = _compatibility(block, obj)
    score = (0.35 * overlap_min) + (0.25 * layout_coverage) + (0.15 * object_coverage) + (0.15 * center) + compatibility
    return MatchCandidate(
        shape_id=obj.shape_id,
        score=max(0.0, min(1.0, score)),
        overlap_min=overlap_min,
        layout_coverage=layout_coverage,
        object_coverage=object_coverage,
        center_score=center,
        compatibility=compatibility,
        reason=reason,
    )


def _overlap_features(a: BBox, b: BBox) -> tuple[float, float, float]:
    ix = max(0.0, min(a.x2, b.x2) - max(a.x, b.x))
    iy = max(0.0, min(a.y2, b.y2) - max(a.y, b.y))
    inter = ix * iy
    if inter <= 0:
        return 0.0, 0.0, 0.0
    overlap_min = inter / max(1.0, min(a.area, b.area))
    layout_coverage = inter / max(1.0, a.area)
    object_coverage = inter / max(1.0, b.area)
    return overlap_min, layout_coverage, object_coverage


def _center_score(a: BBox, b: BBox) -> float:
    ax, ay = a.center
    bx, by = b.center
    distance = math.hypot(ax - bx, ay - by)
    diagonal = math.hypot(max(a.w, b.w), max(a.h, b.h))
    if diagonal <= 0:
        return 0.0
    return max(0.0, 1.0 - (distance / diagonal))


def _compatibility(block: LayoutBlock, obj: XmlObject) -> tuple[float, str]:
    role = _label_role(block.label, block.top_k)
    if role == "table":
        if obj.tag == "graphicFrame":
            return 0.25, "table_graphicFrame"
        if obj.is_textual:
            return -0.25, "table_text_penalty"
        return -0.15, "table_non_table_penalty"
    if role == "footer":
        if obj.is_footer:
            return 0.30, "footer_placeholder"
        return -0.40, "footer_requires_footer"
    if role == "heading":
        if obj.is_title_placeholder:
            return 0.25, "heading_title_placeholder"
        if obj.is_textual:
            bonus = 0.08
            reason = "heading_text"
            if obj.font_pt is not None and obj.font_pt >= 28.0:
                bonus += 0.18
                reason = "heading_text_large_font"
            if _looks_like_numbered_heading_text(obj.text):
                bonus += 0.08
                reason = f"{reason}_numbered"
            return bonus, reason
        return -0.35, "heading_non_text_penalty"
    if role in {"text", "list"}:
        if obj.is_textual:
            return 0.12, "text_shape"
        return -0.25, "text_non_text_penalty"
    if role in {"figure", "picture"}:
        if obj.tag in {"pic", "grpSp"}:
            return 0.18, "figure_visual"
        if obj.tag == "graphicFrame":
            return 0.08, "figure_table_region"
        if obj.is_textual:
            return 0.00, "figure_text_member"
        return -0.05, "figure_weak_member"
    return 0.0, "neutral"


def _label_role(label: str, top_k: dict[str, float] | None = None) -> str:
    key = label.replace(" ", "").replace("_", "").lower()
    if "sectionheader" in key or "title" in key or "header" in key:
        return "heading"
    if top_k and _top_k_heading_score(top_k) >= SURYA_HEADING_TOP_K_THRESHOLD:
        return "heading"
    if "table" in key:
        return "table"
    if "footer" in key or "pagenumber" in key:
        return "footer"
    if "list" in key:
        return "list"
    if "figure" in key:
        return "figure"
    if "picture" in key or "image" in key:
        return "picture"
    if "text" in key:
        return "text"
    return "unknown"


def _top_k_heading_score(top_k: dict[str, float] | None) -> float:
    if not top_k:
        return 0.0
    return max(
        (
            float(score)
            for label, score in top_k.items()
            if "sectionheader" in label.replace(" ", "").replace("_", "").lower()
            or "header" in label.replace(" ", "").replace("_", "").lower()
            or "title" in label.replace(" ", "").replace("_", "").lower()
        ),
        default=0.0,
    )


def _looks_like_numbered_heading_text(text: str) -> bool:
    stripped = " ".join((text or "").split())
    if len(stripped) > 90:
        return False
    return bool(re.match(r"^\d+(?:\.\d+)*\.?\s+", stripped))


def _object_sort_key(obj: XmlObject | None) -> tuple[float, float, int]:
    if obj is None or obj.bbox is None:
        return (float("inf"), float("inf"), 0)
    return (round(obj.bbox.y / 100000.0), obj.bbox.x, obj.xml_index)


def _find_object(objects: Iterable[XmlObject], shape_id: str) -> XmlObject | None:
    for obj in objects:
        if obj.shape_id == shape_id:
            return obj
    return None


def _assign_surya_heading_depths(
    rows: list[dict[str, Any]],
    blocks: list[LayoutBlock],
    slide_size: SlideSize | None = None,
) -> None:
    if not rows:
        return
    slide_height = float(slide_size.cy) if slide_size is not None else max((block.bbox_emu.y2 for block in blocks), default=1.0)
    slide_width = float(slide_size.cx) if slide_size is not None else max((block.bbox_emu.x2 for block in blocks), default=1.0)
    heading_rows = [
        row
        for row in rows
        if isinstance(row.get("heading_sources"), list)
        and "surya_label" in row.get("heading_sources", [])
        and row.get("reading_order_source") in {"surya_match", "surya_region"}
        and isinstance(row.get("bbox"), list)
        and len(row.get("bbox") or []) >= 4
    ]
    if not heading_rows:
        return

    max_height = max(float(row["bbox"][3]) for row in heading_rows) or 1.0
    for row in heading_rows:
        x, y, w, h = [float(value) for value in row["bbox"][:4]]
        normalized_height = min(1.0, h / max_height)
        normalized_width = min(1.0, w / max(1.0, slide_width))
        topness = max(0.0, min(1.0, 1.0 - (y / max(1.0, slide_height))))
        score = (0.45 * normalized_height) + (0.30 * normalized_width) + (0.25 * topness)
        row["surya_heading_primary_score"] = score

    primary = max(heading_rows, key=lambda row: float(row.get("surya_heading_primary_score") or 0.0))
    for row in heading_rows:
        row["surya_heading_depth_hint"] = 1 if row is primary else 2


def _reading_row(
    page_num: int,
    block: LayoutBlock | None,
    obj: XmlObject,
    candidate: MatchCandidate | None,
    source: str,
) -> dict[str, Any]:
    role = _label_role(block.label, block.top_k) if block else ""
    has_text = bool(obj.text.strip())
    is_heading = bool(has_text and (obj.is_title_placeholder or obj.is_subtitle_placeholder or role == "heading"))
    heading_sources: list[str] = []
    if is_heading and (obj.is_title_placeholder or obj.is_subtitle_placeholder):
        heading_sources.append("placeholder")
    if is_heading and role == "heading":
        heading_sources.append("surya_label")
    if has_text and obj.is_title_placeholder:
        depth = 1
    elif has_text and (obj.is_subtitle_placeholder or role == "heading"):
        depth = 2
    else:
        depth = None
    row = {
        "page": page_num,
        "shape_id": obj.shape_id,
        "matched_shape_id": obj.shape_id,
        "xml_index": obj.xml_index,
        "tag": obj.tag,
        "name": obj.name,
        "ph_type": obj.ph_type,
        "ph_idx": obj.ph_idx,
        "text": obj.text,
        "xml_text": obj.text,
        "font_pt": obj.font_pt,
        "bbox": obj.bbox.to_list() if obj.bbox else None,
        "reading_order_source": source,
        "layout_block_id": block.block_id if block else None,
        "position": block.model_position if block else None,
        "model_position": block.model_position if block else None,
        "label": block.label if block else None,
        "semantic_role": role or None,
        "confidence": block.confidence if block else None,
        "match_score": candidate.score if candidate else 0.0,
        "match_quality": _match_quality(candidate.score if candidate else 0.0, source),
        "match_reason": candidate.reason if candidate else "xml_spatial_fallback",
        "is_title_placeholder": obj.is_title_placeholder,
        "is_heading_candidate": is_heading,
        "heading_score": 1.0 if is_heading and obj.is_title_placeholder else (0.80 if is_heading else 0.0),
        "heading_depth_hint": depth,
        "surya_heading_depth_hint": None,
        "heading_sources": heading_sources,
        "heading_source": "placeholder" if is_heading and (obj.is_title_placeholder or obj.is_subtitle_placeholder) else ("surya_label" if is_heading else None),
    }
    if candidate:
        row.update(
            {
                "overlap_min": candidate.overlap_min,
                "layout_coverage": candidate.layout_coverage,
                "object_coverage": candidate.object_coverage,
                "center_score": candidate.center_score,
            }
        )
    return row


def _match_quality(score: float, source: str) -> str:
    if source == "xml_append":
        return "unmatched"
    if score >= 0.75:
        return "strong"
    if score >= 0.55:
        return "medium"
    return "weak"


def _block_dict(block: LayoutBlock) -> dict[str, Any]:
    return {
        "block_id": block.block_id,
        "page": block.page,
        "model_position": block.model_position,
        "label": block.label,
        "confidence": block.confidence,
        "top_k": dict(block.top_k),
        "effective_role": _label_role(block.label, block.top_k),
        "surya_heading_top_k_score": _top_k_heading_score(block.top_k),
        "bbox_px": block.bbox_px.to_list(),
        "bbox_emu": block.bbox_emu.to_list(),
    }


def _layout_block_from_dict(data: dict[str, Any]) -> LayoutBlock:
    return LayoutBlock(
        block_id=str(data.get("block_id") or ""),
        page=int(data.get("page") or 0),
        model_position=int(data.get("model_position") or 0),
        label=str(data.get("label") or ""),
        confidence=_safe_float(data.get("confidence")),
        top_k=_top_k_from_raw(data.get("top_k")),
        bbox_px=_bbox_from_xywh(data.get("bbox_px")) or BBox(0, 0, 0, 0),
        bbox_emu=_bbox_from_xywh(data.get("bbox_emu")) or BBox(0, 0, 0, 0),
    )


def _object_dict(obj: XmlObject) -> dict[str, Any]:
    data = asdict(obj)
    data["bbox"] = obj.bbox.to_list() if obj.bbox else None
    data["is_title_placeholder"] = obj.is_title_placeholder
    data["is_subtitle_placeholder"] = obj.is_subtitle_placeholder
    data["is_footer"] = obj.is_footer
    data["is_decorative"] = obj.is_decorative
    return data


def _xml_object_from_dict(data: dict[str, Any]) -> XmlObject:
    return XmlObject(
        xml_index=int(data.get("xml_index") or 0),
        shape_id=str(data.get("shape_id") or ""),
        tag=str(data.get("tag") or ""),
        name=str(data.get("name") or ""),
        ph_type=data.get("ph_type"),
        ph_idx=data.get("ph_idx"),
        bbox=_bbox_from_xywh(data.get("bbox")),
        bbox_source=str(data.get("bbox_source") or "missing"),
        bbox_source_part=data.get("bbox_source_part"),
        text=str(data.get("text") or ""),
        font_pt=_safe_float(data.get("font_pt")),
    )


def _candidate_dict(candidate: MatchCandidate, obj: XmlObject | None) -> dict[str, Any]:
    return {
        "shape_id": candidate.shape_id,
        "score": candidate.score,
        "overlap_min": candidate.overlap_min,
        "layout_coverage": candidate.layout_coverage,
        "object_coverage": candidate.object_coverage,
        "center_score": candidate.center_score,
        "compatibility": candidate.compatibility,
        "reason": candidate.reason,
        "xml_object": _object_dict(obj) if obj is not None else None,
    }


def reorder_slide_xml(slide_xml: Path, shape_order: list[str], output_xml: Path) -> None:
    tree = ET.parse(slide_xml)
    root = tree.getroot()
    sp_tree = root.find(".//p:spTree", NS)
    if sp_tree is None:
        output_xml.parent.mkdir(parents=True, exist_ok=True)
        tree.write(output_xml, encoding="utf-8", xml_declaration=True)
        return

    children = list(sp_tree)
    order = {shape_id: idx for idx, shape_id in enumerate(shape_order)}

    def key(child: ET.Element) -> tuple[int, int]:
        sid = _shape_id(child)
        if sid in order:
            return (0, order[sid])
        return (1, children.index(child))

    sorted_children = sorted(children, key=key)
    for child in children:
        sp_tree.remove(child)
    for child in sorted_children:
        sp_tree.append(deepcopy(child))

    output_xml.parent.mkdir(parents=True, exist_ok=True)
    tree.write(output_xml, encoding="utf-8", xml_declaration=True)
