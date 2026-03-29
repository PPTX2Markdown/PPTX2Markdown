#!/usr/bin/env python3
"""
Normalize Surya outputs into per-page reading order.
"""

from __future__ import annotations

import argparse
from difflib import SequenceMatcher
import json
import re
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple
import xml.etree.ElementTree as ET
import zipfile


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}
REL_NS = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}

REORDERABLE = {"sp", "pic", "graphicFrame", "grpSp", "cxnSp"}
TITLE_TYPES = {"title", "ctrTitle", "subTitle"}


@dataclass
class XmlObject:
    shape_id: str
    tag: str
    ph_type: Optional[str]
    x: float
    y: float
    w: float
    h: float
    cx: float
    cy: float
    text: str
    font_pt: Optional[float]


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def pick_doc_payload(payload: Any, doc_key: Optional[str]) -> Tuple[str, Any]:
    if isinstance(payload, list):
        return "default", payload
    if not isinstance(payload, dict):
        raise ValueError("invalid json payload type")
    if doc_key:
        if doc_key not in payload:
            raise KeyError(f"doc key not found: {doc_key}")
        return doc_key, payload[doc_key]
    keys = list(payload.keys())
    if not keys:
        raise ValueError("empty payload dictionary")
    return keys[0], payload[keys[0]]


def safe_float(v: Any, default: float = 0.0) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def safe_int(v: Any, default: int = 0) -> int:
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


def norm_bbox(raw: Any) -> Optional[List[float]]:
    if not isinstance(raw, list) or len(raw) != 4:
        return None
    x1 = safe_float(raw[0])
    y1 = safe_float(raw[1])
    x2 = safe_float(raw[2])
    y2 = safe_float(raw[3])
    return [min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2)]


def merge_bboxes(bboxes: Sequence[List[float]]) -> Optional[List[float]]:
    if not bboxes:
        return None
    xs1 = [b[0] for b in bboxes]
    ys1 = [b[1] for b in bboxes]
    xs2 = [b[2] for b in bboxes]
    ys2 = [b[3] for b in bboxes]
    return [min(xs1), min(ys1), max(xs2), max(ys2)]


def bbox_area(bbox: Optional[List[float]]) -> float:
    if bbox is None:
        return 0.0
    return max(0.0, bbox[2] - bbox[0]) * max(0.0, bbox[3] - bbox[1])


def intersection_bbox(a: Optional[List[float]], b: Optional[List[float]]) -> Optional[List[float]]:
    if a is None or b is None:
        return None
    x1 = max(a[0], b[0])
    y1 = max(a[1], b[1])
    x2 = min(a[2], b[2])
    y2 = min(a[3], b[3])
    if x2 <= x1 or y2 <= y1:
        return None
    return [x1, y1, x2, y2]


def overlap_ratio(a: Optional[List[float]], b: Optional[List[float]]) -> float:
    inter = intersection_bbox(a, b)
    if inter is None:
        return 0.0
    inter_area = bbox_area(inter)
    denom = min(bbox_area(a), bbox_area(b))
    if denom <= 0:
        return 0.0
    return inter_area / denom


def center_distance_score(a: Optional[List[float]], b: Optional[List[float]]) -> float:
    if a is None or b is None:
        return 0.0
    acx = (a[0] + a[2]) / 2.0
    acy = (a[1] + a[3]) / 2.0
    bcx = (b[0] + b[2]) / 2.0
    bcy = (b[1] + b[3]) / 2.0
    dist = ((acx - bcx) ** 2 + (acy - bcy) ** 2) ** 0.5
    return max(0.0, 1.0 - min(dist / 1.5, 1.0))


def text_similarity(a: str, b: str) -> float:
    left = normalize_text(a)
    right = normalize_text(b)
    if not left or not right:
        return 0.0
    if left in right or right in left:
        return 1.0
    return SequenceMatcher(None, left, right).ratio()


def local_name(tag: str) -> str:
    return tag.split("}", 1)[-1]


def normalize_text(s: str) -> str:
    s = re.sub(r"\s+", " ", (s or "").strip())
    return s


def first_off_and_ext(elem: ET.Element) -> Tuple[Optional[ET.Element], Optional[ET.Element]]:
    off = None
    ext = None
    for p in (
        "./p:spPr/a:xfrm",
        "./p:grpSpPr/a:xfrm",
        "./p:xfrm",
        ".//a:xfrm",
    ):
        xfrm = elem.find(p, NS)
        if xfrm is None:
            continue
        off = xfrm.find("a:off", NS)
        ext = xfrm.find("a:ext", NS)
        if off is not None or ext is not None:
            return off, ext
    return off, ext


def bbox_from_off_ext(off: Optional[ET.Element], ext: Optional[ET.Element]) -> Optional[List[float]]:
    if off is None or ext is None:
        return None
    x = safe_float(off.attrib.get("x"))
    y = safe_float(off.attrib.get("y"))
    w = safe_float(ext.attrib.get("cx"))
    h = safe_float(ext.attrib.get("cy"))
    if w <= 0 or h <= 0:
        return None
    return [x, y, x + w, y + h]


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


def _resolve_related_part(source_xml: Path, rel_target: str) -> Path:
    return (source_xml.parent / rel_target).resolve()


def _relationship_target(source_xml: Path, rel_type_suffix: str) -> Optional[Path]:
    rels_path = source_xml.parent / "_rels" / f"{source_xml.name}.rels"
    if not rels_path.exists():
        return None
    root = ET.parse(rels_path).getroot()
    for rel in root.findall("r:Relationship", REL_NS):
        rel_type = str(rel.attrib.get("Type", ""))
        if not rel_type.endswith(rel_type_suffix):
            continue
        target = rel.attrib.get("Target")
        if not target:
            continue
        return _resolve_related_part(source_xml, target)
    return None


def _ph_type_match(a: Optional[str], b: Optional[str]) -> bool:
    if a == b:
        return True
    if a in TITLE_TYPES and b in TITLE_TYPES:
        return True
    return False


def _find_placeholder_bbox_in_part(part_xml: Path, ph_type: Optional[str], ph_idx: Optional[str]) -> Optional[List[float]]:
    if not part_xml.exists():
        return None
    root = ET.parse(part_xml).getroot()
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        return None

    best_bbox: Optional[List[float]] = None
    best_score = -1
    for ch in list(sp_tree):
        tag = local_name(ch.tag)
        if tag not in REORDERABLE:
            continue
        _, ph_path = get_nvpr_paths(tag)
        ph = ch.find(ph_path, NS)
        if ph is None:
            continue
        cand_type = ph.attrib.get("type")
        cand_idx = ph.attrib.get("idx")
        off, ext = first_off_and_ext(ch)
        bbox = bbox_from_off_ext(off, ext)
        if bbox is None:
            continue

        if ph_idx is not None:
            if cand_idx != ph_idx:
                continue
            score = 4
        else:
            score = 1 if cand_idx is None else 0

        if ph_type is not None:
            if not _ph_type_match(ph_type, cand_type):
                continue
            score += 2
        else:
            score += 1 if cand_type is None else 0

        if score > best_score:
            best_score = score
            best_bbox = bbox
    return best_bbox


def find_inherited_placeholder_bbox(slide_xml: Path, ph_type: Optional[str], ph_idx: Optional[str]) -> Optional[List[float]]:
    if ph_type is None and ph_idx is None:
        return None
    layout_xml = _relationship_target(slide_xml, "/slideLayout")
    if layout_xml is not None:
        bbox = _find_placeholder_bbox_in_part(layout_xml, ph_type, ph_idx)
        if bbox is not None:
            return bbox
        master_xml = _relationship_target(layout_xml, "/slideMaster")
        if master_xml is not None:
            bbox = _find_placeholder_bbox_in_part(master_xml, ph_type, ph_idx)
            if bbox is not None:
                return bbox
    return None


def extract_shape_text(elem: ET.Element) -> str:
    parts: List[str] = []
    for t in elem.findall(".//a:t", NS):
        if t.text and t.text.strip():
            parts.append(t.text.strip())
    return normalize_text(" ".join(parts))


def extract_font_pt(elem: ET.Element) -> Optional[float]:
    sizes: List[float] = []
    for rpr in elem.findall(".//a:rPr", NS):
        sz = rpr.attrib.get("sz")
        if sz is None:
            continue
        val = safe_float(sz, -1.0)
        if val > 0:
            sizes.append(val / 100.0)
    for rpr in elem.findall(".//a:endParaRPr", NS):
        sz = rpr.attrib.get("sz")
        if sz is None:
            continue
        val = safe_float(sz, -1.0)
        if val > 0:
            sizes.append(val / 100.0)
    if not sizes:
        return None
    return max(sizes)


def parse_slide_xml_objects(slide_xml: Path) -> List[XmlObject]:
    root = ET.parse(slide_xml).getroot()
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        return []

    out: List[XmlObject] = []
    for ch in list(sp_tree):
        tag = local_name(ch.tag)
        if tag not in REORDERABLE:
            continue
        c_nv_path, ph_path = get_nvpr_paths(tag)
        c_nv_pr = ch.find(c_nv_path, NS)
        ph = ch.find(ph_path, NS)
        shape_id = c_nv_pr.attrib.get("id", "") if c_nv_pr is not None else ""
        ph_type = ph.attrib.get("type") if ph is not None else None

        off, ext = first_off_and_ext(ch)
        bbox = bbox_from_off_ext(off, ext)
        if bbox is None:
            bbox = find_inherited_placeholder_bbox(slide_xml, ph_type, ph.attrib.get("idx") if ph is not None else None)
        if bbox is None:
            x = 0.0
            y = 0.0
            w = 0.0
            h = 0.0
        else:
            x = bbox[0]
            y = bbox[1]
            w = bbox[2] - bbox[0]
            h = bbox[3] - bbox[1]
        cx = x + (w / 2.0)
        cy = y + (h / 2.0)

        out.append(
            XmlObject(
                shape_id=shape_id,
                tag=tag,
                ph_type=ph_type,
                x=x,
                y=y,
                w=w,
                h=h,
                cx=cx,
                cy=cy,
                text=extract_shape_text(ch),
                font_pt=extract_font_pt(ch),
            )
        )
    return out


def parse_slide_size_emu(ppt_root: Path) -> Tuple[float, float]:
    pres = ppt_root / "ppt" / "presentation.xml"
    if not pres.exists():
        return (1.0, 1.0)
    root = ET.parse(pres).getroot()
    sld_sz = root.find("p:sldSz", NS)
    if sld_sz is None:
        return (1.0, 1.0)
    cx = safe_float(sld_sz.attrib.get("cx"), 1.0)
    cy = safe_float(sld_sz.attrib.get("cy"), 1.0)
    return (max(cx, 1.0), max(cy, 1.0))


def xml_object_to_normalized_bbox(
    obj: XmlObject,
    slide_w: float,
    slide_h: float,
) -> Optional[List[float]]:
    if obj.w <= 0 or obj.h <= 0 or slide_w <= 0 or slide_h <= 0:
        return None
    return [
        obj.x / slide_w,
        obj.y / slide_h,
        (obj.x + obj.w) / slide_w,
        (obj.y + obj.h) / slide_h,
    ]


def xml_object_to_emu_bbox(obj: XmlObject) -> Optional[List[float]]:
    if obj.w <= 0 or obj.h <= 0:
        return None
    # XML xfrm(EMU): (x, y, cx, cy) -> bbox(EMU): [left, top, right, bottom]
    # left=x, top=y, right=x+cx, bottom=y+cy
    return [obj.x, obj.y, obj.x + obj.w, obj.y + obj.h]


def layout_bbox_px_to_emu_bbox(
    block_bbox: List[float],
    image_bbox: List[float],
    slide_w: float,
    slide_h: float,
) -> List[float]:
    # image_bbox: full page image extent in px (global frame), usually [0, 0, image_w, image_h].
    # block_bbox: one detected layout block in the same px frame.
    # We map block_bbox from image-px coordinates into slide-EMU coordinates.
    #
    # Concept:
    # - image_bbox provides the conversion frame (origin + total size).
    # - slide_w/slide_h are the same page size in EMU.
    # - So scale is "EMU per pixel":
    #   sx = slide_w / image_w_px, sy = slide_h / image_h_px.
    # - After shifting by image origin, multiplying by sx/sy gives EMU.
    # image_bbox = [left_px, top_px, right_px, bottom_px]
    # iw, ih are image width/height in px.
    iw = max(1e-6, image_bbox[2] - image_bbox[0])
    ih = max(1e-6, image_bbox[3] - image_bbox[1])
    # Scale factors from px -> EMU.
    # sx = slide_width_emu / image_width_px
    # sy = slide_height_emu / image_height_px
    sx = slide_w / iw
    sy = slide_h / ih
    return [
        # Normalize to image origin first, then convert px to EMU.
        # left_emu   = (left_px   - image_left_px) * sx
        (block_bbox[0] - image_bbox[0]) * sx,
        # top_emu    = (top_px    - image_top_px)  * sy
        (block_bbox[1] - image_bbox[1]) * sy,
        # right_emu  = (right_px  - image_left_px) * sx
        (block_bbox[2] - image_bbox[0]) * sx,
        # bottom_emu = (bottom_px - image_top_px)  * sy
        (block_bbox[3] - image_bbox[1]) * sy,
    ]


def extract_pptx_to_temp_root(pptx_path: Path) -> Tuple[Path, Path]:
    if not pptx_path.exists() or not pptx_path.is_file():
        raise FileNotFoundError(f"pptx not found: {pptx_path}")
    temp_dir = Path(tempfile.mkdtemp(prefix="pptx_xml_"))
    extracted_root = temp_dir / pptx_path.stem
    extracted_root.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(pptx_path, "r") as zf:
        zf.extractall(extracted_root)
    return temp_dir, extracted_root


def resolve_results_json(raw: Optional[str], default_dir: Path, kind: str) -> Path:
    if raw:
        p = Path(raw).resolve()
        if p.is_file():
            return p
        if p.is_dir():
            hits = sorted(p.glob("*/results.json"))
            if hits:
                return hits[0].resolve()
        raise FileNotFoundError(f"{kind} results not found from: {p}")

    hits = sorted(default_dir.glob("*/results.json"))
    if hits:
        return hits[0].resolve()
    raise FileNotFoundError(f"{kind} results not found under default dir: {default_dir}")


def match_layout_to_xml(
    block_bbox: List[float],
    image_bbox: List[float],
    xml_objects: Sequence[XmlObject],
    slide_w: float,
    slide_h: float,
) -> Optional[XmlObject]:
    if not xml_objects:
        return None
    block_bbox_emu = layout_bbox_px_to_emu_bbox(block_bbox, image_bbox, slide_w, slide_h)
    best: Optional[XmlObject] = None
    best_score = -1.0
    for obj in xml_objects:
        obj_bbox = xml_object_to_emu_bbox(obj)
        if obj_bbox is None:
            continue
        score = (0.65 * overlap_ratio(block_bbox_emu, obj_bbox)) + (0.20 * center_distance_score(block_bbox_emu, obj_bbox))
        if obj.tag in {"pic", "graphicFrame", "cxnSp"}:
            score -= 0.10
        if score > best_score:
            best_score = score
            best = obj
    return best


def build_reading_order_from_layout(
    layout_rows: Sequence[Dict[str, Any]],
    xml_objects: Sequence[XmlObject],
    slide_w: float,
    slide_h: float,
) -> List[Dict[str, Any]]:
    by_shape_id = {obj.shape_id: obj for obj in xml_objects if obj.shape_id}
    support_rows: Dict[str, List[Dict[str, Any]]] = {}
    fallback_rows: List[Dict[str, Any]] = []
    for row in layout_rows:
        shape_id = str(row.get("matched_shape_id") or "").strip()
        if shape_id:
            support_rows.setdefault(shape_id, []).append(row)
        else:
            fallback_rows.append(row)

    reading: List[Dict[str, Any]] = []
    covered_shape_ids = set()
    for shape_id, rows in support_rows.items():
        if shape_id not in by_shape_id:
            continue
        rows.sort(
            key=lambda r: (
                safe_int(r.get("position"), 10**9),
                ((r.get("bbox") or [0.0, 0.0, 0.0, 0.0])[1]),
                ((r.get("bbox") or [0.0, 0.0, 0.0, 0.0])[0]),
            )
        )
        reading.append(dict(rows[0]))
        covered_shape_ids.add(shape_id)

    used_fallback_rows = set()
    for obj in xml_objects:
        shape_id = obj.shape_id
        if not shape_id or shape_id in covered_shape_ids:
            continue
        obj_text = normalize_text(obj.text)
        if not obj_text or obj.tag in {"pic", "graphicFrame", "cxnSp"}:
            continue
        obj_bbox = xml_object_to_emu_bbox(obj)
        if obj_bbox is None:
            continue
        best_idx = None
        best_score = -1.0
        for idx, row in enumerate(fallback_rows):
            if idx in used_fallback_rows:
                continue
            row_bbox = row.get("bbox") if isinstance(row.get("bbox"), list) else None
            row_text = normalize_text(str(row.get("xml_text") or ""))
            score = (0.45 * overlap_ratio(obj_bbox, row_bbox)) + (0.15 * center_distance_score(obj_bbox, row_bbox))
            if row_text:
                score += 0.40 * text_similarity(obj_text, row_text)
            if score > best_score:
                best_score = score
                best_idx = idx
        if best_idx is None or best_score < 0.35:
            continue
        used_fallback_rows.add(best_idx)
        row = dict(fallback_rows[best_idx])
        row["matched_shape_id"] = shape_id
        row["matched_placeholder_type"] = obj.ph_type
        row["xml_text"] = obj.text
        row["position"] = 10**5 + safe_int(row.get("position"), 10**5)
        reading.append(row)
        covered_shape_ids.add(shape_id)

    reading.sort(
        key=lambda x: (
            safe_int(x.get("position"), 10**9),
            ((x.get("bbox") or [0.0, 0.0, 0.0, 0.0])[1]),
            ((x.get("bbox") or [0.0, 0.0, 0.0, 0.0])[0]),
        )
    )
    return reading


def main() -> int:
    script_dir = Path(__file__).resolve().parent

    parser = argparse.ArgumentParser(description="Normalize Surya layout outputs for reading order.")
    parser.add_argument(
        "--layout-json",
        default=None,
        help="Path to layout results.json OR layout result directory. If omitted, auto-picks under ./output/layout_result.",
    )
    parser.add_argument(
        "--pptx-path",
        default=None,
        help="Optional path to source .pptx. If set, script extracts XML internally.",
    )
    parser.add_argument(
        "--ppt-root",
        default=None,
        help="Path to extracted pptx root directory (contains ppt/slides).",
    )
    parser.add_argument(
        "--output-json",
        default=str(script_dir / "output" / "normalized" / "normalized_results.json"),
        help="Path to output normalized json",
    )
    parser.add_argument(
        "--doc-key",
        default=None,
        help="Optional document key in input JSON. If omitted, first key is used.",
    )
    args = parser.parse_args()

    default_layout_dir = script_dir / "output" / "layout_result"
    layout_path = resolve_results_json(args.layout_json, default_layout_dir, kind="layout")
    output_path = Path(args.output_json).resolve()
    pptx_path = Path(args.pptx_path).resolve() if args.pptx_path else None

    temp_ppt_dir: Optional[Path] = None
    if args.ppt_root:
        ppt_root = Path(args.ppt_root).resolve()
    elif pptx_path is not None:
        temp_ppt_dir, ppt_root = extract_pptx_to_temp_root(pptx_path)
    else:
        raise ValueError("either --ppt-root or --pptx-path is required")

    layout_payload = load_json(layout_path)
    layout_key, layout_pages_raw = pick_doc_payload(layout_payload, args.doc_key)

    if not isinstance(layout_pages_raw, list):
        raise ValueError("layout payload doc value must be a list")

    slide_w, slide_h = parse_slide_size_emu(ppt_root)

    pages: Dict[int, Dict[str, Any]] = {}

    for page in layout_pages_raw:
        if not isinstance(page, dict):
            continue
        page_no = safe_int(page.get("page"), 0)
        if page_no <= 0:
            continue
        image_bbox = norm_bbox(page.get("image_bbox")) or [0.0, 0.0, 1.0, 1.0]

        slide_xml = ppt_root / "ppt" / "slides" / f"slide{page_no}.xml"
        xml_objects = parse_slide_xml_objects(slide_xml) if slide_xml.exists() else []

        bboxes = page.get("bboxes", [])
        if not isinstance(bboxes, list):
            bboxes = []

        layout_rows: List[Dict[str, Any]] = []
        for blk in bboxes:
            if not isinstance(blk, dict):
                continue
            label = str(blk.get("label", "")).strip()
            bbox = norm_bbox(blk.get("bbox"))
            if bbox is None:
                continue
            bbox_emu = layout_bbox_px_to_emu_bbox(bbox, image_bbox, slide_w, slide_h)
            position = safe_int(blk.get("position"), 10**9)
            conf = safe_float(blk.get("confidence"), 0.0)

            xml_obj = match_layout_to_xml(bbox, image_bbox, xml_objects, slide_w, slide_h)
            xml_text = xml_obj.text if xml_obj else ""

            row = {
                "position": position,
                "label": label,
                "confidence": conf,
                "bbox": bbox_emu,
                "xml_text": xml_text,
                "matched_shape_id": xml_obj.shape_id if xml_obj else None,
                "matched_placeholder_type": xml_obj.ph_type if xml_obj else None,
            }
            layout_rows.append(row)

        reading = build_reading_order_from_layout(
            layout_rows=layout_rows,
            xml_objects=xml_objects,
            slide_w=slide_w,
            slide_h=slide_h,
        )
        if not reading:
            reading = list(layout_rows)

        reading.sort(key=lambda x: (x["position"], x["bbox"][1], x["bbox"][0]))
        pages[page_no] = {
            "page": page_no,
            "image_bbox": image_bbox,
            "reading_order": [
                {
                    "order_index": idx,
                    **r,
                }
                for idx, r in enumerate(reading)
            ],
        }

    ordered_pages = [pages[k] for k in sorted(pages.keys())]
    for p in ordered_pages:
        p["counts"] = {
            "reading_blocks": len(p["reading_order"]),
        }

    out = {
        "source": {
            "layout_json": str(layout_path),
            "layout_doc_key": layout_key,
            "pptx_path": str(pptx_path) if pptx_path and pptx_path.exists() else None,
            "ppt_root": str(ppt_root),
        },
        "rules": {
            "signals": [
                "layout_label",
                "layout_confidence",
                "layout_bbox",
                "xml_match",
                "xml_position",
            ],
        },
        "pages": ordered_pages,
        "summary": {
            "page_count": len(ordered_pages),
            "reading_blocks_total": sum(len(p["reading_order"]) for p in ordered_pages),
        },
    }

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote: {output_path}")
    print(
        f"Summary: pages={out['summary']['page_count']} "
        f"reading_blocks={out['summary']['reading_blocks_total']}"
    )

    if temp_ppt_dir is not None and temp_ppt_dir.exists():
        shutil.rmtree(temp_ppt_dir, ignore_errors=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
