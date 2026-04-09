from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, List, Optional, Sequence
import xml.etree.ElementTree as ET

from .constants import NS, REORDERABLE
from .extractor import extract_slide_objects_xml
from .structure import (
    OrderContext,
    SlideObject,
    compute_heading_depths,
    build_order_context,
    bucket,
    heading_score,
    heading_threshold,
    order_objects,
    reason,
)
from .xml_primitives import local_name, natural_key


SCHEMA_VERSION = "1.0"


def confidence(objects: Sequence[SlideObject]) -> str:
    if not objects:
        return "low"
    unknown = sum(1 for obj in objects if obj.coord_source == "unknown")
    ratio = unknown / len(objects)
    if ratio >= 0.3:
        return "low"
    if ratio >= 0.1:
        return "medium"
    return "high"


def object_to_dict(
    obj: SlideObject,
    context: OrderContext,
    heading_depths: Optional[Dict[str, Optional[int]]] = None,
    strict: bool = False,
) -> Dict[str, object]:
    if heading_depths is None:
        depth = compute_heading_depths([obj], context, strict=strict).get(obj.shape_id)
    else:
        depth = heading_depths.get(obj.shape_id)

    score = heading_score(obj, strict=strict)
    is_candidate = depth is not None and score >= heading_threshold(strict=strict)
    return {
        "shape_id": obj.shape_id,
        "xml_index": obj.xml_index,
        "tag": obj.tag,
        "name": obj.name,
        "ph_type": obj.ph_type,
        "ph_idx": obj.ph_idx,
        "x": obj.x,
        "y": obj.y,
        "coord_source": obj.coord_source,
        "text": obj.text,
        "font_pt": obj.font_pt,
        "is_footer": obj.is_footer,
        "is_decorative": obj.is_decorative,
        "is_heading": obj.is_heading,
        "is_title_placeholder": obj.is_title_placeholder,
        "is_heading_candidate": is_candidate,
        "heading_score": round(score, 3),
        "heading_depth_hint": depth,
        "bbox": list(obj.bbox) if obj.bbox is not None else None,
        "bucket": bucket(obj, context),
        "reason": reason(obj, context),
    }


def reorder_tree_by_indexes(tree: ET.ElementTree, ordered_xml_indexes: Sequence[int]) -> None:
    root = tree.getroot()
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        return

    children = list(sp_tree)
    reorderables = [child for child in children if local_name(child.tag) in REORDERABLE]
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
    for child in children:
        if local_name(child.tag) in REORDERABLE:
            if not inserted:
                new_children.extend(reordered_elems)
                inserted = True
            continue
        new_children.append(child)
    sp_tree[:] = new_children


def gather_input_files(target_dir: Path, raw_inputs: Sequence[str]) -> List[Path]:
    files: List[Path] = []

    if raw_inputs:
        for item in raw_inputs:
            path = Path(item)
            candidates = [path, target_dir / item]
            if path.suffix == "":
                candidates.append(target_dir / f"{item}.xml")

            selected: Optional[Path] = None
            for candidate in candidates:
                if candidate.exists() and candidate.is_file():
                    selected = candidate
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
    for file in files:
        unique[str(file)] = file
    return list(unique.values())


def write_outputs(
    slide_xml: Path,
    output_dir: Path,
    mode: str,
    strict: bool = False,
) -> Dict[str, object]:
    objects, meta = extract_slide_objects_xml(slide_xml, strict=strict)
    meta["mode"] = mode
    meta["strict"] = strict
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
        "schema_version": SCHEMA_VERSION,
        "input_xml": str(slide_xml),
        "mode": mode,
        "strict": strict,
        "layout_xml": meta.get("layout_xml"),
        "confidence": confidence(objects),
        "counts": {
            "total": len(objects),
            "text": sum(1 for obj in objects if bool(obj.normalized)),
            "graphicFrame": sum(1 for obj in objects if obj.tag == "graphicFrame"),
            "pic": sum(1 for obj in objects if obj.tag == "pic"),
            "footer": sum(1 for obj in objects if obj.is_footer),
            "decorative": sum(1 for obj in objects if obj.is_decorative),
            "layout_coord_used": sum(1 for obj in objects if obj.coord_source == "layout"),
            "xml_tables": len(meta.get("xml_tables", [])),
            "xml_images": len(meta.get("xml_images", [])),
        },
        "xml_tables": meta.get("xml_tables", []),
        "xml_images": meta.get("xml_images", []),
        "structure_order": [object_to_dict(obj, context, ordered_heading_depths, strict=strict) for obj in ordered],
        "raw_xml_order": [
            object_to_dict(obj, context, raw_heading_depths, strict=strict)
            for obj in sorted(objects, key=lambda x: x.xml_index)
        ],
        "ordered_xml_indexes": ordered_indexes,
        "output_structure_xml": str(xml_path),
    }

    json_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

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
