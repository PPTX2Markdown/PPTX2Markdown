from __future__ import annotations

import copy
import json
import os
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Optional, Sequence

from .constants import NS, REL_NS, REORDERABLE
from .extractor import extract_slide_objects_xml
from .structure import (
    OrderContext,
    SlideObject,
    bucket,
    build_order_context,
    compute_heading_depths,
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
    mode: str = "xml",
) -> Dict[str, object]:
    if heading_depths is None:
        depth = compute_heading_depths([obj], context, strict=strict).get(obj.shape_id)
    else:
        depth = heading_depths.get(obj.shape_id)

    score = heading_score(obj, strict=strict)
    is_candidate = depth is not None and score >= heading_threshold(strict=strict)
    if mode == "raw":
        ordering_bucket = None
        ordering_reason = "Original XML order (diagnostic only)"
    elif mode == "xycut":
        ordering_bucket = None
        ordering_reason = (
            "XYCut bounding-box projection"
            if obj.bbox is not None
            else "Missing bounding box; appended by XML index"
        )
    else:
        ordering_bucket = bucket(obj, context)
        ordering_reason = reason(obj, context)
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
        "source_part": obj.source_part,
        "inheritance_kind": obj.inheritance_kind,
        "text": obj.text,
        "font_pt": obj.font_pt,
        "list_kind": obj.list_kind,
        "list_level": obj.list_level,
        "has_list_semantics": obj.list_kind in {"ul", "ol"},
        "is_footer": obj.is_footer,
        "is_decorative": obj.is_decorative,
        "is_heading": obj.is_heading,
        "is_title_placeholder": obj.is_title_placeholder,
        "is_heading_candidate": is_candidate,
        "heading_score": round(score, 3),
        "heading_depth_hint": depth,
        "heading_depth": depth,
        "bbox": list(obj.bbox) if obj.bbox is not None else None,
        "group_path": list(obj.group_path),
        "z_path": list(obj.z_path),
        "bucket": ordering_bucket,
        "reason": ordering_reason,
    }


def reorder_tree_by_indexes(tree: ET.ElementTree, ordered_xml_indexes: Sequence[int]) -> None:
    root = tree.getroot()
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        return

    children = list(sp_tree)
    reorderables = [child for child in children if local_name(child.tag) in REORDERABLE]

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


def _qn(prefix: str, name: str) -> str:
    return f"{{{NS[prefix]}}}{name}"


def _sub(
    parent: ET.Element,
    prefix: str,
    name: str,
    attrib: Optional[Dict[str, str]] = None,
) -> ET.Element:
    return ET.SubElement(parent, _qn(prefix, name), attrib or {})


def _rels_path_for_part(part_xml: Path) -> Path:
    return part_xml.parent / "_rels" / f"{part_xml.name}.rels"


def _read_relationships(rels_path: Path) -> Dict[str, Dict[str, str]]:
    if not rels_path.exists():
        return {}
    root = ET.parse(rels_path).getroot()
    out: Dict[str, Dict[str, str]] = {}
    for rel in root.findall("rel:Relationship", REL_NS):
        rid = rel.attrib.get("Id")
        if rid:
            out[rid] = dict(rel.attrib)
    return out


def _write_relationships(rels_path: Path, rels: Sequence[Dict[str, str]]) -> None:
    rels_path.parent.mkdir(parents=True, exist_ok=True)
    root = ET.Element(
        "Relationships",
        {"xmlns": "http://schemas.openxmlformats.org/package/2006/relationships"},
    )
    for rel in rels:
        ET.SubElement(root, "Relationship", rel)
    tree = ET.ElementTree(root)
    try:
        ET.indent(tree, space="  ")
    except AttributeError:
        pass
    tree.write(rels_path, encoding="utf-8", xml_declaration=True)


def _source_part_elements(part_xml: Optional[str], source_part: str) -> Dict[str, ET.Element]:
    if not part_xml:
        return {}
    path = Path(part_xml)
    if not path.exists():
        return {}
    root = ET.parse(path).getroot()
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        return {}
    out: Dict[str, ET.Element] = {}
    for child in list(sp_tree):
        if local_name(child.tag) not in REORDERABLE:
            continue
        c_nv_pr = child.find(".//p:cNvPr", NS)
        if c_nv_pr is None:
            continue
        sid = c_nv_pr.attrib.get("id")
        if sid:
            out[f"{source_part}:{sid}"] = child
    return out


def _set_shape_id(elem: ET.Element, obj: SlideObject) -> None:
    c_nv_pr = elem.find(".//p:cNvPr", NS)
    if c_nv_pr is None:
        return
    c_nv_pr.set("id", str(obj.xml_index))
    if obj.name:
        c_nv_pr.set("name", obj.name)


def _apply_bbox(elem: ET.Element, obj: SlideObject) -> None:
    if obj.bbox is None:
        return
    sp_pr = elem.find("./p:spPr", NS)
    if sp_pr is None:
        sp_pr = _sub(elem, "p", "spPr")
    xfrm = sp_pr.find("./a:xfrm", NS)
    if xfrm is None:
        xfrm = _sub(sp_pr, "a", "xfrm")
    off = xfrm.find("./a:off", NS)
    if off is None:
        off = _sub(xfrm, "a", "off")
    ext = xfrm.find("./a:ext", NS)
    if ext is None:
        ext = _sub(xfrm, "a", "ext")
    x1, y1, x2, y2 = obj.bbox
    off.set("x", str(x1))
    off.set("y", str(y1))
    ext.set("cx", str(max(0, x2 - x1)))
    ext.set("cy", str(max(0, y2 - y1)))


def _relative_uri(path: Path, start: Path) -> str:
    return os.path.relpath(path, start=start).replace(os.sep, "/")


def _copy_materialized_shape(
    obj: SlideObject,
    source_elements: Dict[str, ET.Element],
    inherited_rels: Dict[str, str],
    rel_counter: List[int],
) -> Optional[ET.Element]:
    source = source_elements.get(obj.shape_id)
    if source is None:
        return None

    elem = copy.deepcopy(source)
    _set_shape_id(elem, obj)
    _apply_bbox(elem, obj)

    if obj.tag == "pic":
        blip = elem.find(".//a:blip", NS)
        old_rid = blip.attrib.get(f"{{{NS['r']}}}embed") if blip is not None else None
        target = inherited_rels.get(old_rid or "")
        if blip is not None and target:
            rel_counter[0] += 1
            new_rid = f"rIdInherited{rel_counter[0]}"
            blip.set(f"{{{NS['r']}}}embed", new_rid)
            inherited_rels[new_rid] = target
    return elem


def materialize_tree_by_objects(
    tree: ET.ElementTree,
    ordered_objects: Sequence[SlideObject],
    meta: Dict[str, object],
    output_dir: Path,
    output_rels_path: Path,
    source_slide_xml: Path,
) -> None:
    root = tree.getroot()
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        return

    children = list(sp_tree)
    reorderables = [child for child in children if local_name(child.tag) in REORDERABLE]

    idx_to_elem = {i + 1: elem for i, elem in enumerate(reorderables)}
    used_indexes: set[int] = set()
    ordered_elems: List[ET.Element] = []
    source_elements = {
        **_source_part_elements(str(meta.get("layout_xml") or ""), "layout"),
        **_source_part_elements(str(meta.get("master_xml") or ""), "master"),
    }

    inherited_targets: Dict[str, str] = {}
    for source_key in ("layout_xml", "master_xml"):
        part_raw = meta.get(source_key)
        if not isinstance(part_raw, str) or not part_raw:
            continue
        part_xml = Path(part_raw)
        part_rels = _rels_path_for_part(part_xml)
        for rid, rel in _read_relationships(part_rels).items():
            if "image" not in rel.get("Type", ""):
                continue
            target = rel.get("Target")
            if not target:
                continue
            abs_target = (part_rels.parent.parent / target).resolve()
            inherited_targets[rid] = _relative_uri(abs_target, output_dir)
    rel_counter = [0]

    for obj in ordered_objects:
        if obj.inheritance_kind == "materialized":
            synthetic = _copy_materialized_shape(
                obj,
                source_elements,
                inherited_targets,
                rel_counter,
            )
            if synthetic is not None:
                ordered_elems.append(synthetic)
            continue
        elem = idx_to_elem.get(obj.xml_index)
        if elem is None:
            continue
        ordered_elems.append(elem)
        used_indexes.add(obj.xml_index)

    for i, elem in idx_to_elem.items():
        if i not in used_indexes:
            ordered_elems.append(elem)

    new_children: List[ET.Element] = []
    inserted = False
    for child in children:
        if local_name(child.tag) in REORDERABLE:
            if not inserted:
                new_children.extend(ordered_elems)
                inserted = True
            continue
        new_children.append(child)
    if not inserted:
        new_children.extend(ordered_elems)
    sp_tree[:] = new_children

    rels: List[Dict[str, str]] = []
    source_rels = _read_relationships(_rels_path_for_part(source_slide_xml))
    for rel in source_rels.values():
        current = dict(rel)
        target = current.get("Target")
        if target and current.get("TargetMode") != "External":
            abs_target = (_rels_path_for_part(source_slide_xml).parent.parent / target).resolve()
            current["Target"] = _relative_uri(abs_target, output_dir)
        rels.append(current)
    for rid, target in inherited_targets.items():
        if not rid.startswith("rIdInherited"):
            continue
        rels.append(
            {
                "Id": rid,
                "Type": (
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                ),
                "Target": target,
            }
        )
    if rels:
        _write_relationships(output_rels_path, rels)


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
            raise NotADirectoryError(
                f"target_slides path exists but is not a directory: {target_dir}"
            )
        target_dir.mkdir(parents=True, exist_ok=True)
        files = sorted(target_dir.glob("*.xml"), key=natural_key)
        files = [f.resolve() for f in files if f.is_file()]

    unique: Dict[str, Path] = {}
    for file in files:
        unique[str(file)] = file
    return list(unique.values())


def _object_counts(objects: Sequence[SlideObject], meta: Dict[str, object]) -> Dict[str, int]:
    return {
        "total": len(objects),
        "text": sum(1 for obj in objects if bool(obj.normalized)),
        "graphicFrame": sum(1 for obj in objects if obj.tag == "graphicFrame"),
        "pic": sum(1 for obj in objects if obj.tag == "pic"),
        "footer": sum(1 for obj in objects if obj.is_footer),
        "decorative": sum(1 for obj in objects if obj.is_decorative),
        "layout_coord_used": sum(1 for obj in objects if obj.coord_source == "layout"),
        "master_coord_used": sum(1 for obj in objects if obj.coord_source == "master"),
        "materialized": sum(1 for obj in objects if obj.inheritance_kind == "materialized"),
        "materialized_layout": sum(
            1
            for obj in objects
            if obj.inheritance_kind == "materialized" and obj.source_part == "layout"
        ),
        "materialized_master": sum(
            1
            for obj in objects
            if obj.inheritance_kind == "materialized" and obj.source_part == "master"
        ),
        "xml_tables": len(meta.get("xml_tables", [])),
        "xml_images": len(meta.get("xml_images", [])),
        "groups": int(meta.get("group_count", 0)),
    }


def _analysis_report(
    *,
    slide_xml: Path,
    xml_path: Path,
    objects: Sequence[SlideObject],
    ordered: Sequence[SlideObject],
    context: OrderContext,
    heading_depths: Dict[str, Optional[int]],
    raw_heading_depths: Dict[str, Optional[int]],
    ordered_indexes: Sequence[int],
    meta: Dict[str, object],
    mode: str,
    strict: bool,
) -> Dict[str, object]:
    raw_objects = sorted(objects, key=lambda obj: obj.xml_index)
    return {
        "schema_version": SCHEMA_VERSION,
        "input_xml": str(slide_xml),
        "mode": mode,
        "strict": strict,
        "pptx_inheritance": meta.get("pptx_inheritance"),
        "inherited_shapes": meta.get("inherited_shapes"),
        "layout_xml": meta.get("layout_xml"),
        "master_xml": meta.get("master_xml"),
        "confidence": confidence(objects),
        "counts": _object_counts(objects, meta),
        "flattened_groups": bool(meta.get("flattened_groups", False)),
        "xml_tables": meta.get("xml_tables", []),
        "xml_images": meta.get("xml_images", []),
        "structure_order": [
            object_to_dict(obj, context, heading_depths, strict=strict, mode=mode)
            for obj in ordered
        ],
        "raw_xml_order": [
            object_to_dict(obj, context, raw_heading_depths, strict=strict, mode="raw")
            for obj in raw_objects
        ],
        "ordered_xml_indexes": list(ordered_indexes),
        "output_structure_xml": str(xml_path),
    }


def write_outputs(
    slide_xml: Path,
    output_dir: Path,
    mode: str,
    strict: bool = False,
    pptx_inheritance: str = "style",
    inherited_shapes: str = "visible",
) -> Dict[str, object]:
    objects, meta = extract_slide_objects_xml(
        slide_xml,
        strict=strict,
        pptx_inheritance=pptx_inheritance,
        inherited_shapes=inherited_shapes,
    )
    meta["mode"] = mode
    meta["strict"] = strict
    context = build_order_context(objects)
    ordered = order_objects(objects, mode=mode)

    ordered_heading_depths = compute_heading_depths(ordered, context, strict=strict)
    raw_objects = sorted(objects, key=lambda obj: obj.xml_index)
    raw_heading_depths = compute_heading_depths(raw_objects, context, strict=strict)
    ordered_indexes = [obj.xml_index for obj in ordered]

    tree = ET.parse(slide_xml)
    if any(obj.inheritance_kind == "materialized" for obj in ordered):
        materialize_tree_by_objects(
            tree,
            ordered,
            meta,
            output_dir,
            output_dir / "_rels" / f"{slide_xml.stem}.reordered.xml.rels",
            slide_xml,
        )
    elif not bool(meta.get("flattened_groups", False)):
        reorder_tree_by_indexes(tree, ordered_indexes)

    stem = slide_xml.stem
    json_path = output_dir / f"{stem}.structure_analysis.json"
    xml_path = output_dir / f"{stem}.reordered.xml"

    report = _analysis_report(
        slide_xml=slide_xml,
        xml_path=xml_path,
        objects=objects,
        ordered=ordered,
        context=context,
        heading_depths=ordered_heading_depths,
        raw_heading_depths=raw_heading_depths,
        ordered_indexes=ordered_indexes,
        meta=meta,
        mode=mode,
        strict=strict,
    )

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
