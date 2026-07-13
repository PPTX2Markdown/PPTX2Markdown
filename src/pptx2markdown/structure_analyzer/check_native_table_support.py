#!/usr/bin/env python3
"""Inspect whether PPTX native structure can replace image-first table screening.

This reports, per slide:
  - native PPTX tables (<a:tbl> in graphicFrame)
  - picture shapes (<p:pic>)
  - pictures that are fully overlaid on native tables
  - pictures that remain standalone candidates for image verification
"""

from __future__ import annotations

import argparse
import json
import shutil
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from pptx2markdown.ooxml_security import resolve_package_part, safe_extract_ooxml_archive

from .constants import NS, REL_NS, REORDERABLE
from .xml_primitives import (
    extract_bbox_emu,
    get_nvpr_paths,
    local_name,
    natural_key,
)

OVERLAY_PICTURE_MIN_RATIO = 0.80


def bbox_area(bbox: Tuple[int, int, int, int]) -> int:
    return max(0, bbox[2] - bbox[0]) * max(0, bbox[3] - bbox[1])


def intersection_area(left: Tuple[int, int, int, int], right: Tuple[int, int, int, int]) -> int:
    x1 = max(left[0], right[0])
    y1 = max(left[1], right[1])
    x2 = min(left[2], right[2])
    y2 = min(left[3], right[3])
    return max(0, x2 - x1) * max(0, y2 - y1)


def bbox_center(bbox: Tuple[int, int, int, int]) -> Tuple[float, float]:
    return ((bbox[0] + bbox[2]) / 2.0, (bbox[1] + bbox[3]) / 2.0)


def bbox_contains_point(bbox: Tuple[int, int, int, int], point: Tuple[float, float]) -> bool:
    x, y = point
    return bbox[0] <= x <= bbox[2] and bbox[1] <= y <= bbox[3]


def extract_pptx_to_temp_root(pptx_path: Path) -> Tuple[Path, Path]:
    if not pptx_path.exists() or not pptx_path.is_file():
        raise FileNotFoundError(f"pptx not found: {pptx_path}")
    temp_dir = Path(tempfile.mkdtemp(prefix="pptx_native_probe_"))
    extracted_root = temp_dir / pptx_path.stem
    extracted_root.mkdir(parents=True, exist_ok=True)
    safe_extract_ooxml_archive(pptx_path, extracted_root)
    return temp_dir, extracted_root


def load_rels_map(slide_xml: Path) -> Dict[str, str]:
    rels_path = slide_xml.parent / "_rels" / f"{slide_xml.name}.rels"
    if not rels_path.exists():
        return {}
    root = ET.parse(rels_path).getroot()
    out: Dict[str, str] = {}
    for rel in root.findall("rel:Relationship", REL_NS):
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        target_mode = rel.attrib.get("TargetMode")
        if not rel_id or not target or target_mode == "External":
            continue
        out[rel_id] = target
    return out


def resolve_media_target(slide_xml: Path, rel_target: Optional[str]) -> Optional[Path]:
    if not rel_target:
        return None
    return resolve_package_part(slide_xml, rel_target)


def inspect_slide(slide_xml: Path) -> Dict[str, Any]:
    root = ET.parse(slide_xml).getroot()
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        return {
            "slide": slide_xml.stem,
            "input_xml": str(slide_xml),
            "error": "Missing p:cSld/p:spTree",
        }

    rels_map = load_rels_map(slide_xml)
    native_tables: List[Dict[str, Any]] = []
    pictures: List[Dict[str, Any]] = []
    reorderable_count = 0

    for child in list(sp_tree):
        tag = local_name(child.tag)
        if tag not in REORDERABLE:
            continue
        reorderable_count += 1

        c_nv_path, _ = get_nvpr_paths(tag)
        c_nv_pr = child.find(c_nv_path, NS)
        shape_id = c_nv_pr.attrib.get("id", "") if c_nv_pr is not None else ""
        name = c_nv_pr.attrib.get("name", "") if c_nv_pr is not None else ""
        bbox = extract_bbox_emu(child)

        if tag == "graphicFrame" and bbox is not None and child.find(".//a:tbl", NS) is not None:
            native_tables.append({"shape_id": shape_id, "name": name, "bbox": list(bbox)})
            continue

        if tag != "pic" or bbox is None:
            continue

        blip = child.find(".//a:blip", NS)
        embed = blip.attrib.get(f"{{{NS['r']}}}embed") if blip is not None else None
        rel_target = rels_map.get(embed, "") if embed else ""
        media_path = resolve_media_target(slide_xml, rel_target)
        pictures.append(
            {
                "shape_id": shape_id,
                "name": name,
                "bbox": list(bbox),
                "rel_id": embed,
                "rel_target": rel_target or None,
                "media_path": str(media_path) if media_path else None,
                "media_ext": media_path.suffix.lower() if media_path else None,
            }
        )

    overlay_pictures: List[Dict[str, Any]] = []
    standalone_pictures: List[Dict[str, Any]] = []

    for picture in pictures:
        picture_bbox = tuple(picture["bbox"])
        picture_area = bbox_area(picture_bbox)
        best_match: Optional[Dict[str, Any]] = None
        best_ratio = 0.0

        for table in native_tables:
            table_bbox = tuple(table["bbox"])
            overlap = intersection_area(table_bbox, picture_bbox)
            overlap_ratio = float(overlap) / float(picture_area) if picture_area > 0 else 0.0
            center_inside = bbox_contains_point(table_bbox, bbox_center(picture_bbox))
            if not center_inside or overlap_ratio < OVERLAY_PICTURE_MIN_RATIO:
                continue
            if overlap_ratio > best_ratio:
                best_ratio = overlap_ratio
                best_match = table

        enriched = dict(picture)
        if best_match is not None:
            enriched["overlay_table_shape_id"] = best_match["shape_id"]
            enriched["overlay_table_name"] = best_match["name"]
            enriched["overlay_ratio"] = round(best_ratio, 4)
            overlay_pictures.append(enriched)
        else:
            standalone_pictures.append(enriched)

    return {
        "slide": slide_xml.stem,
        "input_xml": str(slide_xml),
        "counts": {
            "reorderable_objects": reorderable_count,
            "native_tables": len(native_tables),
            "pictures": len(pictures),
            "overlay_pictures": len(overlay_pictures),
            "standalone_pictures": len(standalone_pictures),
        },
        "native_tables": native_tables,
        "overlay_pictures": overlay_pictures,
        "standalone_pictures": standalone_pictures,
        "native_first_viable": bool(native_tables or overlay_pictures),
    }


def inspect_ppt_root(ppt_root: Path) -> Dict[str, Any]:
    slides_dir = ppt_root / "ppt" / "slides"
    if not slides_dir.exists() or not slides_dir.is_dir():
        raise FileNotFoundError(f"slides directory not found: {slides_dir}")

    slide_xmls = sorted(slides_dir.glob("slide*.xml"), key=natural_key)
    slides = [inspect_slide(slide_xml) for slide_xml in slide_xmls]

    totals = {
        "slides": len(slides),
        "native_tables": sum(int(row.get("counts", {}).get("native_tables", 0)) for row in slides),
        "pictures": sum(int(row.get("counts", {}).get("pictures", 0)) for row in slides),
        "overlay_pictures": sum(
            int(row.get("counts", {}).get("overlay_pictures", 0)) for row in slides
        ),
        "standalone_pictures": sum(
            int(row.get("counts", {}).get("standalone_pictures", 0)) for row in slides
        ),
        "slides_with_native_tables": sum(
            1 for row in slides if int(row.get("counts", {}).get("native_tables", 0)) > 0
        ),
        "slides_with_pictures": sum(
            1 for row in slides if int(row.get("counts", {}).get("pictures", 0)) > 0
        ),
        "slides_needing_image_verifier": sum(
            1 for row in slides if int(row.get("counts", {}).get("standalone_pictures", 0)) > 0
        ),
    }
    totals["native_first_picture_reduction_ratio"] = round(
        1.0 - (float(totals["standalone_pictures"]) / float(totals["pictures"]))
        if totals["pictures"] > 0
        else 0.0,
        4,
    )

    return {
        "ppt_root": str(ppt_root),
        "summary": totals,
        "slides": slides,
        "decision": {
            "native_first_viable": totals["native_tables"] > 0 or totals["overlay_pictures"] > 0,
            "recommended_strategy": (
                "Use native PPTX tables first, skip overlay pictures, and send only "
                "standalone pictures to image verification."
            ),
        },
    }


def render_text_report(report: Dict[str, Any]) -> str:
    summary = report["summary"]
    lines = [
        "Native Table Support Probe",
        f"ppt_root: {report['ppt_root']}",
        (
            "summary: "
            f"slides={summary['slides']} "
            f"native_tables={summary['native_tables']} "
            f"pictures={summary['pictures']} "
            f"overlay_pictures={summary['overlay_pictures']} "
            f"standalone_pictures={summary['standalone_pictures']} "
            f"picture_reduction={summary['native_first_picture_reduction_ratio']:.2%}"
        ),
        (
            "decision: "
            f"native_first_viable={report['decision']['native_first_viable']} | "
            f"{report['decision']['recommended_strategy']}"
        ),
        "",
    ]

    for slide in report["slides"]:
        counts = slide.get("counts", {})
        if not any(
            int(counts.get(key, 0)) > 0
            for key in ("native_tables", "pictures", "overlay_pictures", "standalone_pictures")
        ):
            continue

        lines.append(
            f"{slide['slide']}: "
            f"native_tables={counts.get('native_tables', 0)} "
            f"pictures={counts.get('pictures', 0)} "
            f"overlay={counts.get('overlay_pictures', 0)} "
            f"standalone={counts.get('standalone_pictures', 0)}"
        )

        for picture in slide.get("overlay_pictures", []):
            lines.append(
                "  overlay-picture: "
                f"{picture.get('name') or picture.get('shape_id')} -> "
                f"{picture.get('overlay_table_name') or picture.get('overlay_table_shape_id')} "
                f"({picture.get('overlay_ratio', 0.0):.2f}) "
                f"{picture.get('rel_target') or ''}".rstrip()
            )

        for picture in slide.get("standalone_pictures", []):
            lines.append(
                "  standalone-picture: "
                f"{picture.get('name') or picture.get('shape_id')} "
                f"{picture.get('rel_target') or ''}".rstrip()
            )

    return "\n".join(lines)


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Probe whether native PPTX structure can drive table candidate screening."
    )
    parser.add_argument("--ppt-root", help="Extracted PPTX root containing ppt/slides.")
    parser.add_argument("--pptx-path", help="Source .pptx file to extract temporarily.")
    parser.add_argument("--output", help="Optional path to write JSON report.")
    parser.add_argument(
        "--format",
        choices=("text", "json"),
        default="text",
        help="Stdout format. JSON is always used for --output.",
    )
    args = parser.parse_args(argv)
    if bool(args.ppt_root) == bool(args.pptx_path):
        parser.error("exactly one of --ppt-root or --pptx-path is required")
    return args


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    temp_dir: Optional[Path] = None
    if args.ppt_root:
        ppt_root = Path(args.ppt_root).resolve()
    else:
        temp_dir, ppt_root = extract_pptx_to_temp_root(Path(args.pptx_path).resolve())

    try:
        report = inspect_ppt_root(ppt_root)
        if args.output:
            output_path = Path(args.output).resolve()
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_text(
                json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
            )
        if args.format == "json":
            print(json.dumps(report, ensure_ascii=False, indent=2))
        else:
            print(render_text_report(report))
        return 0
    finally:
        if temp_dir is not None:
            shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
