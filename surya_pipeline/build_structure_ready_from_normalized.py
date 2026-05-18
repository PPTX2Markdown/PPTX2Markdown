#!/usr/bin/env python3
"""
Build reordered slide XML from normalized Surya output.
Reading-order matching is the only ranking objective.
"""

from __future__ import annotations

import argparse
from difflib import SequenceMatcher
import json
import re
import shutil
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple
import xml.etree.ElementTree as ET
import zipfile
import sys

repo_root = Path(__file__).resolve().parent.parent
if str(repo_root) not in sys.path:
    sys.path.insert(0, str(repo_root))

from pptx_inheritance import resolve_effective_slide


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
REORDERABLE = {"sp", "pic", "graphicFrame", "grpSp", "cxnSp"}


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def local_name(tag: str) -> str:
    return tag.split("}", 1)[-1]


def normalize_text(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def parse_int(v: Any, default: int = 10**18) -> int:
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


def parse_float(v: Any, default: float = 0.0) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def canonical_text(s: str) -> str:
    text = normalize_text(s).lower()
    text = re.sub(r"^\d+(?:[.)]|\.\d+)*\s*", "", text)
    text = re.sub(r"[^0-9a-zA-Z가-힣]+", "", text)
    return text


def text_similarity(a: str, b: str) -> float:
    left = canonical_text(a)
    right = canonical_text(b)
    if not left or not right:
        return 0.0
    if left == right:
        return 1.0
    if left in right or right in left:
        shorter = min(len(left), len(right))
        longer = max(len(left), len(right))
        if longer > 0:
            return max(0.85, shorter / longer)
    return SequenceMatcher(None, left, right).ratio()


def extract_pptx_to_temp_root(pptx_path: Path) -> Tuple[Path, Path]:
    if not pptx_path.exists() or not pptx_path.is_file():
        raise FileNotFoundError(f"pptx not found: {pptx_path}")
    temp_dir = Path(tempfile.mkdtemp(prefix="pptx_xml_"))
    extracted_root = temp_dir / pptx_path.stem
    extracted_root.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(pptx_path, "r") as zf:
        zf.extractall(extracted_root)
    return temp_dir, extracted_root


def parse_slide_xml_objects(slide_xml: Path, pptx_inheritance: str = "style") -> List[Dict[str, Any]]:
    effective_slide = resolve_effective_slide(
        slide_xml,
        pptx_inheritance=pptx_inheritance,
        inherited_shapes="none",
    )
    out: List[Dict[str, Any]] = []
    for shape in effective_slide.shapes:
        out.append(
            {
                "shape_id": shape.shape_id,
                "xml_index": shape.xml_index,
                "tag": shape.tag,
                "name": shape.name,
                "ph_type": shape.ph_type,
                "ph_idx": shape.ph_idx,
                "x": shape.x,
                "y": shape.y,
                "bbox": list(shape.bbox) if shape.bbox is not None else None,
                "text": shape.normalized,
                "font_pt": shape.font_pt,
                "list_kind": shape.list_kind,
                "list_level": shape.list_level,
                "has_list_semantics": shape.has_list_semantics,
                "is_footer": shape.is_footer,
                "is_decorative": shape.is_decorative,
                "is_title_placeholder": shape.is_title_placeholder,
            }
        )
    return out


def load_normalized_pages(normalized_json: Path) -> Dict[int, Dict[str, Any]]:
    payload = load_json(normalized_json)
    pages = payload.get("pages", [])
    out: Dict[int, Dict[str, Any]] = {}
    if not isinstance(pages, list):
        return out
    for row in pages:
        if not isinstance(row, dict):
            continue
        page_no = parse_int(row.get("page"), 0)
        if page_no > 0:
            out[page_no] = row
    return out


def choose_unique_shape_matches(reading_order: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    seen: set[str] = set()
    out: List[Dict[str, Any]] = []
    for row in reading_order:
        if not isinstance(row, dict):
            continue
        shape_id = str(row.get("matched_shape_id") or "").strip()
        if not shape_id or shape_id in seen:
            continue
        seen.add(shape_id)
        out.append(row)
    return out


def find_alias_anchor(
    obj: Dict[str, Any],
    structure_rows: Sequence[Dict[str, Any]],
) -> Optional[Dict[str, Any]]:
    obj_text = normalize_text(str(obj.get("text", "")))
    if not obj_text:
        return None

    best: Optional[Dict[str, Any]] = None
    best_score = 0.0
    obj_is_title = bool(obj.get("is_title_placeholder"))
    obj_text_len = len(canonical_text(obj_text))
    for row in structure_rows:
        if row.get("coord_source") == "xml-unmatched":
            continue
        row_text = normalize_text(str(row.get("text", "")))
        if not row_text:
            continue
        score = text_similarity(obj_text, row_text)
        if obj_text_len <= 8 and len(canonical_text(row_text)) <= 8:
            score += 0.10
        if parse_int(obj.get("y"), 10**18) != 10**18 and parse_int(row.get("y"), 10**18) != 10**18:
            dy = abs(parse_int(obj.get("y"), 10**18) - parse_int(row.get("y"), 10**18))
            if dy <= 300000:
                score += 0.05
        if score > best_score:
            best_score = score
            best = row

    threshold = 0.92 if obj_text_len <= 6 else 0.82
    if obj_is_title:
        threshold = 0.65
    if best_score < threshold:
        return None
    return best


def bucket_for_object(obj: Dict[str, Any]) -> int:
    if obj.get("is_footer"):
        return 4
    if obj.get("is_decorative"):
        return 5
    return 2


def build_ordered_xml_indexes(
    xml_objects: Sequence[Dict[str, Any]],
    reading_order: Sequence[Dict[str, Any]],
) -> Tuple[List[int], List[Dict[str, Any]], Dict[str, int]]:
    by_shape_id = {str(obj.get("shape_id", "")): obj for obj in xml_objects}
    matched_ids: List[str] = []
    matched_rows: List[Dict[str, Any]] = []
    alias_rows_by_anchor: Dict[str, List[Dict[str, Any]]] = {}
    unmatched_rows: List[Dict[str, Any]] = []

    unique_matches = choose_unique_shape_matches(reading_order)
    for rank, row in enumerate(unique_matches):
        shape_id = str(row.get("matched_shape_id") or "").strip()
        obj = by_shape_id.get(shape_id)
        if obj is None:
            continue
        matched_ids.append(shape_id)
        matched_rows.append(
            {
                "shape_id": obj["shape_id"],
                "xml_index": obj["xml_index"],
                "surya_rank": parse_int(row.get("position"), rank),
                "tag": obj["tag"],
                "name": obj["name"],
                "ph_type": obj["ph_type"],
                "ph_idx": obj["ph_idx"],
                "x": obj["x"],
                "y": obj["y"],
                "coord_source": "surya-normalized",
                "text": normalize_text(str(row.get("xml_text") or row.get("ocr_text") or obj.get("text", ""))),
                "bbox": row.get("bbox") if isinstance(row.get("bbox"), list) else obj.get("bbox"),
                "label": row.get("label"),
                "font_pt": obj.get("font_pt"),
                "list_kind": obj.get("list_kind"),
                "list_level": obj.get("list_level"),
                "has_list_semantics": bool(obj.get("has_list_semantics")),
                "is_footer": bool(obj.get("is_footer")),
                "is_decorative": bool(obj.get("is_decorative")),
                "is_title_placeholder": bool(obj.get("is_title_placeholder")),
                "reason": "Matched by normalized reading_order",
            }
        )

    unmatched_count = 0
    alias_count = 0
    for obj in xml_objects:
        shape_id = str(obj.get("shape_id", ""))
        if shape_id in matched_ids:
            continue
        alias_anchor = None
        if (
            not bool(obj.get("is_footer"))
            and not bool(obj.get("is_decorative"))
            and normalize_text(str(obj.get("text", "")))
        ):
            alias_anchor = find_alias_anchor(obj, matched_rows)
        if alias_anchor is not None:
            matched_ids.append(shape_id)
            alias_count += 1
            alias_rows_by_anchor.setdefault(str(alias_anchor.get("shape_id", "")), []).append(
                {
                    "shape_id": obj["shape_id"],
                    "xml_index": obj["xml_index"],
                    "surya_rank": alias_anchor.get("surya_rank"),
                    "tag": obj["tag"],
                    "name": obj["name"],
                    "ph_type": obj["ph_type"],
                    "ph_idx": obj["ph_idx"],
                    "x": obj["x"],
                    "y": obj["y"],
                    "coord_source": "surya-alias",
                    "text": normalize_text(str(obj.get("text", ""))),
                    "bbox": obj.get("bbox"),
                    "label": alias_anchor.get("label"),
                    "font_pt": obj.get("font_pt"),
                    "list_kind": obj.get("list_kind"),
                    "list_level": obj.get("list_level"),
                    "has_list_semantics": bool(obj.get("has_list_semantics")),
                    "is_footer": bool(obj.get("is_footer")),
                    "is_decorative": bool(obj.get("is_decorative")),
                    "is_title_placeholder": bool(obj.get("is_title_placeholder")),
                    "reason": "Alias match to reading_order text block",
                }
            )
            continue
        unmatched_count += 1
        unmatched_rows.append(
            {
                "shape_id": obj["shape_id"],
                "xml_index": obj["xml_index"],
                "surya_rank": None,
                "tag": obj["tag"],
                "name": obj["name"],
                "ph_type": obj["ph_type"],
                "ph_idx": obj["ph_idx"],
                "x": obj["x"],
                "y": obj["y"],
                "coord_source": "xml-unmatched",
                "text": normalize_text(str(obj.get("text", ""))),
                "bbox": obj.get("bbox"),
                "label": None,
                "font_pt": obj.get("font_pt"),
                "list_kind": obj.get("list_kind"),
                "list_level": obj.get("list_level"),
                "has_list_semantics": bool(obj.get("has_list_semantics")),
                "is_footer": bool(obj.get("is_footer")),
                "is_decorative": bool(obj.get("is_decorative")),
                "is_title_placeholder": bool(obj.get("is_title_placeholder")),
                "reason": "Unmatched XML object appended after reading_order matches",
            }
        )

    structure_rows: List[Dict[str, Any]] = []
    for row in matched_rows:
        structure_rows.append(row)
    alias_rows: List[Dict[str, Any]] = []
    for row in matched_rows:
        aliases = alias_rows_by_anchor.get(str(row.get("shape_id", "")), [])
        aliases.sort(
            key=lambda item: (
                parse_int(item.get("surya_rank"), 10**18),
                parse_int(item.get("y"), 10**18),
                parse_int(item.get("x"), 10**18),
                parse_int(item.get("xml_index"), 10**18),
            )
        )
        alias_rows.extend(aliases)
    structure_rows.extend(alias_rows)
    unmatched_rows.sort(key=lambda row: parse_int(row.get("xml_index"), 10**18))
    structure_rows.extend(unmatched_rows)

    for row in structure_rows:
        row["bucket"] = bucket_for_object(row)

    ordered_xml_indexes = [int(row["xml_index"]) for row in structure_rows]
    stats = {
        "matched_shapes": len(matched_ids),
        "unmatched_shapes": unmatched_count,
        "alias_matches": alias_count,
        "duplicate_shape_matches": max(0, len([r for r in reading_order if r.get("matched_shape_id")]) - len(matched_ids)),
    }
    return ordered_xml_indexes, structure_rows, stats


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


def write_xml_pretty(tree: ET.ElementTree, output_xml: Path) -> None:
    # Keep reordered XML human-readable for debugging and diff reviews.
    if hasattr(ET, "indent"):
        ET.indent(tree, space="  ")
    tree.write(output_xml, encoding="utf-8", xml_declaration=True)


def write_structure_ready_outputs(
    slide_xml: Path,
    page_payload: Dict[str, Any],
    output_dir: Path,
    pptx_inheritance: str = "style",
) -> Dict[str, Any]:
    output_dir.mkdir(parents=True, exist_ok=True)
    xml_objects = parse_slide_xml_objects(slide_xml, pptx_inheritance=pptx_inheritance)
    reading_order = page_payload.get("reading_order", [])
    if not isinstance(reading_order, list):
        reading_order = []
    ordered_xml_indexes, structure_rows, stats = build_ordered_xml_indexes(xml_objects, reading_order)

    tree = ET.parse(slide_xml)
    reorder_tree_by_indexes(tree, ordered_xml_indexes)

    stem = slide_xml.stem
    output_xml = output_dir / f"{stem}.reordered.xml"
    output_json = output_dir / f"{stem}.structure_analysis.json"
    write_xml_pretty(tree, output_xml)
    output_json.write_text(
        json.dumps(
            {
                "schema_version": "surya-structure-ready-1.0",
                "page": page_payload.get("page"),
                "input_xml": str(slide_xml),
                "output_structure_xml": str(output_xml),
                "pptx_inheritance": pptx_inheritance,
                "structure_order": structure_rows,
                "ordered_xml_indexes": ordered_xml_indexes,
                "counts": {
                    "total": len(structure_rows),
                    "matched_shapes": stats["matched_shapes"],
                    "unmatched_shapes": stats["unmatched_shapes"],
                    "alias_matches": stats["alias_matches"],
                    "list_properties": sum(1 for row in structure_rows if row.get("has_list_semantics")),
                },
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return {
        "page": page_payload.get("page"),
        "input_xml": str(slide_xml),
        "output_xml": str(output_xml),
        "output_json": str(output_json),
        "matched_shapes": stats["matched_shapes"],
        "unmatched_shapes": stats["unmatched_shapes"],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Build reordered XML from normalized Surya output.")
    parser.add_argument("--normalized-json", required=True, help="Path to normalized Surya output JSON.")
    parser.add_argument("--ppt-root", default=None, help="Extracted PPT root directory containing ppt/slides.")
    parser.add_argument("--pptx-path", default=None, help="Optional source PPTX path. Used if --ppt-root is absent.")
    parser.add_argument("--output-dir", required=True, help="Output directory for reordered XML files.")
    parser.add_argument(
        "--placeholder-inheritance",
        "--pptx-inheritance",
        dest="pptx_inheritance",
        choices=("none", "geometry", "style", "placeholder", "semantic"),
        default="style",
        help="Placeholder inheritance depth used for XML object extraction.",
    )
    args = parser.parse_args()

    normalized_json = Path(args.normalized_json).resolve()
    output_dir = Path(args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / "structure_analysis_manifest.json").unlink(missing_ok=True)

    temp_ppt_dir: Optional[Path] = None
    if args.ppt_root:
        ppt_root = Path(args.ppt_root).resolve()
    elif args.pptx_path:
        temp_ppt_dir, ppt_root = extract_pptx_to_temp_root(Path(args.pptx_path).resolve())
    else:
        raise ValueError("either --ppt-root or --pptx-path is required")

    pages = load_normalized_pages(normalized_json)
    processed: List[Dict[str, Any]] = []
    failed: List[Dict[str, Any]] = []

    ET.register_namespace("a", NS["a"])
    ET.register_namespace("p", NS["p"])
    ET.register_namespace("r", NS["r"])

    for page_no in sorted(pages.keys()):
        slide_xml = ppt_root / "ppt" / "slides" / f"slide{page_no}.xml"
        if not slide_xml.exists():
            failed.append({"page": page_no, "error": f"slide xml not found: {slide_xml}"})
            continue
        try:
            row = write_structure_ready_outputs(
                slide_xml,
                pages[page_no],
                output_dir,
                pptx_inheritance=args.pptx_inheritance,
            )
            processed.append(row)
            print(f"[surya-struct] Processed: slide{page_no}.xml")
        except Exception as exc:  # noqa: BLE001
            failed.append({"page": page_no, "input_xml": str(slide_xml), "error": str(exc)})
            print(f"[surya-struct] Failed: slide{page_no}.xml -> {exc}")
    print(f"[surya-struct] Completed: processed={len(processed)} failed={len(failed)}")

    if temp_ppt_dir is not None and temp_ppt_dir.exists():
        shutil.rmtree(temp_ppt_dir, ignore_errors=True)
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
