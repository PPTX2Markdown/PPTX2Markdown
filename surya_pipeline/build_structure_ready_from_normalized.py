from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from layout_matcher import load_presentation_root, reorder_slide_xml, slide_xml_paths


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build structure-ready reordered slide XML files from normalized Surya matches."
    )
    parser.add_argument("--normalized-json", required=True, type=Path)
    parser.add_argument("--ppt-root", type=Path)
    parser.add_argument("--pptx-path", type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument(
        "--placeholder-inheritance",
        "--pptx-inheritance",
        dest="pptx_inheritance",
        choices=("none", "geometry", "style", "placeholder", "semantic"),
        default="style",
        help="Accepted for compatibility with the Surya pipeline driver.",
    )
    return parser.parse_args()


def _load(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _page_map(payload: dict[str, Any]) -> dict[int, dict[str, Any]]:
    pages = payload.get("pages")
    if isinstance(pages, list):
        out: dict[int, dict[str, Any]] = {}
        for page in pages:
            if isinstance(page, dict):
                try:
                    out[int(page.get("page"))] = page
                except (TypeError, ValueError):
                    pass
        return out
    rows = payload.get("reading_order")
    out = {}
    if isinstance(rows, list):
        by_page: dict[int, list[dict[str, Any]]] = {}
        for row in rows:
            if not isinstance(row, dict):
                continue
            try:
                page_num = int(row.get("page"))
            except (TypeError, ValueError):
                continue
            by_page.setdefault(page_num, []).append(row)
        for page_num, page_rows in by_page.items():
            out[page_num] = {"page": page_num, "reading_order": page_rows}
    return out


def _shape_order(rows: list[dict[str, Any]]) -> list[str]:
    order: list[str] = []
    seen: set[str] = set()
    for row in sorted(rows, key=lambda r: int(r.get("order_index") or 0)):
        shape_id = str(row.get("shape_id") or row.get("matched_shape_id") or "").strip()
        if not shape_id or shape_id in seen:
            continue
        seen.add(shape_id)
        order.append(shape_id)
    return order


def main() -> None:
    args = parse_args()
    ppt_root, _ = load_presentation_root(ppt_root=args.ppt_root, pptx_path=args.pptx_path)
    payload = _load(args.normalized_json)
    pages = _page_map(payload)
    slides = slide_xml_paths(ppt_root)
    args.output_dir.mkdir(parents=True, exist_ok=True)

    manifest: dict[str, Any] = {
        "schema_version": "structure_ready_from_surya_layout_match_v2",
        "normalized_json": str(args.normalized_json),
        "ppt_root": str(ppt_root),
        "slides": [],
    }

    for page_num, slide_xml in enumerate(slides, start=1):
        page = pages.get(page_num, {"page": page_num, "reading_order": []})
        rows = page.get("reading_order") if isinstance(page.get("reading_order"), list) else []
        output_xml = args.output_dir / f"slide{page_num}.reordered.xml"
        sidecar_json = args.output_dir / f"slide{page_num}.structure_analysis.json"
        shape_order = _shape_order(rows)
        reorder_slide_xml(slide_xml, shape_order, output_xml)
        sidecar = {
            "schema_version": "structure_analysis_surya_layout_match_v2",
            "page": page_num,
            "input_xml": str(slide_xml),
            "output_xml": str(output_xml),
            "structure_order": rows,
            "reading_order": rows,
            "layout_blocks": page.get("layout_blocks", []),
            "unmatched_layout_blocks": page.get("unmatched_layout_blocks", []),
            "decorative_objects": page.get("decorative_objects", []),
        }
        sidecar_json.write_text(json.dumps(sidecar, ensure_ascii=False, indent=2), encoding="utf-8")
        manifest["slides"].append(
            {
                "page": page_num,
                "input_xml": str(slide_xml),
                "output_xml": str(output_xml),
                "sidecar_json": str(sidecar_json),
                "shape_count": len(shape_order),
            }
        )

    manifest_path = args.output_dir / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[OK] structure-ready: {args.output_dir}")


if __name__ == "__main__":
    main()
