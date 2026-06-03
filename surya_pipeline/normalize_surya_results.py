from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from layout_matcher import (
    build_match_candidate_pages,
    build_normalized_pages,
    build_surya_bbox_pages,
    build_xml_bbox_pages,
    load_json,
    load_presentation_root,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Normalize Surya layout output and match layout blocks to PPTX XML objects."
    )
    parser.add_argument("--layout-json", required=True, type=Path)
    parser.add_argument("--ppt-root", type=Path)
    parser.add_argument("--pptx-path", type=Path)
    parser.add_argument("--output-json", required=True, type=Path)
    return parser.parse_args()


def artifact_paths(output_json: Path) -> tuple[Path, Path, Path]:
    if output_json.name in {"normalized.json", "04_normalized.json"}:
        return (
            output_json.with_name("02_xml_bboxes.json"),
            output_json.with_name("01_surya_bboxes.json"),
            output_json.with_name("03_match_candidates.json"),
        )
    artifact_base = output_json.stem
    if artifact_base.endswith("_normalized"):
        artifact_base = artifact_base[: -len("_normalized")]
    return (
        output_json.with_name(f"{artifact_base}_xml_bboxes.json"),
        output_json.with_name(f"{artifact_base}_surya_bboxes.json"),
        output_json.with_name(f"{artifact_base}_match_candidates.json"),
    )


def main() -> None:
    args = parse_args()
    ppt_root, _ = load_presentation_root(ppt_root=args.ppt_root, pptx_path=args.pptx_path)
    layout_json = load_json(args.layout_json)
    xml_bbox_pages = build_xml_bbox_pages(ppt_root)
    surya_bbox_pages = build_surya_bbox_pages(layout_json, ppt_root)
    match_candidate_pages = build_match_candidate_pages(xml_bbox_pages, surya_bbox_pages)
    pages = build_normalized_pages(layout_json, ppt_root)
    xml_bbox_json, surya_bbox_json, match_candidates_json = artifact_paths(args.output_json)
    payload: dict[str, Any] = {
        "schema_version": "surya_layout_match_v2",
        "layout_json": str(args.layout_json),
        "ppt_root": str(ppt_root),
        "pptx_path": str(args.pptx_path) if args.pptx_path else None,
        "matching_inputs": {
            "surya_bboxes_json": str(surya_bbox_json),
            "xml_bboxes_json": str(xml_bbox_json),
            "match_candidates_json": str(match_candidates_json),
        },
        "pages": pages,
        "reading_order": [row for page in pages for row in page.get("reading_order", [])],
    }
    args.output_json.parent.mkdir(parents=True, exist_ok=True)
    xml_bbox_json.write_text(
        json.dumps(
            {
                "schema_version": "pptx_xml_bbox_v1",
                "ppt_root": str(ppt_root),
                "pages": xml_bbox_pages,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    surya_bbox_json.write_text(
        json.dumps(
            {
                "schema_version": "surya_bbox_v1",
                "layout_json": str(args.layout_json),
                "ppt_root": str(ppt_root),
                "pages": surya_bbox_pages,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    match_candidates_json.write_text(
        json.dumps(
            {
                "schema_version": "surya_xml_match_candidates_v1",
                "xml_bboxes_json": str(xml_bbox_json),
                "surya_bboxes_json": str(surya_bbox_json),
                "pages": match_candidate_pages,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    args.output_json.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"[OK] surya bboxes: {surya_bbox_json}")
    print(f"[OK] xml bboxes: {xml_bbox_json}")
    print(f"[OK] match candidates: {match_candidates_json}")
    print(f"[OK] normalized: {args.output_json}")


if __name__ == "__main__":
    main()
