#!/usr/bin/env python3
"""
Extract XYCut structure order per slide XML and write:
1) structure-analysis JSON
2) reordered slide XML

Usage:
  python extract_structure_analysis.py [slide1.xml slide2.xml ...]

If no positional args are given, all *.xml in the canonical workspace
``target_slides`` directory are processed. Outputs are written under the same
workspace by default (override with --output-dir).
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from pptx2markdown.workspace_paths import WorkspacePaths, ensure_directory

from .pipeline import gather_input_files, write_outputs
from .xml_primitives import natural_key, register_xml_namespaces


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract XYCut structure from slide XML and write JSON + reordered XML."
    )
    parser.add_argument(
        "inputs",
        nargs="*",
        help="Input slide XML file(s). If omitted, scan <work-dir>/target_slides.",
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        help="Output directory. Default: <work-dir>/structure_analysis",
    )
    parser.add_argument(
        "--work-dir",
        default=None,
        help="Shared intermediate workspace. Default: ./.pptx2markdown",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Use stricter xml heading detection and candidacy thresholds.",
    )
    parser.add_argument(
        "--placeholder-inheritance",
        dest="pptx_inheritance",
        choices=("none", "geometry", "style"),
        default="style",
        help=(
            "Placeholder inheritance depth for markdown extraction. "
            "none uses slide XML only; geometry inherits placeholder type/bbox; "
            "style also inherits text style signals such as font size and list semantics."
        ),
    )
    parser.add_argument(
        "--inherited-shapes",
        choices=("none", "visible", "all"),
        default="visible",
        help=(
            "Whether to materialize layout/master-only shapes. "
            "visible keeps slideshow-visible text/images while filtering placeholder prompts; "
            "all keeps every shape."
        ),
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    paths = WorkspacePaths.from_base(work_dir=args.work_dir)
    target_dir = paths.target_slides
    output_dir = (
        Path(args.output_dir).expanduser().resolve()
        if args.output_dir
        else paths.structure_analysis
    )

    register_xml_namespaces()

    inputs = gather_input_files(target_dir, args.inputs)
    if not inputs:
        print("No input XML files found.")
        print(f"Checked default directory: {target_dir.resolve()}")
        return

    ensure_directory(output_dir, label="structure-analysis output directory")
    manifest = {"processed": [], "failed": []}

    for slide_xml in sorted(inputs, key=natural_key):
        try:
            row = write_outputs(
                slide_xml,
                output_dir,
                strict=args.strict,
                pptx_inheritance=args.pptx_inheritance,
                inherited_shapes=args.inherited_shapes,
            )
            manifest["processed"].append(row)
            print(f"Processed: {slide_xml.name}")
        except Exception as exc:  # noqa: BLE001
            manifest["failed"].append({"input_xml": str(slide_xml), "error": str(exc)})
            print(f"Failed: {slide_xml.name} -> {exc}")

    manifest_path = output_dir / "structure_analysis_manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Wrote manifest: {manifest_path.resolve()}")
    print(f"Output directory: {output_dir.resolve()}")
    print(f"Processed: {len(manifest['processed'])}, Failed: {len(manifest['failed'])}")


if __name__ == "__main__":
    main()
