#!/usr/bin/env python3
"""Run table pipeline end-to-end from PPTX inputs.

Usage:
    ./run.py [input.pptx ...]

Behavior:
    - If inputs are omitted, scans <work-dir>/target_pptx/*.pptx.
    - Stores all generated files under <work-dir>/table_pipeline.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from pptx2markdown.table_pipeline import parse as table_parse
from pptx2markdown.table_pipeline import render as table_render
from pptx2markdown.workspace_paths import WorkspacePaths, ensure_directory

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _natural_key(name: str) -> tuple[object, ...]:
    parts = re.split(r"(\d+)", name)
    out: list[object] = []
    for part in parts:
        if part.isdigit():
            out.append(int(part))
        else:
            out.append(part.lower())
    return tuple(out)


def _safe_stem(path: Path) -> str:
    return re.sub(r"[^0-9A-Za-z._-]+", "_", path.stem).strip("._") or "package"


def _collect_default_pptx_inputs(base_dir: Path) -> list[Path]:
    target_dir = WorkspacePaths.from_base(work_dir=base_dir).target_pptx
    if not target_dir.exists():
        return []
    if not target_dir.is_dir():
        raise NotADirectoryError(f"target_pptx path is not a directory: {target_dir}")
    return sorted(
        [
            path.resolve()
            for path in target_dir.glob("*.pptx")
            if path.is_file() and not path.name.startswith("~$")
        ],
        key=lambda p: _natural_key(p.name),
    )


def _resolve_pptx_input(raw: str, base_dir: Path) -> Path | None:
    raw_path = Path(raw)
    candidates: list[Path] = []

    # 1) explicit/relative path as-is (from current working directory)
    candidates.append(raw_path.expanduser())

    # 2) name-based lookup under the shared workspace target_pptx directory
    target_dir = WorkspacePaths.from_base(work_dir=base_dir).target_pptx
    candidates.append(target_dir / raw)
    if raw_path.suffix.lower() != ".pptx":
        candidates.append(target_dir / f"{raw}.pptx")

    for cand in candidates:
        resolved = cand.resolve()
        if resolved.exists() and resolved.is_file() and resolved.suffix.lower() == ".pptx":
            if resolved.name.startswith("~$"):
                continue
            return resolved
    return None


def _safe_extract_pptx(pptx_path: Path, dest_dir: Path) -> None:
    with zipfile.ZipFile(pptx_path) as zf:
        for member in zf.infolist():
            member_path = Path(member.filename)
            if member_path.is_absolute() or ".." in member_path.parts:
                raise ValueError(f"unsafe archive entry: {member.filename}")
        zf.extractall(dest_dir)


def _stage_pptx_packages(base_dir: Path, pptx_paths: list[Path]) -> list[tuple[Path, Path]]:
    staged_root = WorkspacePaths.from_base(work_dir=base_dir).target_slides
    ensure_directory(staged_root, label="target_slides directory")

    staged: list[tuple[Path, Path]] = []
    for pptx_path in pptx_paths:
        pkg = _safe_stem(pptx_path)
        pkg_dir = staged_root / pkg
        if pkg_dir.exists():
            shutil.rmtree(pkg_dir)
        pkg_dir.mkdir(parents=True, exist_ok=True)
        _safe_extract_pptx(pptx_path, pkg_dir)
        slides_dir = pkg_dir / "ppt" / "slides"
        if not slides_dir.exists() or not slides_dir.is_dir():
            raise ValueError(f"invalid pptx package after extraction: {pptx_path}")
        staged.append((pptx_path, pkg_dir))
    return staged


def _collect_table_elements(slide_xml_bytes: bytes) -> list[ET.Element]:
    root = ET.fromstring(slide_xml_bytes)
    tables: list[ET.Element] = []
    if root.tag == f"{{{A_NS}}}tbl":
        tables.append(root)
    tables.extend(root.findall(f".//{{{A_NS}}}tbl"))
    return tables


def _pretty_xml_bytes(elem: ET.Element) -> bytes:
    from xml.dom import minidom

    raw = ET.tostring(elem, encoding="utf-8")
    pretty = minidom.parseString(raw).toprettyxml(indent="  ", encoding="utf-8")
    lines = [line for line in pretty.splitlines() if line.strip()]
    return b"\n".join(lines) + b"\n"


def run_pipeline(pptx_paths: list[Path], work_dir: Path | None = None) -> int:
    if not pptx_paths:
        return 0

    paths = WorkspacePaths.from_base(work_dir=work_dir)
    extract_dir = ensure_directory(
        paths.table_extract_results, label="table extract-results directory"
    )
    parsing_dir = ensure_directory(
        paths.table_parsing_results, label="table parsing-results directory"
    )
    tables_dir = ensure_directory(
        paths.table_markdown_results, label="table markdown-results directory"
    )

    manifest: dict[str, object] = {
        "source_mode": "explicit_inputs" if pptx_paths else "target_pptx_default",
        "source_count": len(pptx_paths),
        "packages": [],
        "summary": {
            "slides_total": 0,
            "tables_extracted": 0,
            "tables_parsed": 0,
            "tables_rendered": 0,
            "errors": 0,
        },
    }

    summary = manifest["summary"]
    assert isinstance(summary, dict)
    packages = manifest["packages"]
    assert isinstance(packages, list)

    staged_packages = _stage_pptx_packages(paths.work_dir, pptx_paths)

    for pptx_path, pkg_dir in staged_packages:
        pkg = pkg_dir.name
        pkg_row: dict[str, object] = {
            "source": str(pptx_path),
            "package": pkg,
            "staged_package_dir": str(pkg_dir),
            "slides": [],
        }
        pkg_slides = pkg_row["slides"]
        assert isinstance(pkg_slides, list)

        try:
            slide_paths = sorted(
                [p for p in (pkg_dir / "ppt" / "slides").glob("slide*.xml") if p.is_file()],
                key=lambda p: _natural_key(p.name),
            )
            for slide_path in slide_paths:
                summary["slides_total"] = int(summary.get("slides_total", 0)) + 1
                slide_stem = slide_path.stem
                slide_xml = slide_path.read_bytes()
                tables = _collect_table_elements(slide_xml)

                slide_row: dict[str, object] = {
                    "slide": str(slide_path.relative_to(pkg_dir)),
                    "table_count": len(tables),
                    "extract_files": [],
                    "parsing_files": [],
                    "markdown_files": [],
                }

                for idx, table_elem in enumerate(tables, start=1):
                    base_name = f"{pkg}__{slide_stem}_{idx:04d}"
                    extract_path = extract_dir / f"{base_name}.xml"
                    json_path = parsing_dir / f"{base_name}_grid.json"
                    md_path = tables_dir / f"{base_name}_grid.md"

                    extract_path.write_bytes(_pretty_xml_bytes(table_elem))
                    summary["tables_extracted"] = int(summary.get("tables_extracted", 0)) + 1

                    try:
                        parsed = table_parse.parse_table_element(
                            table_elem,
                            source=f"{pptx_path}#{slide_path.relative_to(pkg_dir)}:{idx}",
                        )
                        json_path.write_text(
                            json.dumps(parsed, ensure_ascii=False, indent=2),
                            encoding="utf-8",
                        )
                        summary["tables_parsed"] = int(summary.get("tables_parsed", 0)) + 1

                        md = table_render.render_parsed_table_to_markdown(
                            parsed_table=parsed,
                            header_rows=1,
                            fill_merged=table_render.FILL_HEADER,
                        )
                        md_path.write_text(md, encoding="utf-8")
                        summary["tables_rendered"] = int(summary.get("tables_rendered", 0)) + 1
                    except Exception as exc:
                        summary["errors"] = int(summary.get("errors", 0)) + 1
                        slide_row.setdefault("errors", []).append(
                            {
                                "table_index": idx,
                                "error": str(exc),
                            }
                        )
                        continue

                    extract_files = slide_row["extract_files"]
                    parsing_files = slide_row["parsing_files"]
                    markdown_files = slide_row["markdown_files"]
                    assert isinstance(extract_files, list)
                    assert isinstance(parsing_files, list)
                    assert isinstance(markdown_files, list)
                    extract_files.append(extract_path.name)
                    parsing_files.append(json_path.name)
                    markdown_files.append(md_path.name)

                pkg_slides.append(slide_row)
        except Exception as exc:
            summary["errors"] = int(summary.get("errors", 0)) + 1
            pkg_row["error"] = str(exc)

        packages.append(pkg_row)

    manifest_path = parsing_dir / "manifest.from_pptx.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"[OK] processed pptx files: {len(pptx_paths)}")
    print(f"[OK] extract results: {extract_dir}")
    print(f"[OK] parsing results: {parsing_dir}")
    print(f"[OK] markdown tables: {tables_dir}")
    print(
        "[OK] summary: "
        f"slides={summary['slides_total']} "
        f"extracted={summary['tables_extracted']} "
        f"parsed={summary['tables_parsed']} "
        f"rendered={summary['tables_rendered']} "
        f"errors={summary['errors']}"
    )
    return 1 if int(summary.get("errors", 0)) > 0 else 0


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Run table pipeline from the shared workspace target_pptx directory"
    )
    parser.add_argument(
        "inputs",
        nargs="*",
        help="Optional .pptx file paths. If omitted, scans <work-dir>/target_pptx",
    )
    parser.add_argument(
        "--work-dir",
        default=None,
        help="Shared intermediate workspace. Default: ./.pptx2markdown",
    )
    args = parser.parse_args(argv[1:])

    paths = WorkspacePaths.from_base(work_dir=args.work_dir)

    if args.inputs:
        pptx_paths: list[Path] = []
        base_dir = paths.work_dir
        for raw in args.inputs:
            p = _resolve_pptx_input(raw, base_dir=base_dir)
            if p is None:
                print(f"[ERROR] invalid pptx input: {raw}", file=sys.stderr)
                return 1
            pptx_paths.append(p)
    else:
        pptx_paths = _collect_default_pptx_inputs(paths.work_dir)
        if not pptx_paths:
            print(f"[ERROR] no .pptx files found in {paths.target_pptx}", file=sys.stderr)
            return 1

    return run_pipeline(pptx_paths, work_dir=paths.work_dir)


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
