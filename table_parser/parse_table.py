#!/usr/bin/env python3
"""Parse OpenXML table XML (<a:tbl>) into a 2D grid JSON.

Usage:
    ./parse_table.py [input_table.xml ...]

Behavior:
    - Creates ./parsing_results if missing.
    - Writes parsed JSON to ./parsing_results/<input_stem>_grid.json.
    - If input is omitted, scans ../table_extractor/extract_results for all *.xml files.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
import re
import sys
import threading
import time
from typing import Callable, TypeVar
import xml.etree.ElementTree as ET

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"a": A_NS}


def _usage() -> str:
    return "Usage: ./parse_table.py [input_table.xml ...]"


def _int_attr(elem: ET.Element, name: str, default: int = 0) -> int:
    val = elem.get(name)
    if val is None:
        return default
    try:
        return int(val)
    except ValueError:
        return default


def _cell_text(tc: ET.Element) -> str:
    paragraphs: list[str] = []
    for p in tc.findall(".//a:txBody/a:p", NS):
        runs = p.findall(".//a:t", NS)
        if not runs:
            continue
        text = "".join(t.text or "" for t in runs)
        text = re.sub(r"\s+", " ", text).strip()
        if text:
            paragraphs.append(text)
    return "\n".join(paragraphs)


def _find_table(root: ET.Element) -> ET.Element | None:
    if root.tag == f"{{{A_NS}}}tbl":
        return root
    return root.find(".//a:tbl", NS)


def parse_table_xml(input_path: Path) -> dict[str, object]:
    tree = ET.parse(input_path)
    root = tree.getroot()
    table = _find_table(root)
    if table is None:
        raise ValueError("No <a:tbl> found in input XML.")

    col_elems = table.findall("./a:tblGrid/a:gridCol", NS)
    col_widths = [_int_attr(c, "w", 0) for c in col_elems]
    n_cols = len(col_widths)

    tr_elems = table.findall("./a:tr", NS)
    n_rows = len(tr_elems)

    # Active vertical merges by column.
    active_v: dict[int, dict[str, object]] = {}
    rows_out: list[list[dict[str, object]]] = []
    origin_cells: list[dict[str, object]] = []

    for r_idx, tr in enumerate(tr_elems):
        tc_elems = tr.findall("./a:tc", NS)
        row_cells: list[dict[str, object]] = []
        current_origin_col: int | None = None
        touched_v_cols: set[int] = set()

        for c_idx, tc in enumerate(tc_elems):
            grid_span = _int_attr(tc, "gridSpan", 1)
            row_span = _int_attr(tc, "rowSpan", 1)
            is_h_merge = tc.get("hMerge") == "1"
            is_v_merge = tc.get("vMerge") == "1"
            text = _cell_text(tc)

            if is_h_merge:
                origin_col = current_origin_col if current_origin_col is not None else c_idx - 1
                row_cells.append(
                    {
                        "type": "hMerge",
                        "origin": [r_idx, max(origin_col, 0)],
                        "text": text,
                    }
                )
                current_origin_col = max(origin_col, 0)
                continue

            if is_v_merge:
                if c_idx in active_v:
                    origin = active_v[c_idx]["origin"]
                    remaining = int(active_v[c_idx]["remaining"]) - 1
                    touched_v_cols.add(c_idx)
                    if remaining <= 0:
                        del active_v[c_idx]
                    else:
                        active_v[c_idx]["remaining"] = remaining
                else:
                    origin = [max(r_idx - 1, 0), c_idx]
                row_cells.append(
                    {
                        "type": "vMerge",
                        "origin": origin,
                        "text": text,
                    }
                )
                current_origin_col = None
                continue

            # Origin cell.
            cell = {
                "type": "origin",
                "origin": [r_idx, c_idx],
                "text": text,
                "rowspan": row_span,
                "colspan": grid_span,
            }
            row_cells.append(cell)
            origin_cells.append(
                {
                    "row": r_idx,
                    "col": c_idx,
                    "text": text,
                    "rowspan": row_span,
                    "colspan": grid_span,
                }
            )
            current_origin_col = c_idx

            if row_span > 1:
                for covered_col in range(c_idx, c_idx + grid_span):
                    active_v[covered_col] = {
                        "origin": [r_idx, c_idx],
                        "remaining": row_span - 1,
                    }

        # If XML omitted vMerge placeholders, backfill from active spans.
        if n_cols and len(row_cells) < n_cols:
            for c_idx in range(len(row_cells), n_cols):
                if c_idx in active_v and c_idx not in touched_v_cols:
                    origin = active_v[c_idx]["origin"]
                    remaining = int(active_v[c_idx]["remaining"]) - 1
                    if remaining <= 0:
                        del active_v[c_idx]
                    else:
                        active_v[c_idx]["remaining"] = remaining
                    row_cells.append({"type": "vMerge", "origin": origin, "text": ""})
                else:
                    row_cells.append({"type": "empty", "text": ""})

        rows_out.append(row_cells)

    return {
        "source": str(input_path),
        "n_rows": n_rows,
        "n_cols": n_cols,
        "column_widths": col_widths,
        "rows": rows_out,
        "origin_cells": origin_cells,
    }


def _collect_default_inputs(base_dir: Path) -> list[Path]:
    input_dir = base_dir.parent / "table_extractor" / "extract_results"
    input_dir.mkdir(parents=True, exist_ok=True)
    return sorted(path.resolve() for path in input_dir.glob("*.xml") if path.is_file())


def _wait_for_files(file_paths: list[Path], timeout_sec: float = 5.0) -> None:
    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        if all(path.exists() and path.stat().st_size > 0 for path in file_paths):
            return
        time.sleep(0.05)
    missing = [str(path) for path in file_paths if not path.exists() or path.stat().st_size <= 0]
    raise RuntimeError(f"output file not ready: {', '.join(missing)}")


def _parse_to_payload(input_path: Path) -> tuple[str, bytes]:
    parsed = parse_table_xml(input_path)
    output_name = f"{input_path.stem}_grid.json"
    payload = json.dumps(parsed, ensure_ascii=False, indent=2).encode("utf-8")
    return output_name, payload


def _write_payload(output_dir: Path, output_name: str, payload: bytes) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / output_name
    with output_path.open("wb") as f:
        f.write(payload)
        f.flush()
        os.fsync(f.fileno())

    try:
        dir_fd = os.open(str(output_dir), os.O_RDONLY)
        try:
            os.fsync(dir_fd)
        finally:
            os.close(dir_fd)
    except OSError:
        pass

    _wait_for_files([output_path])
    return output_path


T = TypeVar("T")


def _run_with_spinner(label: str, action: Callable[[], T]) -> T:
    stop = threading.Event()
    start_time = time.time()
    min_visible_sec = 0.8

    def _spin() -> None:
        frames = ("|", "/", "-", "\\")
        idx = 0
        print(f"[RUN] {label}", flush=True)
        while not stop.is_set():
            print(f"\r[RUN] {label} {frames[idx % len(frames)]}", end="", flush=True)
            idx += 1
            time.sleep(0.12)
        print(f"\r[DONE] {label}    ", flush=True)

    use_spinner = sys.stdout.isatty()
    spinner_thread: threading.Thread | None = None
    if use_spinner:
        spinner_thread = threading.Thread(target=_spin, daemon=True)
        spinner_thread.start()
    else:
        print(f"[RUN] {label}", flush=True)

    try:
        return action()
    finally:
        elapsed = time.time() - start_time
        if use_spinner and elapsed < min_visible_sec:
            time.sleep(min_visible_sec - elapsed)
        stop.set()
        if spinner_thread is not None:
            spinner_thread.join(timeout=1.0)
        elif not use_spinner:
            print(f"[DONE] {label}", flush=True)


def main(argv: list[str]) -> int:
    base_dir = Path(__file__).resolve().parent
    output_dir = base_dir / "parsing_results"
    output_dir.mkdir(parents=True, exist_ok=True)

    input_paths: list[Path] = []
    if len(argv) >= 2:
        for arg in argv[1:]:
            input_path = Path(arg).expanduser().resolve()
            if not input_path.exists() or not input_path.is_file():
                print(f"[ERROR] input file not found: {input_path}", file=sys.stderr)
                return 1
            input_paths.append(input_path)
    else:
        input_paths = _collect_default_inputs(base_dir)
        if not input_paths:
            print("[ERROR] no input XML files found in ../table_extractor/extract_results.", file=sys.stderr)
            return 1

    manifest: dict[str, object] = {
        "source_mode": (
            "single_input"
            if len(input_paths) == 1 and len(argv) >= 2
            else ("multi_input" if len(argv) >= 2 else "extract_results_default")
        ),
        "source_count": len(input_paths),
        "total_parsed_count": 0,
        "success_count": 0,
        "error_count": 0,
        "sources": [],
        "outputs": [],
    }

    success_count = 0
    error_count = 0

    for input_path in input_paths:
        try:
            output_name, payload = _run_with_spinner(
                f"SCRIPT: parsing {input_path.name}",
                lambda p=input_path: _parse_to_payload(input_path=p),
            )
            print(f"[STAGE] SCRIPT complete: {input_path.name}")
        except Exception as exc:
            error_count += 1
            print(f"[ERROR] script failed: {input_path.name} ({exc})", file=sys.stderr)
            manifest_sources = manifest["sources"]
            assert isinstance(manifest_sources, list)
            manifest_sources.append(
                {
                    "source": str(input_path),
                    "status": "script_error",
                    "error": str(exc),
                }
            )
            continue

        try:
            output_path = _run_with_spinner(
                f"WRITE: persisting {input_path.name}",
                lambda n=output_name, b=payload: _write_payload(output_dir=output_dir, output_name=n, payload=b),
            )
            print(f"[STAGE] WRITE complete: {input_path.name} (file={output_path.name})")
        except Exception as exc:
            error_count += 1
            print(f"[ERROR] write failed: {input_path.name} ({exc})", file=sys.stderr)
            manifest_sources = manifest["sources"]
            assert isinstance(manifest_sources, list)
            manifest_sources.append(
                {
                    "source": str(input_path),
                    "status": "write_error",
                    "error": str(exc),
                    "output_file": output_name,
                }
            )
            continue

        success_count += 1
        print(f"[OK] parsed 1 table xml: {input_path.name}")

        manifest_sources = manifest["sources"]
        assert isinstance(manifest_sources, list)
        manifest_sources.append(
            {
                "source": str(input_path),
                "status": "parsed",
                "output_file": output_path.name,
            }
        )

        manifest_outputs = manifest["outputs"]
        assert isinstance(manifest_outputs, list)
        manifest_outputs.append(
            {
                "source": str(input_path),
                "file": output_path.name,
            }
        )

    manifest["total_parsed_count"] = success_count
    manifest["success_count"] = success_count
    manifest["error_count"] = error_count
    manifest_path = output_dir / "manifest.json"
    manifest_payload = json.dumps(manifest, ensure_ascii=False, indent=2).encode("utf-8")
    _write_payload(output_dir=output_dir, output_name=manifest_path.name, payload=manifest_payload)

    print(
        f"[OK] parse completed: {success_count} parsed, {error_count} failed, "
        f"{len(input_paths)} input file(s)"
    )
    print(f"[OK] output directory: {output_dir}")
    return 1 if error_count > 0 else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
