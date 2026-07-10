#!/usr/bin/env python3
"""Extract OpenXML table nodes (<a:tbl>) from slide XML files.

Usage:
    ./extract.py [input_slide.xml ...]

Behavior:
    - Creates ./artifacts/extract_results if missing.
    - Writes extracted table XML files to ./artifacts/extract_results.
    - Output name format: <slide_stem>_0001.xml, <slide_stem>_0002.xml, ...
    - Writes manifest.json to ./artifacts/extract_results.
    - If input is omitted, scans ./target_slides for slide*.xml files.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
import sys
import threading
import time
from typing import Callable, TypeVar
from xml.dom import minidom
import xml.etree.ElementTree as ET

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

NSMAP = {
    "a": A_NS,
    "p": P_NS,
    "r": R_NS,
}

for prefix, uri in NSMAP.items():
    ET.register_namespace(prefix, uri)


def _usage() -> str:
    return "Usage: ./extract.py [input_slide.xml ...]"


def _pretty_xml_bytes(elem: ET.Element) -> bytes:
    raw = ET.tostring(elem, encoding="utf-8")
    pretty = minidom.parseString(raw).toprettyxml(indent="  ", encoding="utf-8")
    lines = [line for line in pretty.splitlines() if line.strip()]
    return b"\n".join(lines) + b"\n"


def _wait_for_files(file_paths: list[Path], timeout_sec: float = 5.0) -> None:
    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        if all(path.exists() and path.stat().st_size > 0 for path in file_paths):
            return
        time.sleep(0.05)
    missing = [str(path) for path in file_paths if not path.exists() or path.stat().st_size <= 0]
    raise RuntimeError(f"output file not ready: {', '.join(missing)}")


def _collect_table_payloads(input_path: Path) -> list[tuple[str, bytes]]:
    tree = ET.parse(input_path)
    root = tree.getroot()

    tables = []
    if root.tag == f"{{{A_NS}}}tbl":
        tables.append(root)
    tables.extend(root.findall(f".//{{{A_NS}}}tbl"))

    payloads: list[tuple[str, bytes]] = []
    slide_stem = input_path.stem
    for idx, table in enumerate(tables, start=1):
        filename = f"{slide_stem}_{idx:04d}.xml"
        payloads.append((filename, _pretty_xml_bytes(table)))
    return payloads


def _write_payloads(output_dir: Path, payloads: list[tuple[str, bytes]]) -> list[str]:
    output_dir.mkdir(parents=True, exist_ok=True)

    written_files: list[str] = []
    written_paths: list[Path] = []
    for filename, payload in payloads:
        out_path = output_dir / filename
        with out_path.open("wb") as f:
            f.write(payload)
            f.flush()
            os.fsync(f.fileno())
        written_files.append(filename)
        written_paths.append(out_path)

    # Ensure directory metadata is also flushed so file visibility is stable.
    try:
        dir_fd = os.open(str(output_dir), os.O_RDONLY)
        try:
            os.fsync(dir_fd)
        finally:
            os.close(dir_fd)
    except OSError:
        pass

    _wait_for_files(written_paths)
    return written_files


def _collect_default_inputs(base_dir: Path) -> list[Path]:
    target_dir = base_dir / "target_slides"
    target_dir.mkdir(parents=True, exist_ok=True)
    return sorted(path.resolve() for path in target_dir.glob("slide*.xml") if path.is_file())


T = TypeVar("T")


def _run_with_spinner(label: str, action: Callable[[], T]) -> T:
    stop = threading.Event()
    start_time = time.time()
    min_visible_sec = 0.8

    def _spin() -> None:
        frames = ("|", "/", "-", "\\")
        idx = 0
        # Print a visible "running" line first so users can see progress immediately.
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
    base_dir = Path.cwd()
    output_dir = base_dir / "artifacts" / "extract_results"
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
            print("[ERROR] no slide XML files found in ./target_slides (expected slide*.xml).", file=sys.stderr)
            return 1

    manifest: dict[str, object] = {
        "source_mode": (
            "single_input"
            if len(input_paths) == 1 and len(argv) >= 2
            else ("multi_input" if len(argv) >= 2 else "target_slides_default")
        ),
        "source_count": len(input_paths),
        "total_table_count": 0,
        "sources": [],
        "tables": [],
    }

    total_count = 0
    for input_path in input_paths:
        payloads = _run_with_spinner(
            f"SCRIPT: parsing {input_path.name}",
            lambda: _collect_table_payloads(input_path=input_path),
        )
        count = len(payloads)
        print(f"[STAGE] SCRIPT complete: {input_path.name} (tables={count})")

        written_files: list[str] = []
        if count == 0:
            print(f"[INFO] no table found: {input_path.name}")
            print(f"[STAGE] WRITE skipped: {input_path.name} (no outputs)")
        else:
            written_files = _run_with_spinner(
                f"WRITE: persisting {input_path.name}",
                lambda: _write_payloads(output_dir=output_dir, payloads=payloads),
            )
            print(f"[STAGE] WRITE complete: {input_path.name} (files={len(written_files)})")
            print(f"[OK] extracted {len(written_files)} table(s): {input_path.name}")

        total_count += len(written_files)

        manifest_sources = manifest["sources"]
        assert isinstance(manifest_sources, list)
        manifest_sources.append(
            {
                "source": str(input_path),
                "table_count": count,
                "files": written_files,
            }
        )

        manifest_tables = manifest["tables"]
        assert isinstance(manifest_tables, list)
        for idx, filename in enumerate(written_files, start=1):
            manifest_tables.append(
                {
                    "source": str(input_path),
                    "index": idx,
                    "file": filename,
                }
            )

    manifest["total_table_count"] = total_count
    (output_dir / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(f"[OK] extraction completed: {total_count} table(s) from {len(input_paths)} file(s)")
    print(f"[OK] output directory: {output_dir}")
    return 0


if __name__ == "__main__":
    print(
        "[INFO] extract.py is an internal pipeline module.\n"
        "Use `python run.py` from table_pipeline/ as the entrypoint."
    )
    raise SystemExit(1)
