#!/usr/bin/env python3
"""
End-to-end Surya pipeline for all PPTX files in target_pptx.

Pipeline:
1) convert_pptx_to_pdf.py: target_pptx -> target_pdf
2) surya_layout on each PDF
3) normalize_surya_results.py for each stem
4) build_structure_ready_from_normalized.py for each stem
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Dict, List, Set
from xml.dom import minidom


def run_cmd(cmd: List[str], cwd: Path, allow_fail: bool = False) -> bool:
    proc = subprocess.run(cmd, cwd=str(cwd), text=True, capture_output=True)
    if proc.returncode != 0:
        if allow_fail:
            print(
                "[WARN] command failed but continuing\n"
                f"cmd: {' '.join(cmd)}\n"
                f"stdout:\n{proc.stdout}\n"
                f"stderr:\n{proc.stderr}"
            )
            return False
        raise RuntimeError(
            "command failed\n"
            f"cmd: {' '.join(cmd)}\n"
            f"stdout:\n{proc.stdout}\n"
            f"stderr:\n{proc.stderr}"
        )
    if proc.stdout.strip():
        print(proc.stdout.strip())
    return True


def run_surya_cli(kind: str, input_path: Path, output_dir: Path, cwd: Path) -> None:
    cli_map = {
        "layout": ("surya.scripts.detect_layout", "detect_layout_cli"),
    }
    if kind not in cli_map:
        raise ValueError(f"unsupported surya cli kind: {kind}")
    module_name, cli_name = cli_map[kind]
    launcher = (
        "import sys; "
        f"from {module_name} import {cli_name} as cli; "
        "cli.main(args=sys.argv[1:], standalone_mode=False)"
    )
    run_cmd(
        [
            sys.executable,
            "-c",
            launcher,
            str(input_path),
            "--output_dir",
            str(output_dir),
        ],
        cwd=cwd,
    )


def prettify_json_file(json_path: Path) -> None:
    if not json_path.exists() or not json_path.is_file():
        return
    try:
        data = json.loads(json_path.read_text(encoding="utf-8"))
    except Exception:
        return
    json_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def normalize_pptx_selector(raw: str) -> str:
    s = (raw or "").strip()
    if not s:
        return s
    return Path(s).stem


def prettify_xml_file(xml_path: Path) -> None:
    try:
        raw = xml_path.read_bytes()
        dom = minidom.parseString(raw)
    except Exception:
        return
    pretty = dom.toprettyxml(indent="  ", encoding="utf-8")
    lines = pretty.decode("utf-8").splitlines()
    cleaned = "\n".join(line for line in lines if line.strip()) + "\n"
    xml_path.write_text(cleaned, encoding="utf-8")


def prettify_xml_tree(root_dir: Path) -> None:
    for xml_path in sorted(root_dir.rglob("*.xml")):
        prettify_xml_file(xml_path)


def export_pptx_bundle_from_pptx(pptx_path: Path, target_slides_root: Path) -> int:
    stem = pptx_path.stem
    dest_dir = target_slides_root / stem
    if dest_dir.exists():
        shutil.rmtree(dest_dir, ignore_errors=True)
    dest_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(pptx_path, "r") as zf:
        zf.extractall(dest_dir)
    prettify_xml_tree(dest_dir)

    slide_xmls = sorted((dest_dir / "ppt" / "slides").glob("slide*.xml"))
    return len(slide_xmls)


def main() -> int:
    parser = argparse.ArgumentParser(description="Run convert + Surya + normalize for target_pptx/*.pptx")
    parser.add_argument(
        "--force",
        action="store_true",
        help="Re-run layout/normalize/reorder even if output files already exist.",
    )
    parser.add_argument(
        "targets",
        nargs="*",
        help="Optional target PPTX selectors. If omitted, all target_pptx/*.pptx are processed.",
    )
    args = parser.parse_args()

    surya_dir = Path(__file__).resolve().parent

    target_pptx = surya_dir / "target_pptx"
    target_pdf = surya_dir / "target_pdf"
    target_slides = surya_dir / "target_slides"
    output_layout = surya_dir / "output" / "layout_result"
    output_norm = surya_dir / "output" / "normalized"
    output_struct = surya_dir / "output" / "structure_ready"

    for required_dir in [
        target_pptx,
        target_pdf,
        target_slides,
        output_layout,
        output_norm,
        output_struct,
    ]:
        if required_dir.exists() and not required_dir.is_dir():
            raise NotADirectoryError(f"required path exists but is not a directory: {required_dir}")
        required_dir.mkdir(parents=True, exist_ok=True)

    selected_stems: Set[str] = {normalize_pptx_selector(x) for x in args.targets if normalize_pptx_selector(x)}

    # Build stem -> pptx map
    pptx_map: Dict[str, Path] = {}
    for p in sorted(target_pptx.glob("*.pptx")):
        pptx_map[p.stem] = p

    if selected_stems:
        missing = sorted(stem for stem in selected_stems if stem not in pptx_map)
        if missing:
            raise FileNotFoundError(
                f"requested --pptx not found in {target_pptx}: {', '.join(missing)}"
            )
        pptx_map = {stem: path for stem, path in pptx_map.items() if stem in selected_stems}

    # Side output: extract full pptx-xml bundle under target_slides/<stem>/
    valid_pptx_map: Dict[str, Path] = {}
    for stem, pptx_path in sorted(pptx_map.items()):
        try:
            xml_count = export_pptx_bundle_from_pptx(pptx_path, target_slides)
        except zipfile.BadZipFile:
            print(f"[WARN] invalid pptx zip, skipping: {pptx_path}")
            continue
        valid_pptx_map[stem] = pptx_path
        print(f"[slides] {stem}: extracted pptx bundle ({xml_count} slides) -> {target_slides / stem}")
    pptx_map = valid_pptx_map
    if not pptx_map:
        print("[INFO] No valid PPTX files available after slide-xml extraction.")
        return 0

    # Step 1: PPTX -> PDF
    print("[1/4] Converting PPTX to PDF...")
    temp_input_dir: Path | None = None
    if selected_stems:
        temp_input_dir = Path(tempfile.mkdtemp(prefix="surya_select_pptx_"))
        for pptx_path in sorted(pptx_map.values()):
            shutil.copy2(pptx_path, temp_input_dir / pptx_path.name)
        run_cmd(
            [
                sys.executable,
                "convert_pptx_to_pdf.py",
                "--input-dir",
                str(temp_input_dir),
                "--output-dir",
                str(target_pdf),
            ],
            cwd=surya_dir,
            allow_fail=True,
        )
    else:
        run_cmd([sys.executable, "convert_pptx_to_pdf.py"], cwd=surya_dir, allow_fail=True)

    pdf_files = sorted(target_pdf.glob("*.pdf"))
    if selected_stems:
        pdf_files = [p for p in pdf_files if p.stem in selected_stems]
    if not pdf_files:
        print("[INFO] No PDF files found in target_pdf after conversion.")
        if temp_input_dir is not None and temp_input_dir.exists():
            shutil.rmtree(temp_input_dir, ignore_errors=True)
        return 0

    try:
        __import__("surya")
    except Exception as e:  # noqa: BLE001
        raise RuntimeError("surya package not importable from current Python. Activate the project venv first.") from e

    # Step 2 + 3 + 4 per PDF
    for idx, pdf_path in enumerate(pdf_files, start=1):
        stem = pdf_path.stem
        print(f"[PDF {idx}/{len(pdf_files)}] {pdf_path.name}")

        layout_json = output_layout / stem / "results.json"
        normalized_json = output_norm / f"{stem}_normalized.json"
        structure_dir = output_struct / stem
        reordered_xmls = sorted(structure_dir.glob("slide*.reordered.xml"))

        # Step 2: layout
        if args.force or not layout_json.exists():
            print("  [2/4] Running surya_layout...")
            run_surya_cli("layout", pdf_path, output_layout, cwd=surya_dir)
        else:
            print("  [2/4] Skip surya_layout (exists)")
        prettify_json_file(layout_json)

        # Step 3: normalize
        pptx_path = pptx_map.get(stem)
        if pptx_path is None:
            print(f"  [3/4] Skip normalize/build (matching pptx not found for stem: {stem})")
            continue
        if args.force or not normalized_json.exists():
            print("  [3/4] Running normalize_surya_results.py...")
            ppt_root = target_slides / stem
            if ppt_root.exists():
                ppt_locator_args = ["--ppt-root", str(ppt_root)]
            else:
                ppt_locator_args = ["--pptx-path", str(pptx_path)]
            run_cmd(
                [
                    sys.executable,
                    "normalize_surya_results.py",
                    "--layout-json",
                    str(layout_json),
                    *ppt_locator_args,
                    "--output-json",
                    str(normalized_json),
                ],
                cwd=surya_dir,
            )
        else:
            print("  [3/4] Skip normalize (exists)")
        prettify_json_file(normalized_json)

        # Step 4: structure-ready
        if args.force or not reordered_xmls:
            print("  [4/4] Running build_structure_ready_from_normalized.py...")
            ppt_root = target_slides / stem
            if ppt_root.exists():
                ppt_locator_args = ["--ppt-root", str(ppt_root)]
            else:
                ppt_locator_args = ["--pptx-path", str(pptx_path)]
            run_cmd(
                [
                    sys.executable,
                    "build_structure_ready_from_normalized.py",
                    "--normalized-json",
                    str(normalized_json),
                    *ppt_locator_args,
                    "--output-dir",
                    str(structure_dir),
                ],
                cwd=surya_dir,
            )
        else:
            print("  [4/4] Skip structure-ready (exists)")

    print("[DONE] Pipeline finished.")
    if temp_input_dir is not None and temp_input_dir.exists():
        shutil.rmtree(temp_input_dir, ignore_errors=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
