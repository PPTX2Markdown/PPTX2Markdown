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
import os
import shutil
import subprocess
import sys
import tempfile
import unicodedata
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple
from xml.dom import minidom


def is_ignored_pptx_file(path: Path) -> bool:
    # Skip Office lock/temp files like "~$sample.pptx".
    return path.name.startswith("~$")


def run_cmd(
    cmd: List[str],
    cwd: Path,
    allow_fail: bool = False,
    env_overrides: Optional[Dict[str, str]] = None,
) -> bool:
    env = os.environ.copy()
    if env_overrides:
        env.update(env_overrides)
    proc = subprocess.run(cmd, cwd=str(cwd), text=True, capture_output=True, env=env)
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
    cmd = [
        sys.executable,
        "-c",
        launcher,
        str(input_path),
        "--output_dir",
        str(output_dir),
    ]
    try:
        run_cmd(cmd, cwd=cwd)
    except RuntimeError as err:
        msg = str(err)
        # Some large PDFs can fail on MPS/GPU path with torch.AcceleratorError.
        # Retry once on CPU for stability.
        if ("AcceleratorError" in msg) or ("index 8192 is out of bounds" in msg):
            print("  [2/4] surya_layout failed on accelerator; retrying on CPU...")
            run_cmd(
                cmd,
                cwd=cwd,
                env_overrides={
                    "TORCH_DEVICE": "cpu",
                    "PYTORCH_ENABLE_MPS_FALLBACK": "1",
                },
            )
            return
        raise


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
    return unicodedata.normalize("NFC", Path(s).stem)


def load_pptx_map_from_input_dir(input_dir: Path) -> Dict[str, Path]:
    """Build stem -> source pptx map from direct .pptx files under input_dir."""
    out: Dict[str, Path] = {}
    for p in sorted(input_dir.glob("*.pptx")):
        if not p.is_file() or is_ignored_pptx_file(p):
            continue
        out[unicodedata.normalize("NFC", p.stem)] = p
    return out


def stage_input_as_pptx(
    input_dir: Path,
    selected_stems: Set[str],
) -> Tuple[Path, Dict[str, Path]]:
    staged_dir = Path(tempfile.mkdtemp(prefix="surya_stage_pptx_"))
    direct_pptx_map = load_pptx_map_from_input_dir(input_dir)
    all_stems: Set[str] = set(direct_pptx_map.keys())
    if selected_stems:
        missing = sorted(stem for stem in selected_stems if stem not in all_stems)
        if missing:
            raise FileNotFoundError(
                f"requested --pptx not found in {input_dir}: {', '.join(missing)}"
            )
        all_stems = set(selected_stems)

    staged_map: Dict[str, Path] = {}
    for stem in sorted(all_stems):
        staged_pptx = staged_dir / f"{stem}.pptx"
        source_pptx = direct_pptx_map.get(stem)
        if source_pptx is not None:
            shutil.copy2(source_pptx, staged_pptx)
            staged_map[stem] = staged_pptx
            continue

    return staged_dir, staged_map


def is_ppt_root_dir(path: Path) -> bool:
    return path.exists() and path.is_dir() and (path / "ppt" / "slides").exists()


def has_flat_slide_xmls(path: Path) -> bool:
    return path.exists() and path.is_dir() and any(path.glob("slide*.xml"))


def build_temp_ppt_root_from_flat_slides(flat_slides_dir: Path, stem: str) -> Path:
    """
    Adapt flat slide xml directory (slide*.xml) into ppt-root shape:
      <temp>/<stem>/ppt/slides/slideN.xml
    """
    temp_root = Path(tempfile.mkdtemp(prefix=f"surya_flat_slides_{stem}_"))
    out_root = temp_root / stem
    out_slides = out_root / "ppt" / "slides"
    out_slides.mkdir(parents=True, exist_ok=True)

    for slide_xml in sorted(flat_slides_dir.glob("slide*.xml")):
        shutil.copy2(slide_xml, out_slides / slide_xml.name)

    src_rels_dir = flat_slides_dir / "_rels"
    if src_rels_dir.exists() and src_rels_dir.is_dir():
        out_rels_dir = out_slides / "_rels"
        out_rels_dir.mkdir(parents=True, exist_ok=True)
        for rel_file in sorted(src_rels_dir.glob("*.rels")):
            shutil.copy2(rel_file, out_rels_dir / rel_file.name)

    return out_root


def resolve_existing_ppt_root(
    stem: str,
    target_slides: Path,
) -> Optional[Path]:
    # 1) package-style target_slides/<stem>/ppt/slides
    candidate = target_slides / stem
    if is_ppt_root_dir(candidate):
        return candidate

    # 2) target_slides itself is already a ppt root
    if is_ppt_root_dir(target_slides):
        return target_slides

    return None


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
    parser.add_argument(
        "--target-pptx-dir",
        default=None,
        help=(
            "Input root for PPTX discovery. Only direct *.pptx files are supported. "
            "Default: <surya_pipeline>/target_pptx"
        ),
    )
    parser.add_argument(
        "--target-slides-dir",
        default=None,
        help="Directory for extracted pptx bundles. Default: <surya_pipeline>/target_slides",
    )
    parser.add_argument(
        "--prefer-existing-target-slides",
        action="store_true",
        help=(
            "Prefer existing slide XML roots from --target-slides-dir instead of re-extracting from pptx."
        ),
    )
    args = parser.parse_args()

    surya_dir = Path(__file__).resolve().parent

    target_pptx = Path(args.target_pptx_dir).resolve() if args.target_pptx_dir else (surya_dir / "target_pptx")
    target_pdf = surya_dir / "target_pdf"
    target_slides = Path(args.target_slides_dir).resolve() if args.target_slides_dir else (surya_dir / "target_slides")
    output_root = surya_dir / "output"

    for required_dir in [target_pdf, target_slides, output_root]:
        if required_dir.exists() and not required_dir.is_dir():
            raise NotADirectoryError(f"required path exists but is not a directory: {required_dir}")
        required_dir.mkdir(parents=True, exist_ok=True)
    if target_pptx.exists() and not target_pptx.is_dir():
        raise NotADirectoryError(f"target pptx path exists but is not a directory: {target_pptx}")
    target_pptx.mkdir(parents=True, exist_ok=True)

    selected_stems: Set[str] = {normalize_pptx_selector(x) for x in args.targets if normalize_pptx_selector(x)}
    staged_pptx_dir, pptx_map = stage_input_as_pptx(target_pptx, selected_stems)

    # Resolve per-stem ppt roots, preferring existing roots when requested.
    valid_pptx_map: Dict[str, Path] = {}
    ppt_root_map: Dict[str, Path] = {}
    temp_ppt_roots: List[Path] = []
    for stem, pptx_path in sorted(pptx_map.items()):
        ppt_root: Optional[Path] = None
        if args.prefer_existing_target_slides:
            ppt_root = resolve_existing_ppt_root(stem=stem, target_slides=target_slides)
            if ppt_root is None and has_flat_slide_xmls(target_slides):
                ppt_root = build_temp_ppt_root_from_flat_slides(target_slides, stem)
                temp_ppt_roots.append(ppt_root.parent)
                print(f"[slides] {stem}: using flat slide xmls from {target_slides} via temp ppt-root {ppt_root}")
            elif ppt_root is not None:
                print(f"[slides] {stem}: reusing existing ppt-root -> {ppt_root}")

        if ppt_root is None:
            # Side output: extract full pptx-xml bundle under target_slides/<stem>/
            try:
                xml_count = export_pptx_bundle_from_pptx(pptx_path, target_slides)
            except zipfile.BadZipFile:
                print(f"[WARN] invalid pptx zip, skipping: {pptx_path}")
                continue
            ppt_root = target_slides / stem
            print(f"[slides] {stem}: extracted pptx bundle ({xml_count} slides) -> {ppt_root}")

        try:
            _ = sorted((ppt_root / "ppt" / "slides").glob("slide*.xml"))
        except Exception:
            print(f"[WARN] failed to inspect ppt-root, skipping: {ppt_root}")
            continue
        valid_pptx_map[stem] = pptx_path
        ppt_root_map[stem] = ppt_root
    pptx_map = valid_pptx_map
    if not pptx_map:
        print("[INFO] No valid PPTX files available after slide-xml extraction.")
        if staged_pptx_dir.exists():
            shutil.rmtree(staged_pptx_dir, ignore_errors=True)
        for tmp in temp_ppt_roots:
            if tmp.exists():
                shutil.rmtree(tmp, ignore_errors=True)
        return 0

    # Step 1: PPTX -> PDF
    print("[1/4] Converting PPTX to PDF...")
    run_cmd(
        [
            sys.executable,
            "convert_pptx_to_pdf.py",
            "--input-dir",
            str(staged_pptx_dir),
            "--output-dir",
            str(target_pdf),
        ],
        cwd=surya_dir,
        allow_fail=True,
    )

    pdf_files = sorted(target_pdf.glob("*.pdf"))
    if selected_stems:
        pdf_files = [p for p in pdf_files if p.stem in selected_stems]
    if not pdf_files:
        print("[INFO] No PDF files found in target_pdf after conversion.")
        if staged_pptx_dir.exists():
            shutil.rmtree(staged_pptx_dir, ignore_errors=True)
        return 0

    try:
        __import__("surya")
    except Exception as e:  # noqa: BLE001
        raise RuntimeError("surya package not importable from current Python. Activate the project venv first.") from e

    # Step 2 + 3 + 4 per PDF
    for idx, pdf_path in enumerate(pdf_files, start=1):
        stem = pdf_path.stem
        print(f"[PDF {idx}/{len(pdf_files)}] {pdf_path.name}")

        package_dir = output_root / stem
        layout_result_json = package_dir / "results.json"
        layout_json = package_dir / "00_raw_surya_result.json"
        normalized_json = package_dir / "04_normalized.json"
        structure_dir = package_dir / "05_structure_ready"
        reordered_xmls = sorted(structure_dir.glob("slide*.reordered.xml"))

        # Step 2: layout
        if not layout_json.exists() and layout_result_json.exists():
            layout_result_json.replace(layout_json)
        if args.force or not layout_json.exists():
            print("  [2/4] Running surya_layout...")
            run_surya_cli("layout", pdf_path, output_root, cwd=surya_dir)
        else:
            print("  [2/4] Skip surya_layout (exists)")
        if layout_result_json.exists():
            layout_result_json.replace(layout_json)
        prettify_json_file(layout_json)

        # Step 3: normalize
        pptx_path = pptx_map.get(stem)
        if pptx_path is None:
            print(f"  [3/4] Skip normalize/build (matching pptx not found for stem: {stem})")
            continue
        if args.force or not normalized_json.exists():
            print("  [3/4] Running normalize_surya_results.py...")
            ppt_root = ppt_root_map.get(stem)
            if ppt_root is not None and ppt_root.exists():
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
            ppt_root = ppt_root_map.get(stem)
            if ppt_root is not None and ppt_root.exists():
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
    if staged_pptx_dir.exists():
        shutil.rmtree(staged_pptx_dir, ignore_errors=True)
    for tmp in temp_ppt_roots:
        if tmp.exists():
            shutil.rmtree(tmp, ignore_errors=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
