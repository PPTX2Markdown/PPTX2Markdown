#!/usr/bin/env python3
"""
Batch-convert PPTX files to PDF using LibreOffice (soffice).

Default behavior (script-relative):
- Input directory:  <this_file_dir>/target_pptx
- Output directory: <this_file_dir>/target_pdf
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import List


def resolve_soffice_cmd() -> str | None:
    candidates: List[Path] = []
    env_path = os.getenv("SOFFICE_PATH", "").strip()
    if env_path:
        candidates.append(Path(env_path))
    if sys.platform == "darwin":
        candidates.extend(
            [
                Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
                Path.home() / "Applications" / "LibreOffice.app" / "Contents" / "MacOS" / "soffice",
            ]
        )
    elif sys.platform.startswith("win"):
        candidates.extend(
            [
                Path("C:/Program Files/LibreOffice/program/soffice.exe"),
                Path("C:/Program Files (x86)/LibreOffice/program/soffice.exe"),
            ]
        )
    soffice_in_path = shutil.which("soffice")
    if soffice_in_path:
        candidates.append(Path(soffice_in_path))
    soffice_exe_in_path = shutil.which("soffice.exe")
    if soffice_exe_in_path:
        candidates.append(Path(soffice_exe_in_path))
    for cand in candidates:
        if cand.exists() and cand.is_file():
            return str(cand)
    return None


def convert_one(soffice_cmd: str, pptx_path: Path, output_dir: Path) -> subprocess.CompletedProcess[str]:
    profile_dir = Path(tempfile.mkdtemp(prefix="libreoffice-profile-"))
    cmd = [
        soffice_cmd,
        f"-env:UserInstallation={profile_dir.resolve().as_uri()}",
        "--headless",
        "--invisible",
        "--nodefault",
        "--nologo",
        "--nolockcheck",
        "--norestore",
        "--convert-to",
        "pdf:impress_pdf_Export",
        str(pptx_path),
        "--outdir",
        str(output_dir),
    ]
    env = os.environ.copy()
    if sys.platform == "darwin":
        env["HOME"] = "/tmp"
    try:
        return subprocess.run(cmd, capture_output=True, text=True, env=env)
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)


def main() -> int:
    script_dir = Path(__file__).resolve().parent
    default_input_dir = script_dir / "target_pptx"
    default_output_dir = script_dir / "target_pdf"

    parser = argparse.ArgumentParser(description="Convert .pptx files to .pdf in batch.")
    parser.add_argument(
        "--input-dir",
        default=str(default_input_dir),
        help="Directory containing .pptx files. Default is script-relative target_pptx.",
    )
    parser.add_argument(
        "--output-dir",
        default=str(default_output_dir),
        help="Directory where .pdf files are written. Default is script-relative target_pdf.",
    )
    args = parser.parse_args()

    input_dir = Path(args.input_dir).resolve()
    output_dir = Path(args.output_dir).resolve()

    if input_dir.exists() and not input_dir.is_dir():
        print(f"[ERROR] input path exists but is not a directory: {input_dir}")
        return 1
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    soffice_cmd = resolve_soffice_cmd()
    if not soffice_cmd:
        print("[ERROR] 'soffice' command not found.")
        print("Install LibreOffice and ensure 'soffice' is available in PATH.")
        return 1

    pptx_files: List[Path] = sorted(p for p in input_dir.glob("*.pptx") if p.is_file())
    if not pptx_files:
        print(f"[INFO] no .pptx files found in: {input_dir}")
        return 0

    failed = 0
    for pptx in pptx_files:
        proc = convert_one(soffice_cmd, pptx, output_dir)
        pdf_name = f"{pptx.stem}.pdf"
        pdf_path = output_dir / pdf_name
        if proc.returncode == 0 and pdf_path.exists():
            print(f"[OK] {pptx.name} -> {pdf_name}")
        else:
            failed += 1
            print(f"[FAIL] {pptx.name}")
            if proc.stdout.strip():
                print(f"  stdout: {proc.stdout.strip()}")
            if proc.stderr.strip():
                print(f"  stderr: {proc.stderr.strip()}")

    print(
        f"[SUMMARY] total={len(pptx_files)} success={len(pptx_files) - failed} failed={failed} output={output_dir}"
    )
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
