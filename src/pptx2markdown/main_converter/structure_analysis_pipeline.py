from __future__ import annotations

import json
import logging
import os
import shutil
import stat
import subprocess
import sys
import time
from pathlib import Path
from typing import Dict, Optional, Sequence, Tuple

logger = logging.getLogger(__name__)


def _clear_readonly_and_retry(func, path, exc_info) -> None:
    exc = exc_info[1]
    if not isinstance(exc, PermissionError):
        raise exc
    os.chmod(path, stat.S_IWRITE)
    func(path)


def _remove_tree_robust(path: Path, retries: int = 3, delay_sec: float = 0.2) -> None:
    last_error: Optional[Exception] = None
    for attempt in range(retries):
        try:
            shutil.rmtree(path, onerror=_clear_readonly_and_retry)
            return
        except FileNotFoundError:
            return
        except PermissionError as exc:
            last_error = exc
            if attempt + 1 == retries:
                break
            time.sleep(delay_sec)
    if last_error is not None:
        raise last_error


def _is_usable_python(executable: str) -> bool:
    candidate = str(executable or "").strip()
    if not candidate:
        return False
    try:
        proc = subprocess.run(
            [candidate, "-c", "import sys"],
            capture_output=True,
            text=True,
        )
    except OSError:
        return False
    return proc.returncode == 0


def resolve_python_executable() -> str:
    """
    Pick a Python interpreter that is valid on the current OS.

    Priority:
    1) explicit override via PPTX2MARKDOWN_PYTHON
    2) the interpreter running the current process
    """
    candidates = []

    env_python = str(os.environ.get("PPTX2MARKDOWN_PYTHON", "")).strip()
    if env_python:
        candidates.append(env_python)

    if sys.executable:
        candidates.append(sys.executable)

    for candidate in candidates:
        if _is_usable_python(candidate):
            return candidate

    raise RuntimeError(
        "could not find a usable Python interpreter. "
        "Run this tool with the intended virtualenv active, or set PPTX2MARKDOWN_PYTHON."
    )


def run_structure_analysis_stage(
    work_root: Path,
    package_name: str,
    slide_xmls: Sequence[Path],
    strict: bool = False,
    pptx_inheritance: str = "style",
    inherited_shapes: str = "visible",
) -> Tuple[Dict[str, Path], Path]:
    ro_output = work_root / "structure_analysis" / package_name
    if ro_output.exists():
        _remove_tree_robust(ro_output)
    ro_output.mkdir(parents=True, exist_ok=True)

    py_exe = resolve_python_executable()
    cmd = [
        py_exe,
        "-m",
        "pptx2markdown.structure_analyzer.extract_structure_analysis",
        "--output-dir",
        str(ro_output),
    ]
    if strict:
        cmd.append("--strict")
    cmd.extend(["--placeholder-inheritance", pptx_inheritance])
    cmd.extend(["--inherited-shapes", inherited_shapes])
    cmd.extend(str(p.resolve()) for p in slide_xmls)

    env = os.environ.copy()
    env.setdefault("PYTHONUTF8", "1")
    env.setdefault("PYTHONIOENCODING", "utf-8")
    proc = subprocess.run(
        cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", env=env
    )
    if proc.returncode != 0:
        raise RuntimeError(
            "structure_analysis stage failed\n"
            f"cmd: {' '.join(cmd)}\n"
            f"stdout:\n{proc.stdout}\n"
            f"stderr:\n{proc.stderr}"
        )

    manifest_path = ro_output / "structure_analysis_manifest.json"
    if not manifest_path.exists():
        raise FileNotFoundError(f"structure_analysis manifest not found: {manifest_path}")

    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    failed = manifest.get("failed", [])
    if isinstance(failed, list) and failed:
        raise RuntimeError(
            f"structure_analysis stage reported failures: {json.dumps(failed, ensure_ascii=False)}"
        )

    mapping: Dict[str, Path] = {}
    processed = manifest.get("processed", [])
    if isinstance(processed, list):
        for row in processed:
            if not isinstance(row, dict):
                continue
            src = row.get("input_xml")
            out = row.get("output_xml")
            if isinstance(src, str) and isinstance(out, str):
                mapping[str(Path(src).resolve())] = Path(out).resolve()
    return mapping, ro_output
