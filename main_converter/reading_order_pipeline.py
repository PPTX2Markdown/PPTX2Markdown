from __future__ import annotations

import json
import logging
import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Dict, Optional, Sequence, Tuple

from fs_utils import remove_tree_robust

logger = logging.getLogger(__name__)


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


def resolve_python_executable(repo_root: Path) -> str:
    """
    Pick a Python interpreter that is valid on the current OS.

    Priority:
    1) explicit override via PPTX2MARKDOWN_PYTHON
    2) the interpreter running the current process
    3) repo-local venv for the current platform
    """
    candidates = []

    env_python = str(os.environ.get("PPTX2MARKDOWN_PYTHON", "")).strip()
    if env_python:
        candidates.append(env_python)

    if sys.executable:
        candidates.append(sys.executable)

    if os.name == "nt":
        candidates.append(str(repo_root / ".venv" / "Scripts" / "python.exe"))
        candidates.append(str(repo_root / ".venv" / "python.exe"))
    else:
        candidates.append(str(repo_root / ".venv" / "bin" / "python"))

    for candidate in candidates:
        if _is_usable_python(candidate):
            return candidate

    raise RuntimeError(
        "could not find a usable Python interpreter. "
        "Run this tool with the intended virtualenv active, or set PPTX2MARKDOWN_PYTHON."
    )


def run_structure_analysis_stage(
    repo_root: Path,
    package_name: str,
    slide_xmls: Sequence[Path],
    strict: bool = False,
    mode: str = "xml",
    pptx_inheritance: str = "style",
    inherited_shapes: str = "visible",
) -> Tuple[Dict[str, Path], Path]:
    ro_script = repo_root / "structure_analyzer" / "extract_structure_analysis.py"
    if not ro_script.exists():
        raise FileNotFoundError(f"structure_analyzer script not found: {ro_script}")

    ro_output = repo_root / "structure_analyzer" / "output" / package_name
    if ro_output.exists():
        remove_tree_robust(ro_output)
    ro_output.mkdir(parents=True, exist_ok=True)

    py_exe = resolve_python_executable(repo_root)
    cmd = [
        py_exe,
        str(ro_script),
        "--mode",
        mode,
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
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", env=env)
    if proc.returncode != 0:
        raise RuntimeError(
            "structure_analysis stage failed\n"
            f"cmd: {' '.join(cmd)}\n"
            f"stdout:\n{proc.stdout}\n"
            f"stderr:\n{proc.stderr}"
        )

    manifest_path = ro_output / "structure_analysis_manifest.json"
    if not manifest_path.exists():
        legacy_manifest = ro_output / "reading_order_manifest.json"
        if legacy_manifest.exists():
            manifest_path = legacy_manifest
        else:
            raise FileNotFoundError(f"structure_analysis manifest not found: {manifest_path}")

    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    failed = manifest.get("failed", [])
    if isinstance(failed, list) and failed:
        raise RuntimeError(f"structure_analysis stage reported failures: {json.dumps(failed, ensure_ascii=False)}")

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


def resolve_surya_structure_dir(surya_root: Path, package_name: str) -> Path:
    def has_reordered_xmls(root: Path) -> bool:
        return root.exists() and root.is_dir() and any(root.glob("slide*.reordered.xml"))

    manifest_here = surya_root / "structure_analysis_manifest.json"
    if manifest_here.exists():
        return surya_root

    if has_reordered_xmls(surya_root):
        return surya_root

    candidate = surya_root / package_name
    if has_reordered_xmls(candidate):
        return candidate

    manifest_there = candidate / "structure_analysis_manifest.json"
    if manifest_there.exists():
        return candidate

    for nested_structure in (candidate / "05_structure_ready", candidate / "structure_ready"):
        if has_reordered_xmls(nested_structure):
            return nested_structure

        nested_manifest = nested_structure / "structure_analysis_manifest.json"
        if nested_manifest.exists():
            return nested_structure

    raise FileNotFoundError(
        f"surya structure-ready output not found for package '{package_name}' under {surya_root}"
    )


def run_surya_pipeline_stage(
    surya_root: Path,
    force: bool = False,
    targets: Optional[Sequence[str]] = None,
    target_pptx_dir: Optional[Path] = None,
    target_slides_dir: Optional[Path] = None,
) -> Path:
    run_script = surya_root / "run_surya_pipeline.py"
    if not run_script.exists():
        raise FileNotFoundError(f"surya pipeline script not found: {run_script}")

    repo_root = Path(__file__).resolve().parent.parent
    py_exe = resolve_python_executable(repo_root)
    cmd = [py_exe, str(run_script)]
    if target_pptx_dir is not None:
        cmd.extend(["--target-pptx-dir", str(target_pptx_dir)])
    if target_slides_dir is not None:
        cmd.extend(["--target-slides-dir", str(target_slides_dir)])
        cmd.append("--prefer-existing-target-slides")
    if force:
        cmd.append("--force")
    if targets:
        cmd.extend(str(t) for t in targets if str(t).strip())

    logger.info("[surya] Running pipeline: %s", " ".join(cmd))
    proc = subprocess.run(cmd, text=True, cwd=str(surya_root))
    if proc.returncode != 0:
        raise RuntimeError("surya pipeline failed\n" f"cmd: {' '.join(cmd)}\n")

    output_root = surya_root / "output"
    legacy_structure_root = output_root / "structure_ready"
    if output_root.exists() and output_root.is_dir():
        return output_root
    if legacy_structure_root.exists() and legacy_structure_root.is_dir():
        return legacy_structure_root
    raise FileNotFoundError(f"surya output not found after pipeline run: {output_root}")


def prepare_surya_structure_root(
    force: bool = False,
    reuse_existing_output: bool = False,
    targets: Optional[Sequence[str]] = None,
    target_pptx_dir: Optional[Path] = None,
    target_slides_dir: Optional[Path] = None,
) -> Path:
    repo_root = Path(__file__).resolve().parent.parent
    candidate = repo_root / "surya_pipeline"
    output_root = candidate / "output"
    structure_root = output_root / "structure_ready"
    analyzer_root = repo_root / "structure_analyzer" / "output"

    def has_structure_ready_output(root: Path) -> bool:
        return (
            root.exists()
            and root.is_dir()
            and (
                any(root.glob("slide*.reordered.xml"))
                or (root / "structure_analysis_manifest.json").exists()
            )
        )

    def has_all_target_outputs(root: Path) -> bool:
        target_names = [str(t).strip() for t in (targets or []) if str(t).strip()]
        if target_names:
            return all(
                has_structure_ready_output(root / name)
                or has_structure_ready_output(root / name / "05_structure_ready")
                or has_structure_ready_output(root / name / "structure_ready")
                for name in target_names
            )
        return has_structure_ready_output(root) or any(
            has_structure_ready_output(child)
            or has_structure_ready_output(child / "05_structure_ready")
            or has_structure_ready_output(child / "structure_ready")
            for child in root.glob("*")
        )

    if reuse_existing_output:
        if (candidate / "structure_analysis_manifest.json").exists():
            return candidate
        if has_all_target_outputs(output_root):
            return output_root
        if structure_root.exists() and structure_root.is_dir():
            return structure_root
        if has_all_target_outputs(analyzer_root):
            return analyzer_root
        raise FileNotFoundError(
            "reused surya output not found. Expected an existing cache under "
            f"{structure_root} or compatible reordered XML under {analyzer_root}. "
            "Run once without --reuse-surya-cache to regenerate Surya cache."
        )

    if (candidate / "run_surya_pipeline.py").exists():
        return run_surya_pipeline_stage(
            surya_root=candidate,
            force=force,
            targets=targets,
            target_pptx_dir=target_pptx_dir,
            target_slides_dir=target_slides_dir,
        )

    raise FileNotFoundError(
        f"valid surya input not found under fixed surya_pipeline path: {candidate}"
    )
