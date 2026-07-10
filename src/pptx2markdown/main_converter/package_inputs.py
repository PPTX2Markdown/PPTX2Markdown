"""Prepare PPT/PPTX inputs for the main conversion pipeline."""

from __future__ import annotations

import json
import logging
import re
import shutil
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

from .converter_models import PreparedPackage
from .ppt_to_pptx import PptConversionError, convert_ppt_to_pptx

logger = logging.getLogger(__name__)


def default_target_dir(base_dir: Path) -> Path:
    """Return the directory that stores extracted PPTX packages."""
    local_target = base_dir / "target_slides"
    if local_target.exists() and not local_target.is_dir():
        raise NotADirectoryError(
            f"target_slides path exists but is not a directory: {local_target}"
        )
    local_target.mkdir(parents=True, exist_ok=True)
    return local_target


def default_pptx_input_dir(base_dir: Path) -> Path:
    """Return the default directory for source PPT/PPTX inputs."""
    local_target = base_dir / "target_pptx"
    if local_target.exists() and not local_target.is_dir():
        raise NotADirectoryError(f"target_pptx path exists but is not a directory: {local_target}")
    local_target.mkdir(parents=True, exist_ok=True)
    return local_target


def default_ppt_conversion_cache_dir(base_dir: Path) -> Path:
    """Return the cache directory used for legacy .ppt conversion."""
    cache_dir = base_dir / ".cache" / "ppt_to_pptx"
    if cache_dir.exists() and not cache_dir.is_dir():
        raise NotADirectoryError(
            f"ppt conversion cache path exists but is not a directory: {cache_dir}"
        )
    cache_dir.mkdir(parents=True, exist_ok=True)
    return cache_dir


def natural_key(name: str) -> Tuple[object, ...]:
    """Sort strings with embedded numbers in human order."""
    parts = re.split(r"(\d+)", name)
    key: List[object] = []
    for part in parts:
        key.append(int(part) if part.isdigit() else part.lower())
    return tuple(key)


def package_marker_matches(pkg_dir: Path, pptx_path: Path) -> bool:
    """Return whether an extracted package matches the current source file."""
    marker = pkg_dir / ".pptx_source.json"
    if not marker.exists():
        return False
    try:
        payload = json.loads(marker.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return False
    if not isinstance(payload, dict):
        return False

    slides_dir = pkg_dir / "ppt" / "slides"
    if not slides_dir.exists():
        return False
    try:
        stat = pptx_path.stat()
    except OSError:
        return False

    return (
        payload.get("source_path") == str(pptx_path.resolve())
        and payload.get("size") == stat.st_size
        and payload.get("mtime_ns") == stat.st_mtime_ns
    )


def safe_extract_pptx(pptx_path: Path, dest_dir: Path) -> None:
    """Extract a PPTX zip after rejecting path traversal entries."""
    with zipfile.ZipFile(pptx_path) as archive:
        for member in archive.infolist():
            member_path = Path(member.filename)
            if member_path.is_absolute() or ".." in member_path.parts:
                raise ValueError(f"Unsafe archive entry: {member.filename}")
        archive.extractall(dest_dir)


def extract_pptx_to_target(
    package: PreparedPackage,
    extraction_root: Path,
    allow_replace_unmanaged: bool = False,
) -> PreparedPackage:
    """Extract a source PPTX into the managed target_slides directory."""
    pptx_path = package.source_pptx_path
    stat = pptx_path.stat()
    extraction_root.mkdir(parents=True, exist_ok=True)

    pkg_name = package.package_dir_name
    pkg_dir = extraction_root / pkg_name
    pkg_dir = _resolve_extract_destination(
        pkg_name=pkg_name,
        pkg_dir=pkg_dir,
        extraction_root=extraction_root,
        allow_replace_unmanaged=allow_replace_unmanaged,
    )

    if package_marker_matches(pkg_dir, pptx_path):
        return package.with_package_dir(pkg_dir.resolve())

    marker = pkg_dir / ".pptx_source.json"
    if pkg_dir.exists():
        if marker.exists() or allow_replace_unmanaged:
            shutil.rmtree(pkg_dir)
        else:
            raise FileExistsError(
                "target package directory already exists and is not managed "
                f"by converter: {pkg_dir}"
            )

    pkg_dir.mkdir(parents=True, exist_ok=True)
    safe_extract_pptx(pptx_path, pkg_dir)

    slides_dir = pkg_dir / "ppt" / "slides"
    if not slides_dir.exists() or not slides_dir.is_dir():
        shutil.rmtree(pkg_dir, ignore_errors=True)
        raise ValueError(f"Not a valid pptx package after extraction: {pptx_path}")

    _write_package_marker(marker, pptx_path, stat.st_size, stat.st_mtime_ns)
    return package.with_package_dir(pkg_dir.resolve())


def _resolve_extract_destination(
    *,
    pkg_name: str,
    pkg_dir: Path,
    extraction_root: Path,
    allow_replace_unmanaged: bool,
) -> Path:
    if not pkg_dir.is_symlink():
        return pkg_dir

    try:
        pkg_dir.unlink()
        return pkg_dir
    except PermissionError:
        if not allow_replace_unmanaged:
            raise

    suffix = "__raw"
    index = 0
    while True:
        candidate_name = f"{pkg_name}{suffix}" if index == 0 else f"{pkg_name}{suffix}{index}"
        candidate = extraction_root / candidate_name
        if not candidate.exists() and not candidate.is_symlink():
            return candidate
        index += 1


def _write_package_marker(marker: Path, pptx_path: Path, size: int, mtime_ns: int) -> None:
    marker.write_text(
        json.dumps(
            {
                "source_path": str(pptx_path.resolve()),
                "size": size,
                "mtime_ns": mtime_ns,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


def is_ignored_presentation_file(path: Path) -> bool:
    """Return whether a file is an Office lock/temp presentation."""
    return path.name.startswith("~$")


def is_supported_presentation_file(path: Path) -> bool:
    """Return whether a file can be used as a presentation input."""
    return path.suffix.lower() in {".pptx", ".ppt"} and not is_ignored_presentation_file(path)


def resolve_input_presentation_path(cwd: Path, item: str) -> Tuple[Optional[Path], List[Path]]:
    """Resolve a CLI input string into an existing PPT/PPTX path."""
    search_roots = [Path.cwd(), cwd, default_pptx_input_dir(cwd), default_target_dir(cwd)]
    candidates: List[Path] = []
    seen: set[str] = set()

    def add_candidate(path: Path) -> None:
        key_path = path if path.is_absolute() else Path.cwd() / path
        key = str(key_path.resolve())
        if key not in seen:
            seen.add(key)
            candidates.append(path)

    raw_path = Path(item)
    add_candidate(raw_path)
    for root in search_roots:
        add_candidate(root / item)

    for candidate in candidates:
        found = _existing_supported_file(candidate)
        if found is not None:
            return found, candidates
    return None, candidates


def _existing_supported_file(candidate: Path) -> Optional[Path]:
    if candidate.exists() and candidate.is_file() and is_supported_presentation_file(candidate):
        return candidate.resolve()

    if candidate.suffix:
        return None

    for suffix in (".pptx", ".ppt"):
        with_suffix = candidate.with_suffix(suffix)
        if (
            with_suffix.exists()
            and with_suffix.is_file()
            and is_supported_presentation_file(with_suffix)
        ):
            return with_suffix.resolve()
    return None


def normalize_presentation_to_pptx(cwd: Path, path: Path, ppt_converter: str) -> Path:
    """Return a PPTX path, converting legacy .ppt inputs if needed."""
    suffix = path.suffix.lower()
    if suffix == ".pptx":
        return path
    if suffix != ".ppt":
        raise ValueError(f"unsupported presentation input: {path}")

    result = convert_ppt_to_pptx(
        path,
        default_ppt_conversion_cache_dir(cwd),
        mode=ppt_converter,  # type: ignore[arg-type]
    )
    action = "reused" if result.reused_cache else "converted"
    logger.info(
        "[ppt-convert] %s %s -> %s (%s)",
        action,
        path.name,
        result.pptx_path.name,
        result.converter,
    )
    return result.pptx_path


def prepare_package_inputs(
    cwd: Path,
    raw_inputs: Sequence[str],
    force_extract: bool = False,
    ppt_converter: str = "auto",
) -> Tuple[List[PreparedPackage], List[Dict[str, object]]]:
    """Resolve, normalize, and extract presentation inputs."""
    if not raw_inputs:
        return [], []

    prepared: List[PreparedPackage] = []
    missing_inputs: List[Dict[str, object]] = []
    extraction_root = default_target_dir(cwd)

    for item in raw_inputs:
        picked_file, candidates = resolve_input_presentation_path(cwd, item)
        if picked_file is None:
            missing_inputs.append(_missing_input_row(item, candidates))
            continue

        try:
            picked_file = normalize_presentation_to_pptx(cwd, picked_file, ppt_converter)
        except PptConversionError as exc:
            logger.error("[ppt-convert] failed: %s", exc)
            raise ValueError("ppt conversion failed") from exc

        package = PreparedPackage(
            package_dir=extraction_root / picked_file.stem,
            source_pptx_path=picked_file,
        )
        prepared.append(
            extract_pptx_to_target(
                package,
                extraction_root,
                allow_replace_unmanaged=force_extract,
            )
        )

    return prepared, missing_inputs


def _missing_input_row(item: str, candidates: Sequence[Path]) -> Dict[str, object]:
    return {
        "input": item,
        "checked": [
            str(path.resolve()) if path.is_absolute() else str((Path.cwd() / path).resolve())
            for path in candidates
        ],
    }


def collect_target_presentation_inputs() -> List[str]:
    """Collect presentation inputs from the user's current directory."""
    by_stem: Dict[str, Path] = {}

    for path in Path.cwd().iterdir():
        if not path.is_file() or not is_supported_presentation_file(path):
            continue
        resolved = path.resolve()
        key = path.stem.lower()
        existing = by_stem.get(key)
        if existing is None or (
            existing.suffix.lower() == ".ppt" and path.suffix.lower() == ".pptx"
        ):
            by_stem[key] = resolved

    files = sorted(by_stem.values(), key=lambda path: natural_key(path.name))
    return [str(path) for path in files]


def stage_surya_pptx_inputs(packages: Sequence[PreparedPackage], stage_dir: Path) -> Path:
    """Copy package sources into a temporary directory for the Surya pipeline."""
    stage_dir.mkdir(parents=True, exist_ok=True)
    for package in packages:
        staged_pptx = stage_dir / f"{package.name}.pptx"
        if staged_pptx.exists():
            staged_pptx.unlink()
        shutil.copy2(package.source_pptx_path, staged_pptx)
    return stage_dir
