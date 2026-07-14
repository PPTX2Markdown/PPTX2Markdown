from __future__ import annotations

import filecmp
import logging
import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Dict, Iterable, Optional

from .ppt_to_pptx import _resolve_soffice_cmd

_NON_WEB_VECTOR_SUFFIXES = {".emf", ".wmf"}

logger = logging.getLogger(__name__)


def _same_content(path_a: Path, path_b: Path) -> bool:
    try:
        return filecmp.cmp(path_a, path_b, shallow=False)
    except OSError:
        return False


def _find_existing_copy(src: Path, dest_dir: Path) -> Optional[Path]:
    stem = src.stem
    suffix = src.suffix
    candidates = [dest_dir / src.name]
    candidates.extend(sorted(dest_dir.glob(f"{stem}-*{suffix}")))
    for candidate in candidates:
        if candidate.is_file() and _same_content(src, candidate):
            return candidate
    return None


def _copy_asset_to_dir(
    path: str,
    dest_dir: Optional[Path],
    copied_assets: Optional[Dict[str, Path]] = None,
) -> Optional[str]:
    if path.startswith("[unresolved-image") or dest_dir is None:
        return None

    src = Path(path)
    if not src.exists() or not src.is_file():
        return None

    try:
        src_key = str(src.resolve())
    except Exception:
        src_key = str(src)

    if copied_assets is not None and src_key in copied_assets:
        return str(copied_assets[src_key])

    dest_dir.mkdir(parents=True, exist_ok=True)
    existing_copy = _find_existing_copy(src, dest_dir)
    if existing_copy is not None:
        if copied_assets is not None:
            copied_assets[src_key] = existing_copy
        return str(existing_copy)

    dest = dest_dir / src.name
    if dest.exists():
        try:
            same_file = dest.resolve() == src.resolve()
        except Exception:
            same_file = False
        if not same_file:
            same_file = _same_content(src, dest)
        if not same_file:
            stem = src.stem
            suffix = src.suffix
            n = 2
            while dest.exists():
                candidate = dest_dir / f"{stem}-{n}{suffix}"
                if candidate.is_file() and _same_content(src, candidate):
                    dest = candidate
                    same_file = True
                    break
                dest = candidate
                n += 1
        if same_file:
            if copied_assets is not None:
                copied_assets[src_key] = dest
            return str(dest)

    shutil.copy2(src, dest)
    if copied_assets is not None:
        copied_assets[src_key] = dest
    return str(dest)


def convert_vector_assets_to_png(
    paths: Iterable[Path],
    dest_dir: Path,
) -> Dict[Path, Path]:
    """Convert all unique EMF/WMF assets in one LibreOffice process."""
    vector_paths: list[Path] = []
    seen_paths: set[Path] = set()
    for raw_path in paths:
        path = Path(raw_path)
        if path.suffix.casefold() not in _NON_WEB_VECTOR_SUFFIXES or not path.is_file():
            continue
        try:
            path = path.resolve()
        except OSError:
            pass
        if path in seen_paths:
            continue
        seen_paths.add(path)
        vector_paths.append(path)

    if not vector_paths:
        return {}

    candidates = [_resolve_soffice_cmd()]
    candidates.extend(shutil.which(name) for name in ("soffice", "libreoffice", "soffice.exe"))
    soffice_commands = list(dict.fromkeys(candidate for candidate in candidates if candidate))
    if not soffice_commands:
        logger.warning(
            "LibreOffice was not found; preserving %d original EMF/WMF asset(s)",
            len(vector_paths),
        )
        return {}

    with tempfile.TemporaryDirectory(prefix="pptx2markdown-vector-") as temporary:
        temporary_root = Path(temporary)
        input_dir = temporary_root / "input"
        input_dir.mkdir()
        staged_assets: list[tuple[Path, Path]] = []
        used_stems: set[str] = set()
        for source in vector_paths:
            stem = source.stem
            candidate_stem = stem
            collision_index = 2
            while candidate_stem.casefold() in used_stems:
                candidate_stem = f"{stem}-{collision_index}"
                collision_index += 1
            used_stems.add(candidate_stem.casefold())
            staged = input_dir / f"{candidate_stem}{source.suffix.casefold()}"
            shutil.copy2(source, staged)
            staged_assets.append((source, staged))

        for attempt, soffice_cmd in enumerate(soffice_commands):
            profile_dir = temporary_root / f"profile-{attempt}"
            output_dir = temporary_root / f"output-{attempt}"
            cache_dir = temporary_root / f"cache-{attempt}"
            profile_dir.mkdir()
            output_dir.mkdir()
            cache_dir.mkdir()
            command = [
                soffice_cmd,
                f"-env:UserInstallation={profile_dir.resolve().as_uri()}",
                "--headless",
                "--invisible",
                "--nodefault",
                "--nologo",
                "--nolockcheck",
                "--norestore",
                "--convert-to",
                "png",
                "--outdir",
                str(output_dir),
                *(str(staged) for _, staged in staged_assets),
            ]
            environment = os.environ.copy()
            environment["XDG_CACHE_HOME"] = str(cache_dir)
            try:
                subprocess.run(
                    command,
                    check=False,
                    capture_output=True,
                    text=True,
                    timeout=60,
                    env=environment,
                )
            except (OSError, subprocess.SubprocessError):
                continue
            converted_assets: Dict[Path, Path] = {}
            for source, staged in staged_assets:
                converted = output_dir / f"{staged.stem}.png"
                if not converted.is_file():
                    continue
                copied = _copy_asset_to_dir(str(converted), dest_dir)
                if copied is not None:
                    converted_assets[source] = Path(copied)
            if converted_assets:
                missing_count = len(vector_paths) - len(converted_assets)
                if missing_count:
                    logger.warning(
                        "LibreOffice did not convert %d of %d EMF/WMF asset(s); "
                        "preserving their original files",
                        missing_count,
                        len(vector_paths),
                    )
                return converted_assets

    logger.warning(
        "LibreOffice could not convert %d EMF/WMF asset(s); preserving originals",
        len(vector_paths),
    )
    return {}


def copy_media_asset(
    path: str,
    media_dir: Optional[Path],
    copied_media: Optional[Dict[str, Path]] = None,
) -> str:
    copied = _copy_asset_to_dir(path, dest_dir=media_dir, copied_assets=copied_media)
    if copied is None:
        return path
    return copied
