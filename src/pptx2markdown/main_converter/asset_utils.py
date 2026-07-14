from __future__ import annotations

import filecmp
import shutil
from pathlib import Path
from typing import Dict, Optional


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


def copy_media_asset(
    path: str,
    media_dir: Optional[Path],
    copied_media: Optional[Dict[str, Path]] = None,
) -> str:
    copied = _copy_asset_to_dir(path, dest_dir=media_dir, copied_assets=copied_media)
    if copied is None:
        return path
    return copied
