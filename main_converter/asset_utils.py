from __future__ import annotations

import shutil
from pathlib import Path
from typing import Dict, Optional


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
    dest = dest_dir / src.name
    if dest.exists():
        try:
            same_file = dest.resolve() == src.resolve()
        except Exception:
            same_file = False
        if not same_file:
            stem = src.stem
            suffix = src.suffix
            n = 2
            while dest.exists():
                dest = dest_dir / f"{stem}-{n}{suffix}"
                n += 1

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
