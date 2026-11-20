"""Transactional writer for portable presentation bundles."""

from __future__ import annotations

import hashlib
import json
import shutil
import tempfile
import uuid
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Literal

from .assets import safe_basename
from .errors import AssetUnavailableError, BundleWriteError, SaveConflictError
from .model import AttachmentBlock, ImageBlock, MarkdownBlockBase
from .render import render_presentation_json, render_presentation_markdown


@dataclass(frozen=True)
class SaveOptions:
    output_format: Literal["markdown", "json"] = "markdown"
    assets: Literal["copy", "omit"] = "copy"
    overwrite: Literal["error", "replace"] = "error"


@dataclass(frozen=True)
class SavedAsset:
    id: str
    path: Path
    sha256: str
    size: int


@dataclass(frozen=True)
class SaveResult:
    source: Path | None
    bundle_dir: Path
    document_path: Path
    manifest_path: Path
    assets: tuple[SavedAsset, ...]


def _referenced_asset_ids(document) -> set[str]:
    referenced: set[str] = set()
    for slide in document.slides:
        for block in slide.blocks:
            if isinstance(block, (ImageBlock, AttachmentBlock)):
                referenced.add(block.asset_id)
            elif isinstance(block, MarkdownBlockBase):
                referenced.update(block.asset_ids)
    return referenced


def _hashed_file(path: Path) -> tuple[str, int]:
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as source:
        while chunk := source.read(1024 * 1024):
            size += len(chunk)
            digest.update(chunk)
    return digest.hexdigest(), size


def _asset_relative_paths(document, referenced: set[str]) -> dict[str, PurePosixPath]:
    descriptors = {asset.id: asset for asset in document.assets}
    assigned: dict[str, PurePosixPath] = {}
    occupied: dict[PurePosixPath, str] = {}
    for asset_id in sorted(referenced):
        descriptor = descriptors[asset_id]
        directory = "media" if descriptor.kind == "image" else "attachments"
        fallback = "image.bin" if descriptor.kind == "image" else "attachment.bin"
        basename = safe_basename(descriptor.original_name, fallback=fallback)
        candidate = PurePosixPath(directory) / basename
        existing_hash = occupied.get(candidate)
        if existing_hash is not None and existing_hash != descriptor.sha256:
            stem = Path(basename).stem
            suffix = Path(basename).suffix
            candidate = PurePosixPath(directory) / (f"{stem}-{descriptor.sha256[:8]}{suffix}")
        occupied[candidate] = descriptor.sha256
        assigned[asset_id] = candidate
    return assigned


def _safe_bundle_path(bundle_dir: str | Path) -> Path:
    path = Path(bundle_dir).expanduser().resolve()
    if path.parent == path or path.name in {"", ".", ".."}:
        raise BundleWriteError(f"unsafe bundle destination: {path}")
    return path


def _commit_staging(
    staging: Path,
    destination: Path,
    *,
    overwrite: Literal["error", "replace"],
) -> None:
    if destination.exists():
        if overwrite == "error":
            raise SaveConflictError(f"bundle destination already exists: {destination}")
        backup = destination.parent / f".{destination.name}.backup-{uuid.uuid4().hex}"
        destination.rename(backup)
        try:
            staging.rename(destination)
        except Exception:
            if not destination.exists() and backup.exists():
                backup.rename(destination)
            raise
        else:
            if backup.is_dir():
                shutil.rmtree(backup)
            else:
                backup.unlink(missing_ok=True)
        return
    staging.rename(destination)


def save_bundle(
    presentation,
    bundle_dir: str | Path,
    *,
    options: SaveOptions | None = None,
) -> SaveResult:
    """Write, validate and commit a complete bundle."""

    options = options or SaveOptions()
    destination = _safe_bundle_path(bundle_dir)
    if destination.exists() and options.overwrite == "error":
        raise SaveConflictError(f"bundle destination already exists: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    staging = Path(tempfile.mkdtemp(prefix=f".{destination.name}.tmp-", dir=destination.parent))
    committed = False
    try:
        document = presentation.document
        stem = safe_basename(Path(document.source.name).stem, fallback="presentation")
        extension = ".md" if options.output_format == "markdown" else ".json"
        document_name = f"{stem}{extension}"
        document_path = staging / document_name
        manifest_path = staging / "bundle-manifest.json"

        referenced = _referenced_asset_ids(document)
        relative_paths = (
            _asset_relative_paths(document, referenced) if options.assets == "copy" else {}
        )
        saved_assets: list[SavedAsset] = []
        materialized: dict[PurePosixPath, str] = {}
        for asset_id, relative in sorted(relative_paths.items()):
            descriptor = presentation.assets.descriptor(asset_id)
            if not descriptor.payload_available:
                raise AssetUnavailableError(f"asset payload is unavailable: {asset_id}")
            output = staging.joinpath(*relative.parts)
            previous_hash = materialized.get(relative)
            if previous_hash is None:
                presentation.assets.materialize(asset_id, output)
                materialized[relative] = descriptor.sha256
            elif previous_hash != descriptor.sha256:
                raise BundleWriteError(f"asset path collision was not resolved: {relative}")
            actual_hash, actual_size = _hashed_file(output)
            if actual_hash != descriptor.sha256 or actual_size != descriptor.size:
                raise BundleWriteError(f"asset verification failed: {asset_id}")
            saved_assets.append(
                SavedAsset(
                    id=asset_id,
                    path=output,
                    sha256=actual_hash,
                    size=actual_size,
                )
            )

        if options.output_format == "markdown":
            output_text = render_presentation_markdown(
                document,
                asset_resolver=(
                    lambda asset_id: (
                        relative_paths[asset_id].as_posix() if asset_id in relative_paths else ""
                    )
                ),
                asset_mode="omit" if options.assets == "omit" else "error",
            )
        else:
            output_text = render_presentation_json(document)
        document_path.write_text(output_text, encoding="utf-8", newline="\n")

        manifest = {
            "schema_version": "1.0",
            "source": document.source.model_dump(mode="json"),
            "document": document_name,
            "output_format": options.output_format,
            "asset_mode": options.assets,
            "assets": {
                asset.id: {
                    "path": relative_paths[asset.id].as_posix(),
                    "sha256": asset.sha256,
                    "size": asset.size,
                    "media_type": asset.media_type,
                }
                for asset in document.assets
                if asset.id in relative_paths
            },
        }
        manifest_path.write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
            encoding="utf-8",
            newline="\n",
        )

        for relative in relative_paths.values():
            resolved = staging.joinpath(*relative.parts).resolve()
            if not resolved.is_relative_to(staging.resolve()) or not resolved.is_file():
                raise BundleWriteError(f"bundle asset path is invalid: {relative}")

        _commit_staging(staging, destination, overwrite=options.overwrite)
        committed = True
        return SaveResult(
            source=presentation.source_path,
            bundle_dir=destination,
            document_path=destination / document_name,
            manifest_path=destination / "bundle-manifest.json",
            assets=tuple(
                SavedAsset(
                    id=asset.id,
                    path=destination / asset.path.relative_to(staging),
                    sha256=asset.sha256,
                    size=asset.size,
                )
                for asset in saved_assets
            ),
        )
    except (SaveConflictError, AssetUnavailableError):
        raise
    except Exception as exc:
        raise BundleWriteError(f"could not write bundle {destination}: {exc}") from exc
    finally:
        if not committed and staging.exists():
            shutil.rmtree(staging, ignore_errors=True)
