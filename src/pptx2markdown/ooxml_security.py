"""Security boundaries for OOXML archives and internal relationships."""

from __future__ import annotations

import shutil
import stat
import urllib.parse
import zipfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath


@dataclass(frozen=True)
class ArchiveSafetyLimits:
    """Conservative limits that prevent PPTX zip bombs from exhausting a host."""

    max_entries: int = 20_000
    max_total_uncompressed_bytes: int = 1_073_741_824  # 1 GiB
    max_member_uncompressed_bytes: int = 268_435_456  # 256 MiB
    max_compression_ratio: float = 200.0
    compression_ratio_min_bytes: int = 10_485_760  # 10 MiB


DEFAULT_ARCHIVE_LIMITS = ArchiveSafetyLimits()


def _normalized_member_path(filename: str) -> PurePosixPath:
    if not filename or "\x00" in filename:
        raise ValueError("unsafe archive entry: empty name or NUL byte")
    normalized = filename.replace("\\", "/")
    member_path = PurePosixPath(normalized)
    first_part = member_path.parts[0] if member_path.parts else ""
    if (
        member_path.is_absolute()
        or ".." in member_path.parts
        or (len(first_part) >= 2 and first_part[1] == ":")
    ):
        raise ValueError(f"unsafe archive entry: {filename}")
    return member_path


def _validate_archive(
    archive: zipfile.ZipFile,
    limits: ArchiveSafetyLimits,
) -> list[tuple[zipfile.ZipInfo, PurePosixPath]]:
    members = archive.infolist()
    if len(members) > limits.max_entries:
        raise ValueError(
            f"archive contains too many entries: {len(members)} > {limits.max_entries}"
        )

    total_size = 0
    seen_paths: set[str] = set()
    validated: list[tuple[zipfile.ZipInfo, PurePosixPath]] = []
    for member in members:
        member_path = _normalized_member_path(member.filename)
        normalized_name = member_path.as_posix().rstrip("/")
        if normalized_name in seen_paths:
            raise ValueError(f"duplicate archive entry: {member.filename}")
        seen_paths.add(normalized_name)

        if member.flag_bits & 0x1:
            raise ValueError(f"encrypted archive entry is not supported: {member.filename}")
        mode = (member.external_attr >> 16) & 0xFFFF
        if mode and stat.S_ISLNK(mode):
            raise ValueError(f"symbolic-link archive entry is not allowed: {member.filename}")
        if member.file_size > limits.max_member_uncompressed_bytes:
            raise ValueError(
                "archive entry exceeds uncompressed size limit: "
                f"{member.filename} ({member.file_size} bytes)"
            )

        total_size += member.file_size
        if total_size > limits.max_total_uncompressed_bytes:
            raise ValueError(f"archive exceeds total uncompressed size limit: {total_size} bytes")
        if (
            member.file_size >= limits.compression_ratio_min_bytes
            and member.compress_size > 0
            and member.file_size / member.compress_size > limits.max_compression_ratio
        ):
            raise ValueError(f"archive entry exceeds compression-ratio limit: {member.filename}")
        validated.append((member, member_path))
    return validated


def safe_extract_ooxml_archive(
    archive_path: Path,
    destination: Path,
    *,
    limits: ArchiveSafetyLimits = DEFAULT_ARCHIVE_LIMITS,
) -> None:
    """Extract an OOXML ZIP while enforcing paths, types, encryption, and sizes."""
    destination.mkdir(parents=True, exist_ok=True)
    root = destination.resolve()
    with zipfile.ZipFile(archive_path) as archive:
        members = _validate_archive(archive, limits)
        copied_total = 0
        for member, member_path in members:
            output_path = destination.joinpath(*member_path.parts)
            resolved_output = output_path.resolve(strict=False)
            try:
                resolved_output.relative_to(root)
            except ValueError as exc:
                raise ValueError(f"unsafe archive entry: {member.filename}") from exc

            if member.is_dir() or member.filename.endswith(("/", "\\")):
                output_path.mkdir(parents=True, exist_ok=True)
                continue
            output_path.parent.mkdir(parents=True, exist_ok=True)
            if output_path.is_symlink():
                raise ValueError(f"archive output is a symbolic link: {output_path}")
            with archive.open(member) as source, output_path.open("wb") as target:
                shutil.copyfileobj(source, target, length=1024 * 1024)
            # Use the resulting file size as a second boundary check instead of
            # trusting only the ZIP directory metadata.
            actual_size = output_path.stat().st_size
            if actual_size > limits.max_member_uncompressed_bytes:
                raise ValueError(
                    f"archive entry exceeded size limit while extracting: {member.filename}"
                )
            copied_total += actual_size
            if copied_total > limits.max_total_uncompressed_bytes:
                raise ValueError("archive exceeded total size limit while extracting")


def find_package_root(path: Path) -> Path | None:
    """Return the nearest extracted OOXML package root for ``path``."""
    start = path if path.is_dir() else path.parent
    for candidate in (start, *start.parents):
        if (candidate / "[Content_Types].xml").is_file():
            return candidate.resolve()
    return None


def _resolve_internal_target(base_dir: Path, package_root: Path, target: str) -> Path | None:
    parsed = urllib.parse.urlsplit(target)
    if parsed.scheme or parsed.netloc or parsed.query or parsed.fragment or not target.strip():
        return None
    decoded = urllib.parse.unquote(target).replace("\\", "/")
    if "\x00" in decoded:
        return None
    if decoded.startswith("/"):
        candidate = package_root / decoded.lstrip("/")
    else:
        candidate = base_dir / decoded
    resolved = candidate.resolve(strict=False)
    try:
        resolved.relative_to(package_root)
    except ValueError:
        return None
    return resolved


def resolve_package_part(
    source_part: Path,
    target: str,
    *,
    require_file: bool = True,
) -> Path | None:
    """Resolve a relationship target relative to a source OOXML part."""
    package_root = find_package_root(source_part)
    if package_root is None:
        return None
    resolved = _resolve_internal_target(source_part.parent, package_root, target)
    if resolved is None or (require_file and not resolved.is_file()):
        return None
    return resolved


def resolve_relationship_target(
    rels_path: Path,
    target: str,
    *,
    require_file: bool = True,
) -> Path | None:
    """Resolve a target from an OOXML ``_rels/*.rels`` sidecar."""
    package_root = find_package_root(rels_path)
    source_part_dir = rels_path.parent.parent
    if package_root is None:
        # Reordered slides live in
        # <work>/structure_analysis/<package>/_rels while their relationships
        # intentionally point back to <work>/target_slides/<package>. Keep that
        # generated sidecar layout usable without accepting arbitrary local paths.
        parsed = urllib.parse.urlsplit(target)
        if (
            parsed.scheme
            or parsed.netloc
            or parsed.query
            or parsed.fragment
            or not target.strip()
            or target.startswith("/")
        ):
            return None
        decoded = urllib.parse.unquote(target).replace("\\", "/")
        if "\x00" in decoded:
            return None
        resolved = (source_part_dir / decoded).resolve(strict=False)
        package_root = find_package_root(resolved)
        analysis_root = source_part_dir.parent
        if (
            package_root is None
            or analysis_root.name != "structure_analysis"
            or package_root.parent.name != "target_slides"
            or package_root.name != source_part_dir.name
            or package_root.parent.parent.resolve() != analysis_root.parent.resolve()
        ):
            return None
        try:
            resolved.relative_to(package_root)
        except ValueError:
            return None
        if require_file and not resolved.is_file():
            return None
        return resolved

    resolved = _resolve_internal_target(source_part_dir, package_root, target)
    if resolved is None or (require_file and not resolved.is_file()):
        return None
    return resolved
