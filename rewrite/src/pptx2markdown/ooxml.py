"""Secure, extraction-free access to OOXML presentation packages."""

from __future__ import annotations

import hashlib
import io
import posixpath
import stat
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import BinaryIO

from .errors import PackageSecurityError, PresentationParseError

RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


@dataclass(frozen=True)
class PackageLimits:
    """Limits enforced before reading any package part."""

    max_entries: int = 20_000
    max_total_uncompressed_bytes: int = 1024 * 1024 * 1024
    max_single_part_bytes: int = 256 * 1024 * 1024
    max_compression_ratio: float = 200.0

    def __post_init__(self) -> None:
        for name, value in vars(self).items():
            if value <= 0:
                raise ValueError(f"{name} must be positive")


@dataclass(frozen=True)
class Relationship:
    id: str
    type: str
    target: str
    external: bool = False


@dataclass(frozen=True)
class SlideEntry:
    page: int
    part_name: str
    relationship_id: str


def normalize_part_name(name: str) -> str:
    """Return a validated package-relative POSIX part name."""

    raw = str(name).replace("\\", "/")
    if raw.startswith("/"):
        raw = raw[1:]
    path = PurePosixPath(raw)
    if not raw or any(part in {"", ".", ".."} for part in path.parts):
        raise PackageSecurityError(f"unsafe OOXML part name: {name!r}")
    return path.as_posix()


def relationship_part_name(source_part: str) -> str:
    """Return the relationship part belonging to an OOXML source part."""

    source = PurePosixPath(normalize_part_name(source_part))
    return (source.parent / "_rels" / f"{source.name}.rels").as_posix()


def resolve_relationship_target(source_part: str, relationship: Relationship) -> str | None:
    """Resolve an internal relationship without allowing package traversal."""

    if relationship.external:
        return None
    target = relationship.target.replace("\\", "/")
    if target.startswith("/"):
        candidate = posixpath.normpath(target).lstrip("/")
    else:
        source = PurePosixPath(normalize_part_name(source_part))
        candidate = posixpath.normpath((source.parent / target).as_posix())
    if candidate == ".." or candidate.startswith("../") or "/../" in f"/{candidate}/":
        raise PackageSecurityError(
            f"relationship escapes package: {source_part!r} -> {relationship.target!r}"
        )
    return normalize_part_name(candidate)


def validate_archive_entries(
    archive: zipfile.ZipFile, limits: PackageLimits
) -> dict[str, zipfile.ZipInfo]:
    """Check every ZIP entry against the limits and return files by part name."""

    infos = archive.infolist()
    if len(infos) > limits.max_entries:
        raise PackageSecurityError(
            f"package contains {len(infos)} entries; limit is {limits.max_entries}"
        )

    total = 0
    seen: set[str] = set()
    validated: dict[str, zipfile.ZipInfo] = {}
    for info in infos:
        name = normalize_part_name(info.filename.rstrip("/"))
        if name in seen:
            raise PackageSecurityError(f"duplicate OOXML part: {name}")
        seen.add(name)

        if info.flag_bits & 0x1:
            raise PackageSecurityError(f"encrypted OOXML part is not allowed: {name}")
        mode = (info.external_attr >> 16) & 0xFFFF
        if stat.S_ISLNK(mode):
            raise PackageSecurityError(f"symlink OOXML entry is not allowed: {name}")
        if info.file_size > limits.max_single_part_bytes:
            raise PackageSecurityError(
                f"OOXML part exceeds {limits.max_single_part_bytes} bytes: {name}"
            )
        total += info.file_size
        if total > limits.max_total_uncompressed_bytes:
            raise PackageSecurityError(
                f"package uncompressed size exceeds {limits.max_total_uncompressed_bytes} bytes"
            )
        if info.file_size:
            ratio = info.file_size / max(info.compress_size, 1)
            if ratio > limits.max_compression_ratio:
                raise PackageSecurityError(
                    f"OOXML part compression ratio {ratio:.1f} exceeds "
                    f"{limits.max_compression_ratio:.1f}: {name}"
                )
        if not info.is_dir():
            validated[name] = info
    return validated


def reject_xml_declarations(data: bytes, name: str) -> None:
    """Refuse XML that declares a DTD or entities."""

    folded = data[:4096].upper()
    if b"<!DOCTYPE" in folded or b"<!ENTITY" in folded:
        raise PackageSecurityError(f"DTD and entity declarations are not allowed: {name}")


def validate_embedded_package(data: bytes, limits: PackageLimits, name: str) -> None:
    """Apply the package checks to an OOXML file embedded in another package."""

    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            infos = validate_archive_entries(archive, limits)
            for part_name, info in infos.items():
                if part_name.endswith((".xml", ".rels")):
                    with archive.open(info) as source:
                        reject_xml_declarations(source.read(4096), f"{name}!{part_name}")
    except zipfile.BadZipFile as exc:
        raise PackageSecurityError(f"embedded package is not a valid ZIP: {name}") from exc


class OOXMLPackage:
    """Validated read-only OOXML package."""

    def __init__(self, path: str | Path, limits: PackageLimits | None = None) -> None:
        self.path = Path(path)
        self.limits = limits or PackageLimits()
        self._archive: zipfile.ZipFile | None = None
        self._infos: dict[str, zipfile.ZipInfo] = {}
        self._relationships: dict[str, dict[str, Relationship]] = {}

    @property
    def source_sha256(self) -> str:
        digest = hashlib.sha256()
        with self.path.open("rb") as source:
            while chunk := source.read(1024 * 1024):
                digest.update(chunk)
        return digest.hexdigest()

    def open(self) -> "OOXMLPackage":
        if self._archive is not None:
            return self
        try:
            archive = zipfile.ZipFile(self.path)
        except (OSError, zipfile.BadZipFile) as exc:
            raise PresentationParseError(self.path, str(exc)) from exc
        self._archive = archive
        try:
            self._validate_archive()
        except Exception:
            archive.close()
            self._archive = None
            raise
        return self

    def _validate_archive(self) -> None:
        self._infos = validate_archive_entries(self._require_archive(), self.limits)

        for required in ("[Content_Types].xml", "ppt/presentation.xml"):
            if required not in self._infos:
                raise PresentationParseError(self.path, f"missing required OOXML part: {required}")

    def exists(self, part_name: str) -> bool:
        return normalize_part_name(part_name) in self._infos

    def info(self, part_name: str) -> zipfile.ZipInfo:
        name = normalize_part_name(part_name)
        try:
            return self._infos[name]
        except KeyError as exc:
            raise KeyError(f"missing OOXML part: {name}") from exc

    def open_part(self, part_name: str) -> BinaryIO:
        archive = self._require_archive()
        return archive.open(self.info(part_name), "r")

    def read(self, part_name: str) -> bytes:
        name = normalize_part_name(part_name)
        try:
            with self.open_part(name) as source:
                return source.read()
        except (OSError, RuntimeError, zipfile.BadZipFile) as exc:
            raise PresentationParseError(self.path, f"cannot read {name}: {exc}") from exc

    def xml(self, part_name: str) -> ET.Element:
        name = normalize_part_name(part_name)
        data = self.read(name)
        reject_xml_declarations(data, name)
        try:
            return ET.fromstring(data)
        except ET.ParseError as exc:
            raise PresentationParseError(self.path, f"invalid XML part {name}: {exc}") from exc

    def relationships(self, source_part: str) -> dict[str, Relationship]:
        source = normalize_part_name(source_part)
        cached = self._relationships.get(source)
        if cached is not None:
            return cached
        rels_name = relationship_part_name(source)
        if not self.exists(rels_name):
            self._relationships[source] = {}
            return {}
        root = self.xml(rels_name)
        relationships: dict[str, Relationship] = {}
        for element in root.findall(f"{{{RELATIONSHIPS_NS}}}Relationship"):
            rid = element.attrib.get("Id", "").strip()
            target = element.attrib.get("Target", "").strip()
            rel_type = element.attrib.get("Type", "").strip()
            if not rid or not target or not rel_type:
                raise PresentationParseError(
                    self.path, f"invalid relationship entry in {rels_name}"
                )
            if rid in relationships:
                raise PresentationParseError(
                    self.path, f"duplicate relationship id {rid!r} in {rels_name}"
                )
            relationships[rid] = Relationship(
                id=rid,
                type=rel_type,
                target=target,
                external=element.attrib.get("TargetMode", "").casefold() == "external",
            )
        self._relationships[source] = relationships
        return relationships

    def related_part(self, source_part: str, relationship_id: str) -> str | None:
        relationship = self.relationships(source_part).get(relationship_id)
        if relationship is None:
            return None
        target = resolve_relationship_target(source_part, relationship)
        if target is None or not self.exists(target):
            return None
        return target

    def slide_entries(self) -> tuple[SlideEntry, ...]:
        root = self.xml("ppt/presentation.xml")
        relationships = self.relationships("ppt/presentation.xml")
        entries: list[SlideEntry] = []
        ns = {"p": PRESENTATION_NS, "r": OFFICE_REL_NS}
        for page, element in enumerate(root.findall("./p:sldIdLst/p:sldId", ns), 1):
            rid = element.attrib.get(f"{{{OFFICE_REL_NS}}}id", "")
            relationship = relationships.get(rid)
            if relationship is None or not relationship.type.endswith("/slide"):
                raise PresentationParseError(
                    self.path, f"slide {page} has no valid slide relationship"
                )
            target = resolve_relationship_target("ppt/presentation.xml", relationship)
            if target is None or not self.exists(target):
                raise PresentationParseError(
                    self.path, f"slide {page} target is missing: {relationship.target}"
                )
            entries.append(SlideEntry(page=page, part_name=target, relationship_id=rid))
        return tuple(entries)

    def _require_archive(self) -> zipfile.ZipFile:
        if self._archive is None:
            raise RuntimeError("OOXML package is not open")
        return self._archive

    def close(self) -> None:
        if self._archive is not None:
            self._archive.close()
            self._archive = None

    def __enter__(self) -> "OOXMLPackage":
        return self.open()

    def __exit__(self, *_: object) -> None:
        self.close()
