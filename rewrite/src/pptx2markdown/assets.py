"""Managed binary asset storage with no externally visible temporary paths."""

from __future__ import annotations

import hashlib
import mimetypes
import re
import shutil
import tempfile
import threading
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import BinaryIO, Iterator, Literal

from .errors import AssetLimitError, AssetUnavailableError
from .model import AssetDescriptor

_CHUNK_SIZE = 1024 * 1024


@dataclass(frozen=True)
class AssetLimits:
    """Resource limits applied while ingesting embedded payloads."""

    max_assets: int = 10_000
    max_single_asset_bytes: int = 128 * 1024 * 1024
    max_total_asset_bytes: int = 512 * 1024 * 1024
    spool_threshold_bytes: int = 8 * 1024 * 1024

    def __post_init__(self) -> None:
        for name, value in vars(self).items():
            if value <= 0:
                raise ValueError(f"{name} must be positive")


@dataclass
class _Payload:
    stream: BinaryIO
    size: int
    sha256: str
    references: int = 1


def safe_basename(name: str, *, fallback: str) -> str:
    """Return a portable basename suitable for an output bundle."""

    normalized = str(name or "").replace("\\", "/")
    candidate = PurePosixPath(normalized).name.strip()
    candidate = re.sub(r"[\x00-\x1f\x7f]+", "", candidate)
    candidate = re.sub(r"[/\\:]+", "_", candidate)
    candidate = candidate.strip(" .")
    if candidate in {"", ".", ".."}:
        candidate = fallback
    stem = Path(candidate).stem.casefold()
    if stem in {
        "con",
        "prn",
        "aux",
        "nul",
        *(f"com{i}" for i in range(1, 10)),
        *(f"lpt{i}" for i in range(1, 10)),
    }:
        candidate = f"_{candidate}"
    return candidate


def safe_source_part(source_part: str) -> str:
    """Normalize and validate an OOXML package-relative part name."""

    raw = str(source_part).replace("\\", "/").lstrip("/")
    path = PurePosixPath(raw)
    if not raw or any(part in {"", ".", ".."} for part in path.parts):
        raise ValueError(f"unsafe asset source part: {source_part!r}")
    return path.as_posix()


_ZIP_DOCUMENTS = {
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".docm": "application/vnd.ms-word.document.macroEnabled.12",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".pptm": "application/vnd.ms-powerpoint.presentation.macroEnabled.12",
    ".odt": "application/vnd.oasis.opendocument.text",
    ".ods": "application/vnd.oasis.opendocument.spreadsheet",
    ".odp": "application/vnd.oasis.opendocument.presentation",
    ".hwpx": "application/hwp+zip",
}


def media_type_for(name: str, data_prefix: bytes = b"") -> str:
    """Infer a conservative media type from magic bytes and the filename."""

    if data_prefix.startswith(b"PK\x03\x04"):
        # Office and OpenDocument files are ZIP archives; their extension is more specific.
        document = _ZIP_DOCUMENTS.get(PurePosixPath(name).suffix.casefold())
        if document:
            return document

    signatures = (
        (b"\x89PNG\r\n\x1a\n", "image/png"),
        (b"\xff\xd8\xff", "image/jpeg"),
        (b"GIF87a", "image/gif"),
        (b"GIF89a", "image/gif"),
        (b"RIFF", "image/webp"),
        (b"%PDF-", "application/pdf"),
        (b"PK\x03\x04", "application/zip"),
    )
    for signature, media_type in signatures:
        if data_prefix.startswith(signature):
            if signature == b"RIFF" and data_prefix[8:12] != b"WEBP":
                continue
            return media_type
    guessed, _ = mimetypes.guess_type(name)
    return guessed or "application/octet-stream"


class AssetStore:
    """Own binary payloads for the lifetime of a parsed presentation."""

    def __init__(self, limits: AssetLimits | None = None) -> None:
        self._limits = limits or AssetLimits()
        self._descriptors: dict[str, AssetDescriptor] = {}
        self._payload_by_asset: dict[str, _Payload] = {}
        self._payload_by_hash: dict[str, _Payload] = {}
        self._logical_keys: dict[str, str] = {}
        self._total_unique_bytes = 0
        self._closed = False
        self._lock = threading.RLock()

    @property
    def closed(self) -> bool:
        return self._closed

    def _ensure_open(self) -> None:
        if self._closed:
            raise AssetUnavailableError("asset store is closed")

    @staticmethod
    def _asset_id(logical_key: str, collision_seed: str | None = None) -> str:
        digest = hashlib.sha256(logical_key.encode("utf-8")).hexdigest()[:16]
        base = f"asset-{digest}"
        if collision_seed is None:
            return base
        suffix = hashlib.sha256(collision_seed.encode("utf-8")).hexdigest()[:8]
        return f"{base}-{suffix}"

    def _allocate_id(self, logical_key: str, content_hash: str) -> str:
        asset_id = self._asset_id(logical_key)
        previous_key = self._logical_keys.get(asset_id)
        if previous_key is None or previous_key == logical_key:
            self._logical_keys[asset_id] = logical_key
            return asset_id
        asset_id = self._asset_id(logical_key, content_hash)
        previous_key = self._logical_keys.get(asset_id)
        if previous_key not in {None, logical_key}:
            raise AssetLimitError("deterministic asset id collision")
        self._logical_keys[asset_id] = logical_key
        return asset_id

    def register_bytes(
        self,
        data: bytes,
        *,
        logical_key: str,
        kind: Literal["image", "attachment"],
        original_name: str,
        source_part: str,
        media_type: str | None = None,
        preserve_payload: bool = True,
    ) -> AssetDescriptor:
        with memoryview(data) as view:
            return self.register_stream(
                _MemoryReader(view),
                logical_key=logical_key,
                kind=kind,
                original_name=original_name,
                source_part=source_part,
                media_type=media_type,
                preserve_payload=preserve_payload,
            )

    def register_stream(
        self,
        source: BinaryIO,
        *,
        logical_key: str,
        kind: Literal["image", "attachment"],
        original_name: str,
        source_part: str,
        media_type: str | None = None,
        preserve_payload: bool = True,
    ) -> AssetDescriptor:
        """Stream one asset into managed storage and return stable metadata."""

        with self._lock:
            self._ensure_open()
            if len(self._descriptors) >= self._limits.max_assets:
                raise AssetLimitError(f"asset count exceeds {self._limits.max_assets}")

            source_part = safe_source_part(source_part)
            fallback = "image.bin" if kind == "image" else "attachment.bin"
            original_name = safe_basename(original_name, fallback=fallback)
            spool = tempfile.SpooledTemporaryFile(
                max_size=self._limits.spool_threshold_bytes,
                mode="w+b",
            )
            digest = hashlib.sha256()
            size = 0
            prefix = b""
            try:
                while chunk := source.read(_CHUNK_SIZE):
                    size += len(chunk)
                    if size > self._limits.max_single_asset_bytes:
                        raise AssetLimitError(
                            f"asset exceeds {self._limits.max_single_asset_bytes} bytes"
                        )
                    if len(prefix) < 16:
                        prefix += chunk[: 16 - len(prefix)]
                    digest.update(chunk)
                    if preserve_payload:
                        spool.write(chunk)
                content_hash = digest.hexdigest()
                asset_id = self._allocate_id(logical_key, content_hash)
                existing = self._descriptors.get(asset_id)
                if existing is not None:
                    if existing.sha256 != content_hash or existing.size != size:
                        raise AssetLimitError(
                            f"logical asset key changed payload during parse: {logical_key}"
                        )
                    spool.close()
                    return existing

                payload: _Payload | None = None
                if preserve_payload:
                    payload = self._payload_by_hash.get(content_hash)
                    if payload is None:
                        if self._total_unique_bytes + size > self._limits.max_total_asset_bytes:
                            raise AssetLimitError(
                                "unique asset payloads exceed "
                                f"{self._limits.max_total_asset_bytes} bytes"
                            )
                        spool.seek(0)
                        payload = _Payload(spool, size, content_hash)
                        self._payload_by_hash[content_hash] = payload
                        self._total_unique_bytes += size
                    else:
                        payload.references += 1
                        spool.close()
                    self._payload_by_asset[asset_id] = payload
                else:
                    spool.close()

                descriptor = AssetDescriptor(
                    id=asset_id,
                    kind=kind,
                    original_name=original_name,
                    media_type=media_type or media_type_for(original_name, prefix),
                    size=size,
                    sha256=content_hash,
                    source_part=source_part,
                    payload_available=preserve_payload,
                )
                self._descriptors[asset_id] = descriptor
                return descriptor
            except Exception:
                spool.close()
                raise

    def descriptor(self, asset_id: str) -> AssetDescriptor:
        with self._lock:
            self._ensure_open()
            try:
                return self._descriptors[asset_id]
            except KeyError as exc:
                raise AssetUnavailableError(f"unknown asset: {asset_id}") from exc

    def descriptors(self) -> tuple[AssetDescriptor, ...]:
        with self._lock:
            self._ensure_open()
            return tuple(self._descriptors[key] for key in sorted(self._descriptors))

    def read_bytes(self, asset_id: str) -> bytes:
        with self._lock:
            payload = self._payload(asset_id)
            payload.stream.seek(0)
            data = payload.stream.read()
            payload.stream.seek(0)
            return data

    @contextmanager
    def open(self, asset_id: str) -> Iterator[BinaryIO]:
        """Yield the owned stream; callers must not close it."""

        with self._lock:
            payload = self._payload(asset_id)
            payload.stream.seek(0)
            try:
                yield payload.stream
            finally:
                payload.stream.seek(0)

    def materialize(self, asset_id: str, destination: Path) -> None:
        """Copy an asset payload to an explicit durable destination."""

        with self._lock:
            payload = self._payload(asset_id)
            destination.parent.mkdir(parents=True, exist_ok=True)
            payload.stream.seek(0)
            with destination.open("xb") as output:
                shutil.copyfileobj(payload.stream, output, length=_CHUNK_SIZE)
            payload.stream.seek(0)

    def _payload(self, asset_id: str) -> _Payload:
        self._ensure_open()
        if asset_id not in self._descriptors:
            raise AssetUnavailableError(f"unknown asset: {asset_id}")
        try:
            return self._payload_by_asset[asset_id]
        except KeyError as exc:
            raise AssetUnavailableError(f"asset payload is unavailable: {asset_id}") from exc

    def close(self) -> None:
        with self._lock:
            if self._closed:
                return
            for payload in self._payload_by_hash.values():
                payload.stream.close()
            self._payload_by_asset.clear()
            self._payload_by_hash.clear()
            self._closed = True

    def __enter__(self) -> "AssetStore":
        self._ensure_open()
        return self

    def __exit__(self, *_: object) -> None:
        self.close()


class _MemoryReader:
    """Minimal streaming reader over a memoryview without duplicating it."""

    def __init__(self, view: memoryview) -> None:
        self._view = view
        self._offset = 0

    def read(self, size: int = -1) -> bytes:
        if self._offset >= len(self._view):
            return b""
        if size < 0:
            end = len(self._view)
        else:
            end = min(self._offset + size, len(self._view))
        chunk = self._view[self._offset : end].tobytes()
        self._offset = end
        return chunk
