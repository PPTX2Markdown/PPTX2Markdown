"""Unpack embedded OLE containers into the file a user actually attached.

PowerPoint stores an embedded file either as itself (``Microsoft_Word_Document.docx``)
or inside an OLE compound file (``oleObject1.bin``). Inside the compound file the
Packager wraps an arbitrary file in an ``Ole10Native`` stream that carries the
original file name; other servers keep their document in a ``Package`` or
``CONTENTS`` stream.
"""

from __future__ import annotations

import io
import struct
import zipfile
from dataclasses import dataclass
from pathlib import PurePosixPath, PureWindowsPath

import olefile

OLE10_NATIVE_STREAM = "\x01Ole10Native"
_DOCUMENT_STREAMS = ("Package", "CONTENTS")


@dataclass(frozen=True)
class UnpackedAttachment:
    name: str
    payload: bytes


def _c_string(data: bytes, offset: int) -> tuple[bytes, int]:
    end = data.find(b"\x00", offset)
    if end < 0:
        raise ValueError("unterminated Ole10Native string")
    return data[offset:end], end + 1


def _file_name(raw: bytes) -> str:
    for encoding in ("utf-8", "cp1252"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = raw.decode("latin-1")
    # The stored name may be a Windows path; keep only its last component.
    return PurePosixPath(PureWindowsPath(text).name).name.strip()


def parse_ole10_native(data: bytes) -> UnpackedAttachment:
    """The original file name and bytes of a Packager ``Ole10Native`` stream."""

    if len(data) < 18:
        raise ValueError("truncated Ole10Native stream")
    declared = struct.unpack_from("<I", data, 0)[0]
    if declared > len(data) - 4:
        raise ValueError("Ole10Native declares more bytes than it holds")
    offset = 6
    name, offset = _c_string(data, offset)
    _, offset = _c_string(data, offset)  # source path at embedding time
    if offset + 8 > len(data):
        raise ValueError("truncated Ole10Native metadata")
    offset += 4
    temp_path_size = struct.unpack_from("<I", data, offset)[0]
    offset += 4
    if temp_path_size > len(data) - offset:
        raise ValueError("invalid Ole10Native temporary path size")
    offset += temp_path_size
    if offset + 4 > len(data):
        raise ValueError("truncated Ole10Native payload size")
    size = struct.unpack_from("<I", data, offset)[0]
    offset += 4
    if size > len(data) - offset:
        raise ValueError("invalid Ole10Native payload size")
    return UnpackedAttachment(_file_name(name) or "attachment.bin", data[offset : offset + size])


def payload_extension(data: bytes, prog_id: str = "") -> str:
    """A file extension for a document stream, from its bytes or the OLE program."""

    if data.startswith(b"%PDF-"):
        return ".pdf"
    if data.startswith(b"PK\x03\x04"):
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as archive:
                names = archive.namelist()
        except zipfile.BadZipFile:
            return ".zip"
        for prefix, extension in (("word/", ".docx"), ("xl/", ".xlsx"), ("ppt/", ".pptx")):
            if any(name.startswith(prefix) for name in names):
                return extension
        return ".zip"
    folded = prog_id.casefold()
    for program, extension in (("excel", ".xls"), ("word", ".doc"), ("powerpoint", ".ppt")):
        if program in folded:
            return extension
    return ".bin"


def unpack_ole(data: bytes, *, prog_id: str, stem: str) -> UnpackedAttachment | None:
    """The file inside an OLE compound file, or None when ``data`` is not one.

    Raises ``ValueError`` when ``data`` is a compound file that cannot be unpacked.
    """

    # Never hand olefile raw bytes: it reads bytes shorter than a minimal compound
    # file as a *file path*, which would let a package name a local file to open.
    if not data.startswith(olefile.MAGIC):
        return None
    try:
        with olefile.OleFileIO(io.BytesIO(data)) as ole:
            if ole.exists(OLE10_NATIVE_STREAM):
                return parse_ole10_native(ole.openstream(OLE10_NATIVE_STREAM).read())
            for stream in _DOCUMENT_STREAMS:
                if ole.exists(stream):
                    payload = ole.openstream(stream).read()
                    return UnpackedAttachment(
                        f"{stem}{payload_extension(payload, prog_id)}", payload
                    )
    except ValueError:
        raise
    except Exception as exc:  # olefile raises many types on malformed input
        raise ValueError(f"unreadable OLE compound file: {exc}") from exc
    raise ValueError("no Ole10Native, Package or CONTENTS stream")
