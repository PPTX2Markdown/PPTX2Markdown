from __future__ import annotations

import io
import re
import shutil
import struct
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Dict, Optional, Tuple

import olefile

from pptx2markdown.ooxml_security import resolve_relationship_target

OLE10_NATIVE_STREAM = "\x01Ole10Native"


def _read_c_string(data: bytes, offset: int) -> Tuple[bytes, int]:
    end = data.find(b"\x00", offset)
    if end < 0:
        raise ValueError("unterminated OLE native string")
    return data[offset:end], end + 1


def _decode_ole_filename(value: bytes) -> str:
    for encoding in ("utf-8", "cp1252", "latin-1"):
        try:
            return value.decode(encoding)
        except UnicodeDecodeError:
            continue
    return value.decode("latin-1", errors="replace")


def parse_ole10_native(data: bytes) -> Tuple[str, bytes]:
    """Extract the original filename and payload from a Packager native stream."""
    if len(data) < 18:
        raise ValueError("truncated Ole10Native stream")
    declared_size = struct.unpack_from("<I", data, 0)[0]
    if declared_size > len(data) - 4:
        raise ValueError("invalid Ole10Native declared size")
    offset = 6
    filename_raw, offset = _read_c_string(data, offset)
    _, offset = _read_c_string(data, offset)  # original source path
    if offset + 8 > len(data):
        raise ValueError("truncated Ole10Native metadata")
    offset += 4  # reserved flags
    temp_path_size = struct.unpack_from("<I", data, offset)[0]
    offset += 4
    if temp_path_size > len(data) - offset:
        raise ValueError("invalid Ole10Native temp path size")
    offset += temp_path_size
    if offset + 4 > len(data):
        raise ValueError("truncated Ole10Native payload size")
    payload_size = struct.unpack_from("<I", data, offset)[0]
    offset += 4
    if payload_size > len(data) - offset:
        raise ValueError("invalid Ole10Native payload size")
    filename = Path(_decode_ole_filename(filename_raw)).name
    if not filename:
        filename = "attachment.bin"
    return filename, data[offset : offset + payload_size]


def infer_payload_extension(data: bytes, prog_id: str = "") -> str:
    if data.startswith(b"%PDF-"):
        return ".pdf"
    if data.startswith(b"PK\x03\x04"):
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as archive:
                names = set(archive.namelist())
        except zipfile.BadZipFile:
            return ".zip"
        if any(name.startswith("word/") for name in names):
            return ".docx"
        if any(name.startswith("xl/") for name in names):
            return ".xlsx"
        if any(name.startswith("ppt/") for name in names):
            return ".pptx"
        return ".zip"
    folded = prog_id.casefold()
    if "excel" in folded:
        return ".xls"
    if "word" in folded:
        return ".doc"
    if "powerpoint" in folded:
        return ".ppt"
    return ".bin"


def _sanitize_filename(filename: str, fallback: str) -> str:
    name = Path(filename.replace("\\", "/")).name.strip()
    name = re.sub(r"[^\w.()-]+", "_", name, flags=re.UNICODE).strip("._")
    return name or fallback


def _relationship_part_path(rels_path: Path, target: str) -> Optional[Path]:
    return resolve_relationship_target(rels_path, target)


def _ole_payload(source: Path, prog_id: str, display_name: str) -> Tuple[str, bytes, str]:
    fallback_name = _sanitize_filename(display_name, source.stem)
    warning = ""
    try:
        with olefile.OleFileIO(source) as ole:
            if ole.exists(OLE10_NATIVE_STREAM):
                native = ole.openstream(OLE10_NATIVE_STREAM).read()
                filename, payload = parse_ole10_native(native)
                return filename, payload, warning
            if ole.exists("Package"):
                payload = ole.openstream("Package").read()
            elif ole.exists("CONTENTS"):
                payload = ole.openstream("CONTENTS").read()
            else:
                raise ValueError("no extractable Package, CONTENTS, or Ole10Native stream")
    except (OSError, IOError, ValueError) as exc:
        warning = f"OLE payload could not be unpacked; preserved container: {exc}"
        return f"{fallback_name}.bin", source.read_bytes(), warning

    extension = infer_payload_extension(payload, prog_id)
    return f"{fallback_name}{extension}", payload, warning


def _unique_attachment_path(directory: Path, filename: str, reserved: set[Path]) -> Path:
    stem = Path(filename).stem
    suffix = Path(filename).suffix
    candidate = directory / filename
    index = 2
    while candidate in reserved:
        candidate = directory / f"{stem}-{index}{suffix}"
        index += 1
    return candidate


def _copy_attachment(
    source: Path,
    *,
    prog_id: str,
    display_name: str,
    attachments_dir: Path,
    copied_attachments: Dict[str, Path],
) -> Tuple[Path, str]:
    key = str(source.resolve())
    existing = copied_attachments.get(key)
    if existing is not None:
        return existing, ""

    warning = ""
    if source.suffix.casefold() == ".bin" and olefile.isOleFile(source):
        filename, payload, warning = _ole_payload(source, prog_id, display_name)
        safe_name = _sanitize_filename(filename, "attachment.bin")
        attachments_dir.mkdir(parents=True, exist_ok=True)
        destination = _unique_attachment_path(
            attachments_dir, safe_name, set(copied_attachments.values())
        )
        destination.write_bytes(payload)
    else:
        safe_name = _sanitize_filename(source.name, "attachment.bin")
        attachments_dir.mkdir(parents=True, exist_ok=True)
        destination = _unique_attachment_path(
            attachments_dir, safe_name, set(copied_attachments.values())
        )
        shutil.copy2(source, destination)
    copied_attachments[key] = destination
    return destination, warning


def _attachment_markdown(path: Path, output_dir: Path) -> str:
    relative = path.relative_to(output_dir).as_posix()
    return f"[attachment: {path.name}]({relative})"


def convert_ole_attachment(
    graphic_frame: ET.Element,
    *,
    rels_path: Optional[Path],
    rels_map: Dict[str, str],
    output_dir: Optional[Path],
    attachments_dir: Optional[Path],
    copied_attachments: Optional[Dict[str, Path]],
    ns: Dict[str, str],
) -> Tuple[Optional[str], Optional[str]]:
    ole_obj = graphic_frame.find(".//p:oleObj", ns)
    if ole_obj is None:
        return None, "OLE object has no p:oleObj payload"
    rid = ole_obj.attrib.get(f"{{{ns['r']}}}id")
    target = rels_map.get(rid or "", "")
    if rels_path is None or not target:
        return None, "OLE object relationship is missing"
    source = _relationship_part_path(rels_path, target)
    if source is None:
        return None, f"OLE attachment target is unavailable: {target}"
    if output_dir is None or attachments_dir is None or copied_attachments is None:
        return None, "OLE attachment output directory is unavailable"
    destination, warning = _copy_attachment(
        source,
        prog_id=ole_obj.attrib.get("progId", ""),
        display_name=ole_obj.attrib.get("name", ""),
        attachments_dir=attachments_dir,
        copied_attachments=copied_attachments,
    )
    return _attachment_markdown(destination, output_dir), warning or None


def convert_media_attachment(
    picture: ET.Element,
    *,
    rels_path: Optional[Path],
    rels_map: Dict[str, str],
    output_dir: Optional[Path],
    attachments_dir: Optional[Path],
    copied_attachments: Optional[Dict[str, Path]],
    ns: Dict[str, str],
) -> Tuple[Optional[str], Optional[str]]:
    media = picture.find(".//a:audioFile", ns)
    if media is None:
        media = picture.find(".//a:videoFile", ns)
    if media is None:
        return None, None
    rid = media.attrib.get(f"{{{ns['r']}}}link")
    target = rels_map.get(rid or "", "")
    if rels_path is None or not target:
        return None, "embedded media relationship is missing"
    source = _relationship_part_path(rels_path, target)
    if source is None:
        return None, f"embedded media target is unavailable: {target}"
    if output_dir is None or attachments_dir is None or copied_attachments is None:
        return None, "embedded media output directory is unavailable"
    destination, warning = _copy_attachment(
        source,
        prog_id="",
        display_name=source.stem,
        attachments_dir=attachments_dir,
        copied_attachments=copied_attachments,
    )
    return _attachment_markdown(destination, output_dir), warning or None


def convert_model3d_attachment(
    graphic_frame: ET.Element,
    *,
    rels_path: Optional[Path],
    rels_map: Dict[str, str],
    output_dir: Optional[Path],
    attachments_dir: Optional[Path],
    copied_attachments: Optional[Dict[str, Path]],
    ns: Dict[str, str],
) -> Tuple[Optional[str], Optional[str]]:
    model = next(
        (
            element
            for element in graphic_frame.iter()
            if element.tag.rsplit("}", 1)[-1] == "model3d"
        ),
        None,
    )
    if model is None:
        return None, "3D model frame has no model3d payload"
    rid = model.attrib.get(f"{{{ns['r']}}}embed")
    target = rels_map.get(rid or "", "")
    if rels_path is None or not target:
        return None, "3D model relationship is missing"
    source = _relationship_part_path(rels_path, target)
    if source is None:
        return None, f"3D model target is unavailable: {target}"
    if output_dir is None or attachments_dir is None or copied_attachments is None:
        return None, "3D model attachment output directory is unavailable"
    destination, warning = _copy_attachment(
        source,
        prog_id="",
        display_name=source.stem,
        attachments_dir=attachments_dir,
        copied_attachments=copied_attachments,
    )
    return _attachment_markdown(destination, output_dir), warning or None
