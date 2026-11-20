from __future__ import annotations

import io
import struct
import unittest
import zipfile
from pathlib import Path

import pptx2markdown
from pptx2markdown.assets import media_type_for
from pptx2markdown.attachments import parse_ole10_native, payload_extension, unpack_ole
from pptx2markdown.model import AttachmentBlock

GOLDEN = Path(__file__).resolve().parents[2] / "tests" / "fixtures" / "golden"


def _ole10_native(name: bytes, payload: bytes, temp_path: bytes = b"C:\\tmp\\x") -> bytes:
    body = (
        b"\x02\x00"
        + name
        + b"\x00"
        + b"C:\\src\\"
        + name
        + b"\x00"
        + b"\x00\x00\x03\x00"
        + struct.pack("<I", len(temp_path))
        + temp_path
        + struct.pack("<I", len(payload))
        + payload
    )
    return struct.pack("<I", len(body)) + body


def _zip(*names: str) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name in names:
            archive.writestr(name, "x")
    return buffer.getvalue()


class OleUnpackTests(unittest.TestCase):
    def test_ole10_native_yields_original_name_and_bytes(self) -> None:
        unpacked = parse_ole10_native(_ole10_native(b"C:\\Users\\me\\Result.json", b'{"a": 1}'))
        self.assertEqual((unpacked.name, unpacked.payload), ("Result.json", b'{"a": 1}'))

    def test_truncated_or_oversized_native_streams_are_rejected(self) -> None:
        stream = _ole10_native(b"a.txt", b"payload")
        for broken in (stream[:10], stream[:-3], struct.pack("<I", 10**6) + stream[4:]):
            with self.subTest(size=len(broken)), self.assertRaises(ValueError):
                parse_ole10_native(broken)

    def test_non_ole_data_is_not_unpacked(self) -> None:
        self.assertIsNone(unpack_ole(b"PK\x03\x04plain zip", prog_id="", stem="x"))

    def test_short_payload_is_never_opened_as_a_local_path(self) -> None:
        local = Path(__file__).resolve()
        self.assertIsNone(unpack_ole(bytes(local), prog_id="", stem="x"))
        with self.assertRaises(ValueError):
            unpack_ole(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1truncated", prog_id="", stem="x")

    def test_document_stream_extension_comes_from_its_bytes_then_program(self) -> None:
        self.assertEqual(payload_extension(_zip("word/document.xml")), ".docx")
        self.assertEqual(payload_extension(_zip("xl/workbook.xml")), ".xlsx")
        self.assertEqual(payload_extension(b"%PDF-1.7"), ".pdf")
        self.assertEqual(payload_extension(b"\xd0\xcf", "Excel.Sheet.8"), ".xls")
        self.assertEqual(payload_extension(b"??"), ".bin")

    def test_office_documents_keep_their_media_type_over_zip(self) -> None:
        self.assertEqual(
            media_type_for("a.xlsx", b"PK\x03\x04"),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.assertEqual(media_type_for("a.zip", b"PK\x03\x04"), "application/zip")


@unittest.skipUnless(GOLDEN.is_dir(), "repository fixtures are not available")
class FixtureAttachmentTests(unittest.TestCase):
    def _attachments(self, name: str) -> list[tuple[str, str | None, str]]:
        with pptx2markdown.parse(GOLDEN / name) as presentation:
            assets = {asset.id: asset for asset in presentation.document.assets}
            return [
                (block.label, block.program, assets[block.asset_id].media_type)
                for slide in presentation.document.slides
                for block in slide.blocks
                if isinstance(block, AttachmentBlock)
            ]

    def test_packager_object_is_unpacked_to_the_attached_file(self) -> None:
        attachments = self._attachments("shapecrawler_table-chart-ole.pptx")
        self.assertIn(("Result.json", "Package", "application/json"), attachments)

    def test_embedded_document_is_labelled_by_its_file_name(self) -> None:
        self.assertEqual(
            self._attachments("python-pptx_ole-object.pptx"),
            [
                (
                    "Microsoft_Excel_Worksheet.xlsx",
                    "Excel.Sheet.12",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            ],
        )


if __name__ == "__main__":
    unittest.main()
