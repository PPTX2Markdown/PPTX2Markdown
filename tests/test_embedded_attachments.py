from __future__ import annotations

import io
import struct
import unittest
import zipfile

from pptx2markdown.main_converter.embedded_attachments import (
    infer_payload_extension,
    parse_ole10_native,
)


class EmbeddedAttachmentTests(unittest.TestCase):
    def test_packager_native_stream_preserves_filename_and_payload(self) -> None:
        filename = b"report.pdf\x00"
        source_path = b"C:\\source\\report.pdf\x00"
        temp_path = b"C:\\temp\\report.pdf\x00"
        payload = b"%PDF-1.7\nexample"
        body = (
            b"\x02\x00"
            + filename
            + source_path
            + b"\x00\x00\x03\x00"
            + struct.pack("<I", len(temp_path))
            + temp_path
            + struct.pack("<I", len(payload))
            + payload
        )
        stream = struct.pack("<I", len(body)) + body

        extracted_name, extracted_payload = parse_ole10_native(stream)

        self.assertEqual(extracted_name, "report.pdf")
        self.assertEqual(extracted_payload, payload)

    def test_unicode_label_replaces_the_code_page_name(self) -> None:
        def utf16(text: str) -> bytes:
            return struct.pack("<I", len(text)) + text.encode("utf-16-le")

        temp_path = b"C:\\t\x00"
        body = (
            b"\x02\x00"
            + "보고서.hwp".encode("cp949")
            + b"\x00C:\\x\x00\x00\x00\x03\x00"
            + struct.pack("<I", len(temp_path))
            + temp_path
            + struct.pack("<I", 3)
            + b"HWP"
            + utf16("C:\\t\\보고서.hwp")
            + utf16("보고서.hwp")
        )
        name, payload = parse_ole10_native(struct.pack("<I", len(body)) + body)
        self.assertEqual((name, payload), ("보고서.hwp", b"HWP"))

    def test_ooxml_package_payload_extension_is_inferred_from_zip_parts(self) -> None:
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            archive.writestr("[Content_Types].xml", "<Types/>")
            archive.writestr("xl/workbook.xml", "<workbook/>")

        self.assertEqual(infer_payload_extension(buffer.getvalue()), ".xlsx")
        self.assertEqual(infer_payload_extension(b"%PDF-1.4\n"), ".pdf")


if __name__ == "__main__":
    unittest.main()
