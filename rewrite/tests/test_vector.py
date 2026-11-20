from __future__ import annotations

import os
import stat
import struct
import sys
import tempfile
import unittest
import zlib
from pathlib import Path
from unittest import mock

import pptx2markdown
from pptx2markdown.assets import AssetStore, media_type_for
from pptx2markdown.libreoffice import find_soffice
from pptx2markdown.model import ImageBlock

from .pptx_factory import build_test_pptx


def rectangle_wmf() -> bytes:
    """A placeable WMF drawing a red square in the middle of a 1000 x 1000 window."""

    def record(function: int, *params: int) -> bytes:
        return struct.pack("<IH", 3 + len(params), function) + struct.pack(
            f"<{len(params)}h", *params
        )

    records = b"".join(
        [
            record(0x020B, 0, 0),  # SETWINDOWORG
            record(0x020C, 1000, 1000),  # SETWINDOWEXT
            record(0x02FC, 0, 0x00FF, 0, 0),  # CREATEBRUSHINDIRECT, solid red
            record(0x012D, 0),  # SELECTOBJECT
            record(0x041B, 800, 800, 200, 200),  # RECTANGLE
            record(0x0000),  # EOF
        ]
    )
    header = struct.pack("<HHHIHIH", 1, 9, 0x0300, (18 + len(records)) // 2, 1, 7, 0)
    placeable = struct.pack("<IHhhhhHI", 0x9AC6CDD7, 0, 0, 0, 1000, 1000, 1440, 0)
    checksum = 0
    for (word,) in struct.iter_unpack("<H", placeable):
        checksum ^= word
    return placeable + struct.pack("<H", checksum) + header + records


def tiny_png() -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return (
            struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))
        )

    header = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", header)
        + chunk(b"IDAT", zlib.compress(b"\x00\x00\x00\x00"))
        + chunk(b"IEND", b"")
    )


SLIDE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:pic><p:nvPicPr><p:cNvPr id="4" name="Diagram" descr="Old diagram"/><p:cNvPicPr/>
      <p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdWmf"/></p:blipFill>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="3000000"/></a:xfrm></p:spPr>
    </p:pic>
  </p:spTree></p:cSld>
</p:sld>"""
RELATIONSHIP = (
    '<Relationship Id="rIdWmf" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/image" Target="../media/image2.wmf"/>'
)

# Stands in for soffice: writes $FAKE_PNG as <outdir>/<stem>.png for each input, or fails.
FAKE_SOFFICE = """#!/bin/sh
out=""; inputs=""
while [ $# -gt 0 ]; do
  case "$1" in --outdir) out="$2"; shift ;; *.wmf|*.emf) inputs="$inputs $1" ;; esac
  shift
done
if [ -n "$FAKE_FAIL" ]; then exit 1; fi
mkdir -p "$out"
for f in $inputs; do cp "$FAKE_PNG" "$out/$(basename "$f" | sed 's/\\.[a-z]*$//').png"; done
"""


class MediaTypeTests(unittest.TestCase):
    def test_metafiles_are_recognised_by_their_bytes_then_extension(self) -> None:
        emf = b"\x01\x00\x00\x00" + b"\x00" * 36 + b" EMF" + b"\x00" * 20
        self.assertEqual(media_type_for("x.bin", emf), "image/emf")
        self.assertEqual(media_type_for("x.bin", rectangle_wmf()[:64]), "image/wmf")
        self.assertEqual(media_type_for("x.bin", b"\x01\x00\x09\x00\x00\x03"), "image/wmf")
        self.assertEqual(media_type_for("picture.EMF"), "image/emf")


class ReplaceTests(unittest.TestCase):
    def test_replacing_keeps_the_id_and_leaves_shared_payloads_alone(self) -> None:
        with AssetStore() as store:
            first = store.register_bytes(
                b"same", logical_key="a", kind="image", original_name="a.wmf", source_part="a"
            )
            second = store.register_bytes(
                b"same", logical_key="b", kind="image", original_name="b.wmf", source_part="b"
            )
            replaced = store.replace(
                first.id, b"png!", original_name="a.png", media_type="image/png"
            )
            self.assertEqual(replaced.id, first.id)
            self.assertEqual((replaced.original_name, replaced.size), ("a.png", 4))
            self.assertEqual(store.read_bytes(first.id), b"png!")
            self.assertEqual(store.read_bytes(second.id), b"same")
            self.assertEqual(store.descriptor(first.id), replaced)


class VectorImageTests(unittest.TestCase):
    def setUp(self) -> None:
        self.directory = tempfile.TemporaryDirectory()
        self.root = Path(self.directory.name)
        self.source = build_test_pptx(
            self.root / "deck.pptx",
            SLIDE,
            extra_slide_relationships=RELATIONSHIP,
            extra_parts={"ppt/media/image2.wmf": rectangle_wmf()},
        )

    def tearDown(self) -> None:
        self.directory.cleanup()

    def _image(self, **options: object) -> tuple[ImageBlock, object, list[str]]:
        with pptx2markdown.parse(self.source, **options) as presentation:
            (block,) = presentation.document.slides[0].blocks
            assert isinstance(block, ImageBlock)
            codes = [warning.code for warning in presentation.document.warnings]
            return block, presentation.assets.descriptor(block.asset_id), codes

    def test_metafiles_are_kept_unless_asked_for(self) -> None:
        _, asset, _ = self._image()
        self.assertEqual((asset.media_type, asset.original_name), ("image/wmf", "image2.wmf"))

    @unittest.skipIf(sys.platform.startswith("win"), "the stand-in converter is a shell script")
    def test_converted_png_takes_the_place_of_the_metafile(self) -> None:
        script = self.root / "soffice"
        script.write_text(FAKE_SOFFICE)
        script.chmod(script.stat().st_mode | stat.S_IEXEC)
        png = self.root / "fake.png"
        png.write_bytes(tiny_png())
        environment = {"SOFFICE_PATH": str(script), "FAKE_PNG": str(png)}
        with mock.patch.dict(os.environ, environment):
            block, asset, codes = self._image(vector_images="png")
            self.assertEqual((asset.media_type, asset.original_name), ("image/png", "image2.png"))
            self.assertEqual(block.alt_text, "Old diagram")
            self.assertNotIn("vector-image-not-converted", codes)
            with mock.patch.dict(os.environ, {"FAKE_FAIL": "1"}):
                _, asset, codes = self._image(vector_images="png")
            self.assertEqual(asset.media_type, "image/wmf")
            self.assertIn("vector-image-not-converted", codes)

    @unittest.skipUnless(find_soffice(), "LibreOffice is not installed")
    def test_real_libreoffice_draws_the_metafile(self) -> None:
        with pptx2markdown.parse(self.source, vector_images="png") as presentation:
            (block,) = presentation.document.slides[0].blocks
            data = presentation.assets.read_bytes(block.asset_id)
        self.assertTrue(data.startswith(b"\x89PNG"))
        try:
            from PIL import Image
        except ImportError:
            return
        import io

        with Image.open(io.BytesIO(data)) as image:
            # Trimmed to the square, not the page LibreOffice draws it on.
            self.assertLess(image.width, 200)
            self.assertEqual(
                image.convert("RGB").getpixel((image.width // 2, image.height // 2)), (255, 0, 0)
            )


if __name__ == "__main__":
    unittest.main()
