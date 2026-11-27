from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pptx2markdown
from pptx2markdown.model import ImageBlock, TextBlock

from .pptx_factory import build_test_pptx

SLIDE = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="4" name="AutoShape 1" descr="Team photo"/><p:cNvSpPr/><p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>
      </p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Our team</a:t></a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>
"""


class PictureFillTests(unittest.TestCase):
    def test_shape_filled_with_a_picture_shows_the_picture_and_its_text(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = build_test_pptx(Path(directory) / "fill.pptx", SLIDE)
            with pptx2markdown.parse(source) as presentation:
                blocks = presentation.document.slides[0].blocks
        self.assertEqual([type(block) for block in blocks], [ImageBlock, TextBlock])
        self.assertEqual(blocks[0].alt_text, "Team photo")
        self.assertEqual(blocks[1].text, "Our team")


if __name__ == "__main__":
    unittest.main()
