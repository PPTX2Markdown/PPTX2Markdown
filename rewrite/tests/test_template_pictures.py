from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pptx2markdown
from pptx2markdown.model import ImageBlock
from pptx2markdown.recipe import Recipe

from .pptx_factory import build_templated_pptx
from .test_inheritance import para, part, sp, xfrm

LOGO = """<p:pic><p:nvPicPr><p:cNvPr id="7" name="Logo" descr="Company logo"/><p:cNvPicPr/>
<p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdLogo"/></p:blipFill>
<p:spPr><a:xfrm><a:off x="8000000" y="100000"/><a:ext cx="900000" cy="400000"/></a:xfrm>
</p:spPr></p:pic>"""
LOGO_RELATIONSHIP = (
    '<Relationship Id="rIdLogo" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
    'Target="../media/logo.png"/>'
)


class TemplatePictureTests(unittest.TestCase):
    def setUp(self) -> None:
        self.directory = tempfile.TemporaryDirectory()
        self.source = build_templated_pptx(
            Path(self.directory.name) / "deck.pptx",
            slide_xml=part("sld", sp(2, para("Body"), xfrm=xfrm(0, 1000000, 4000000, 500000))),
            layout_xml=part("sldLayout", LOGO),
            master_xml=part("sldMaster", ""),
            layout_relationships=LOGO_RELATIONSHIP,
            extra_parts={"ppt/media/logo.png": b"\x89PNG\r\n\x1a\nlogo"},
        )

    def tearDown(self) -> None:
        self.directory.cleanup()

    def test_layout_picture_is_drawn_with_its_layout_relationship(self) -> None:
        with pptx2markdown.parse(self.source) as presentation:
            images = [
                b for b in presentation.document.slides[0].blocks if isinstance(b, ImageBlock)
            ]
            self.assertEqual(len(images), 1)
            self.assertEqual(images[0].source_part, "layout")
            self.assertEqual(images[0].alt_text, "Company logo")
            asset = presentation.assets.descriptor(images[0].asset_id)
            self.assertEqual(asset.source_part, "ppt/media/logo.png")

    def test_thin_band_on_the_layout_is_decoration(self) -> None:
        band = LOGO.replace('id="7" name="Logo"', 'id="8" name="Band"').replace(
            'cx="900000" cy="400000"', 'cx="9906000" cy="72008"'
        )
        source = build_templated_pptx(
            Path(self.directory.name) / "band.pptx",
            slide_xml=part("sld", sp(2, para("Body"), xfrm=xfrm(0, 1000000, 4000000, 500000))),
            layout_xml=part("sldLayout", band),
            master_xml=part("sldMaster", ""),
            layout_relationships=LOGO_RELATIONSHIP,
            extra_parts={"ppt/media/logo.png": b"\x89PNG\r\n\x1a\nlogo"},
        )
        with pptx2markdown.parse(source) as presentation:
            kinds = [block.kind for block in presentation.document.slides[0].blocks]
        self.assertEqual(kinds, ["text"])

    def test_recipe_can_leave_template_pictures_out(self) -> None:
        with pptx2markdown.parse(
            self.source, recipe=Recipe(template_pictures=False)
        ) as presentation:
            kinds = [block.kind for block in presentation.document.slides[0].blocks]
        self.assertNotIn("image", kinds)
        self.assertIn("text", kinds)


if __name__ == "__main__":
    unittest.main()
