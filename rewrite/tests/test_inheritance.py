from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pptx2markdown
from pptx2markdown.model import BoundingBox, HeadingBlock, ListBlock, TextBlock

from .pptx_factory import build_templated_pptx

ROOT_NS = (
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)
GROUP = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'


def sp(shape_id: int, body: str, *, ph: str = "", xfrm: str = "") -> str:
    placeholder = f"<p:nvPr>{ph}</p:nvPr>" if ph else "<p:nvPr/>"
    geometry = f"<a:xfrm>{xfrm}</a:xfrm>" if xfrm else ""
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{shape_id}" name="Shape {shape_id}"/><p:cNvSpPr/>'
        f"{placeholder}</p:nvSpPr><p:spPr>{geometry}</p:spPr>"
        f"<p:txBody><a:bodyPr/><a:lstStyle/>{body}</p:txBody></p:sp>"
    )


def xfrm(x: int, y: int, cx: int, cy: int) -> str:
    return f'<a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/>'


def para(text: str) -> str:
    return f"<a:p><a:r><a:t>{text}</a:t></a:r></a:p>"


def part(root: str, shapes: str, extra: str = "", attrs: str = "") -> str:
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<p:{root} {ROOT_NS}{attrs}><p:cSld><p:spTree>{GROUP}{shapes}</p:spTree></p:cSld>"
        f"{extra}</p:{root}>"
    )


MASTER = part(
    "sldMaster",
    sp(2, "", ph='<p:ph type="title"/>', xfrm=xfrm(100, 200, 8_000_000, 900_000))
    + sp(3, "", ph='<p:ph type="body" idx="1"/>', xfrm=xfrm(100, 1_500_000, 8_000_000, 4_000_000))
    + sp(4, para("MASTER ONLY"), xfrm=xfrm(100, 6_000_000, 3_000_000, 300_000)),
    extra=(
        "<p:txStyles>"
        '<p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle>'
        '<p:bodyStyle><a:lvl1pPr><a:buChar char="•"/><a:defRPr sz="2000"/></a:lvl1pPr>'
        "</p:bodyStyle>"
        '<p:otherStyle><a:lvl1pPr><a:defRPr sz="2800"/></a:lvl1pPr></p:otherStyle>'
        "</p:txStyles>"
    ),
)

LAYOUT = part(
    "sldLayout",
    sp(2, "", ph='<p:ph type="title"/>')
    + sp(3, "", ph='<p:ph idx="1"/>')
    + sp(5, para("Click to edit Master title style"), xfrm=xfrm(0, 0, 100, 100))
    + sp(
        6,
        '<a:p><a:r><a:rPr><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>Docs</a:t></a:r></a:p>',
        xfrm=xfrm(100, 5_500_000, 3_000_000, 300_000),
    ),
)

LAYOUT_LINK = (
    '<Relationship Id="rIdLink"'
    ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"'
    ' Target="https://example.com/docs" TargetMode="External"/>'
)


def slide(attrs: str = "") -> str:
    return part(
        "sld",
        sp(10, para("Agenda"), xfrm=xfrm(100, 50, 3_000_000, 100))
        + sp(11, para("Inherited Title"), ph='<p:ph type="title"/>')
        + sp(12, para("First point") + para("Second point"), ph='<p:ph idx="1"/>')
        + sp(13, para("08/20/2024"), ph='<p:ph type="dt" idx="10"/>', xfrm=xfrm(0, 0, 10, 10)),
        attrs=attrs,
    )


class InheritanceTests(unittest.TestCase):
    def parse(self, slide_xml: str):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        source = build_templated_pptx(
            Path(temporary.name) / "templated.pptx",
            slide_xml=slide_xml,
            layout_xml=LAYOUT,
            master_xml=MASTER,
            layout_relationships=LAYOUT_LINK,
        )
        with pptx2markdown.parse(source) as presentation:
            return presentation.document.slides[0].blocks

    def test_title_placeholder_inherits_geometry_and_becomes_heading(self) -> None:
        blocks = {block.shape_id: block for block in self.parse(slide()) if block.shape_id}
        title = blocks["11"]
        self.assertIsInstance(title, HeadingBlock)
        self.assertEqual(title.level, 1)
        self.assertEqual(title.bbox, BoundingBox(x=100, y=200, width=8_000_000, height=900_000))

    def test_untyped_placeholder_inherits_body_type_and_bullets(self) -> None:
        blocks = {block.shape_id: block for block in self.parse(slide()) if block.shape_id}
        body = blocks["12"]
        self.assertIsInstance(body, ListBlock)
        self.assertEqual([item.text for item in body.items], ["First point", "Second point"])

    def test_inherited_font_size_makes_first_text_block_a_heading(self) -> None:
        blocks = {block.shape_id: block for block in self.parse(slide()) if block.shape_id}
        self.assertEqual(
            blocks["10"],
            HeadingBlock(
                id="s1.b1",
                text="Agenda",
                level=1,
                source="font",
                shape_id="10",
                bbox=BoundingBox(x=100, y=50, width=3_000_000, height=100),
            ),
        )

    def test_footer_placeholders_are_dropped(self) -> None:
        texts = [getattr(block, "text", "") for block in self.parse(slide())]
        self.assertNotIn("08/20/2024", texts)

    def test_layout_and_master_text_shapes_are_drawn_without_prompts(self) -> None:
        inherited = [
            block for block in self.parse(slide()) if block.source_part in {"layout", "master"}
        ]
        self.assertEqual(
            [(block.source_part, block.shape_id, block.text) for block in inherited],
            [
                ("layout", "6", "[Docs](https://example.com/docs)"),
                ("master", "4", "MASTER ONLY"),
            ],
        )
        self.assertTrue(all(isinstance(block, TextBlock) for block in inherited))

    def test_only_heading_on_a_slide_is_promoted_to_h1(self) -> None:
        numbered = sp(
            10,
            '<a:p><a:r><a:rPr sz="2800"/><a:t>1. Scope</a:t></a:r></a:p>',
            xfrm=xfrm(0, 3_000_000, 3_000_000, 300_000),
        )
        blocks = self.parse(part("sld", numbered, attrs=' showMasterSp="0"'))
        headings = [block for block in blocks if isinstance(block, HeadingBlock)]
        self.assertEqual([(h.text, h.level) for h in headings], [("1. Scope", 1)])

    def test_master_shapes_are_hidden_when_the_slide_opts_out(self) -> None:
        blocks = self.parse(slide(attrs=' showMasterSp="0"'))
        self.assertEqual({block.source_part for block in blocks}, {"slide", "layout"})


class MasterPlaceholderMatchTests(unittest.TestCase):
    def test_layout_content_placeholder_does_not_inherit_the_master_date_by_idx(self) -> None:
        # The layout's second content placeholder has idx 2, as has the master's date.
        master = part(
            "sldMaster",
            sp(2, "", ph='<p:ph type="title"/>', xfrm=xfrm(0, 0, 8_000_000, 900_000))
            + sp(
                3,
                "",
                ph='<p:ph type="body" idx="1"/>',
                xfrm=xfrm(0, 1_000_000, 8_000_000, 4_000_000),
            )
            + sp(
                4,
                "",
                ph='<p:ph type="dt" sz="half" idx="2"/>',
                xfrm=xfrm(0, 6_000_000, 2_000_000, 300_000),
            ),
        )
        layout = part(
            "sldLayout",
            sp(
                2,
                "",
                ph='<p:ph sz="half" idx="2"/>',
                xfrm=xfrm(0, 1_000_000, 4_000_000, 4_000_000),
            ),
        )
        slide = part("sld", sp(5, para("Agenda item"), ph='<p:ph sz="half" idx="2"/>'))
        with tempfile.TemporaryDirectory() as directory:
            source = build_templated_pptx(
                Path(directory) / "deck.pptx",
                slide_xml=slide,
                layout_xml=layout,
                master_xml=master,
            )
            with pptx2markdown.parse(source) as presentation:
                blocks = presentation.document.slides[0].blocks
        self.assertEqual([getattr(block, "text", None) for block in blocks], ["Agenda item"])


if __name__ == "__main__":
    unittest.main()
