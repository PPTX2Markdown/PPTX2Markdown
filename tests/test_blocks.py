from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pptx2markdown
from pptx2markdown.model import ListBlock, MathBlock, TextBlock

from .pptx_factory import build_test_pptx

BULLET = '<a:pPr marL="285750" indent="-285750"><a:buChar char="•"/></a:pPr>'


def _paragraph(text: str, properties: str = "") -> str:
    return f'<a:p>{properties}<a:r><a:t xml:space="preserve">{text}</a:t></a:r></a:p>'


def _text_box(shape_id: int, y: int, paragraphs: str) -> str:
    return f"""
      <p:sp>
        <p:nvSpPr><p:cNvPr id="{shape_id}" name="Box {shape_id}"/><p:cNvSpPr txBox="1"/><p:nvPr/>
        </p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="{y}"/><a:ext cx="6000000" cy="900000"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/>{paragraphs}</p:txBody>
      </p:sp>"""


EQUATION = (
    "<a:p><a14:m><m:oMathPara><m:oMath><m:r><m:t>E=mc</m:t></m:r></m:oMath></m:oMathPara>"
    "</a14:m></a:p>"
)

SHAPES = [
    # A plain text box: one block, one line per paragraph.
    _paragraph("Alpha") + _paragraph("Beta") + _paragraph("Gamma"),
    # Text and bullets stay typed blocks, in order.
    _paragraph("Lead")
    + _paragraph("One", BULLET)
    + _paragraph("Two", BULLET)
    + _paragraph("Tail"),
    # Typed triangle bullets are list items; a lone triangle is not.
    _paragraph("▶ First") + _paragraph("▶▶Second") + _paragraph("▶"),
    # A hand-indented paragraph after a bullet continues it; an unindented one does not.
    _paragraph("Wrapped bullet starts", BULLET)
    + _paragraph("      and ends here.")
    + _paragraph("Separate text"),
    # An equation paragraph splits the surrounding text.
    _paragraph("Before") + EQUATION + _paragraph("After"),
    # Indented text with no item before it is ordinary text.
    _paragraph("   Indented only"),
    # Characters stacked one per paragraph form one word: two full-width ones,
    # four or more narrow letters; three narrow letters stay lines.
    _paragraph("구") + _paragraph("분"),
    "".join(_paragraph(letter) for letter in "LIKE"),
    _paragraph("a") + _paragraph("b") + _paragraph("c"),
    # The same by line breaks within one paragraph.
    "<a:p>" + "<a:br/>".join(f"<a:r><a:t>{letter}</a:t></a:r>" for letter in "AXIS") + "</a:p>",
]

SLIDE = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      {"".join(_text_box(index + 2, index * 1000000, body) for index, body in enumerate(SHAPES))}
    </p:spTree>
  </p:cSld>
</p:sld>
"""


def _summary(block) -> tuple:
    if isinstance(block, TextBlock):
        return ("text", block.text)
    if isinstance(block, ListBlock):
        return ("list", tuple((item.text, item.level) for item in block.items))
    if isinstance(block, MathBlock):
        return ("math", block.latex)
    return (block.kind,)


class ShapeBlockTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = build_test_pptx(Path(directory) / "blocks.pptx", SLIDE)
            with pptx2markdown.parse(source) as presentation:
                cls.document = presentation.document
        cls.shapes: dict[str, list[tuple]] = {}
        for block in cls.document.slides[0].blocks:
            cls.shapes.setdefault(block.shape_id, []).append(_summary(block))

    def test_stacked_characters_form_one_word(self) -> None:
        self.assertEqual(
            [self.shapes[str(shape_id)] for shape_id in range(8, 12)],
            [[("text", "구분")], [("text", "LIKE")], [("text", "a\nb\nc")], [("text", "AXIS")]],
        )

    def test_plain_text_box_is_one_block(self) -> None:
        self.assertEqual(self.shapes["2"], [("text", "Alpha\nBeta\nGamma")])

    def test_text_and_bullets_stay_separate_typed_blocks(self) -> None:
        self.assertEqual(
            self.shapes["3"],
            [("text", "Lead"), ("list", (("One", 0), ("Two", 0))), ("text", "Tail")],
        )

    def test_typed_triangle_bullets_become_list_items(self) -> None:
        self.assertEqual(
            self.shapes["4"],
            [("list", (("First", 0), ("Second", 0))), ("text", "▶")],
        )

    def test_indented_paragraph_continues_the_previous_item(self) -> None:
        self.assertEqual(
            self.shapes["5"],
            [
                ("list", (("Wrapped bullet starts\nand ends here.", 0),)),
                ("text", "Separate text"),
            ],
        )

    def test_equation_paragraph_splits_text(self) -> None:
        self.assertEqual(
            self.shapes["6"],
            [("text", "Before"), ("math", "E=mc"), ("text", "After")],
        )

    def test_indentation_without_a_preceding_item_is_plain_text(self) -> None:
        self.assertEqual(self.shapes["7"], [("text", "Indented only")])

    def test_markdown_keeps_one_line_per_paragraph(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = build_test_pptx(Path(directory) / "blocks.pptx", SLIDE)
            with pptx2markdown.parse(source) as presentation:
                markdown = presentation.render_markdown(assets="omit")
        self.assertIn("Alpha\nBeta\nGamma", markdown)
        self.assertIn("- Wrapped bullet starts\nand ends here.", markdown)


if __name__ == "__main__":
    unittest.main()
