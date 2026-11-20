from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from xml.sax.saxutils import escape

import pptx2markdown
from pptx2markdown.inline import emphasis, finalize, finalize_heading, markup
from pptx2markdown.model import ListBlock, TextBlock
from pptx2markdown.render import render_presentation_markdown

from .pptx_factory import build_test_pptx

BULLET = '<a:pPr marL="285750" indent="-285750"><a:buChar char="•"/></a:pPr>'


class FinalizeTests(unittest.TestCase):
    def test_block_starts_are_escaped(self) -> None:
        cases = {
            "# not a heading": r"\# not a heading",
            "> not a quote": r"\> not a quote",
            "- not a list": r"\- not a list",
            "* not a list": r"\* not a list",
            "2024. Revenue grew": r"2024\. Revenue grew",
            "3) third": r"3\) third",
            "---": r"\---",
            "===": r"\===",
            "```python": r"\`\`\`python",
        }
        for source, expected in cases.items():
            with self.subTest(source=source):
                self.assertEqual(finalize(source), expected)

    def test_inline_syntax_is_escaped_only_where_it_could_render(self) -> None:
        self.assertEqual(finalize("a <tag> & b &amp; c"), r"a \<tag> & b \&amp; c")
        self.assertEqual(finalize("see [x](y) and [참고]"), r"see \[x](y) and [참고]")
        self.assertEqual(finalize("[Page_9] marker"), r"\[Page_9] marker")
        self.assertEqual(finalize("*star* and 5 * 3"), r"\*star\* and 5 * 3")
        self.assertEqual(finalize("snake_case _under_"), r"snake_case \_under\_")
        self.assertEqual(finalize(r"C:\path `code`"), r"C:\\path \`code\`")
        self.assertEqual(finalize("~~strike~~ ~approx"), r"\~\~strike\~\~ ~approx")

    def test_dollars_are_escaped_only_when_they_could_pair_as_math(self) -> None:
        self.assertEqual(finalize("costs $5"), "costs $5")
        self.assertEqual(finalize("$5 to $10"), r"\$5 to \$10")

    def test_generated_markup_is_kept_verbatim(self) -> None:
        text = f"{markup('**')}bold *x*{markup('**')} {markup('$a_1$')}"
        self.assertEqual(finalize(text), r"**bold \*x\*** $a_1$")

    def test_emphasis_a_renderer_would_print_literally_is_dropped(self) -> None:
        bold = lambda text: emphasis(text, bold=True, italic=False)  # noqa: E731
        self.assertEqual(finalize(bold("중첩된(nested)") + "테이블"), "중첩된(nested)테이블")
        self.assertEqual(finalize("셀 " + bold("(merged)") + "과"), "셀 (merged)과")
        self.assertEqual(finalize("(" + bold("x") + ")"), "(**x**)")
        self.assertEqual(finalize(bold("품질 문제") + "로"), "**품질 문제**로")

    def test_each_line_is_checked_for_block_starts(self) -> None:
        self.assertEqual(finalize("intro\n- dash\n1. one"), "intro\n\\- dash\n1\\. one")
        self.assertEqual(finalize("- cell", line_start=False), "- cell")

    def test_heading_keeps_inner_numbers_and_escapes_a_closing_hash_run(self) -> None:
        self.assertEqual(finalize_heading("1. Intro to C#"), "1. Intro to C#")
        self.assertEqual(finalize_heading("Tags #"), r"Tags \#")


def _run(text: str, attributes: str = "", link: str = "") -> str:
    hyperlink = f'<a:hlinkClick r:id="{link}"/>' if link else ""
    return (
        f"<a:r><a:rPr {attributes}>{hyperlink}</a:rPr>"
        f'<a:t xml:space="preserve">{escape(text)}</a:t></a:r>'
    )


def _box(shape_id: int, y: int, paragraphs: str) -> str:
    return f"""
      <p:sp>
        <p:nvSpPr><p:cNvPr id="{shape_id}" name="Box {shape_id}"/><p:cNvSpPr txBox="1"/><p:nvPr/>
        </p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="{y}"/><a:ext cx="6000000" cy="900000"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/>{paragraphs}</p:txBody>
      </p:sp>"""


SHAPES = [
    # Literal Markdown syntax in plain text.
    "<a:p>"
    + _run("# not a heading <b>x</b>")
    + "</a:p>"
    + "<a:p>"
    + _run("- not a list")
    + "</a:p>",
    # Bold and italic runs; adjacent runs with one style are merged, spaces stay outside.
    "<a:p>"
    + _run("Plain ")
    + _run("bo", 'b="1"')
    + _run("ld ", 'b="1"')
    + _run("and ")
    + _run("both", 'b="1" i="1"')
    + _run(" end", 'i="1"')
    + "</a:p>",
    # A bold hyperlink whose label holds Markdown syntax, inside a bullet.
    f"<a:p>{BULLET}" + _run("see ") + _run("docs [v2]", 'b="1"', link="rId9") + "</a:p>",
]

SLIDE = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      {"".join(_box(index + 2, index * 1000000, body) for index, body in enumerate(SHAPES))}
    </p:spTree>
  </p:cSld>
</p:sld>
"""

LINK_RELATIONSHIP = (
    '<Relationship Id="rId9" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
    'Target="https://example.test/a (b)" TargetMode="External"/>'
)


class InlineMarkdownTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = build_test_pptx(
                Path(directory) / "inline.pptx",
                SLIDE,
                extra_slide_relationships=LINK_RELATIONSHIP,
            )
            with pptx2markdown.parse(source) as presentation:
                cls.document = presentation.document
        cls.blocks = {block.shape_id: block for block in cls.document.slides[0].blocks}

    def test_literal_syntax_renders_as_text(self) -> None:
        block = self.blocks["2"]
        self.assertIsInstance(block, TextBlock)
        self.assertEqual(block.text, "\\# not a heading \\<b>x\\</b>\n\\- not a list")

    def test_emphasis_runs_become_markdown_emphasis(self) -> None:
        self.assertEqual(self.blocks["3"].text, "Plain **bold** and ***both*** *end*")

    def test_link_label_is_escaped_and_destination_is_wrapped(self) -> None:
        block = self.blocks["4"]
        self.assertIsInstance(block, ListBlock)
        self.assertEqual(
            block.items[0].text, "see **[docs \\[v2\\]](<https://example.test/a (b)>)**"
        )

    def test_markdown_has_no_internal_sentinels(self) -> None:
        markdown = render_presentation_markdown(self.document, asset_mode="omit")
        self.assertNotIn("\ue000", markdown)
        self.assertNotIn("\ue001", markdown)
        self.assertIn("\\# not a heading", markdown)


if __name__ == "__main__":
    unittest.main()
