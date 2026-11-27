from __future__ import annotations

import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path

import pptx2markdown
from pptx2markdown.model import MathBlock, TextBlock
from pptx2markdown.omml import omml_to_latex

from .pptx_factory import build_test_pptx

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

# Mirrors how PowerPoint saves equations: a:p > a14:m > m:oMath(Para), with the
# equation shape inside mc:AlternateContent and a picture-like fallback.
MATH_SLIDE = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>

      <mc:AlternateContent>
        <mc:Choice Requires="a14">
          <p:sp>
            <p:nvSpPr><p:cNvPr id="2" name="Display"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="800000"/></a:xfrm></p:spPr>
            <p:txBody><a:bodyPr/><a:lstStyle/>
              <a:p><a14:m><m:oMathPara><m:oMath>
                <m:sSup>
                  <m:e><m:r><m:t>x</m:t></m:r></m:e>
                  <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
                </m:sSup>
              </m:oMath></m:oMathPara></a14:m></a:p>
            </p:txBody>
          </p:sp>
        </mc:Choice>
        <mc:Fallback>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="2" name="Display"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="800000"/></a:xfrm></p:spPr>
            <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Fallback</a:t></a:r></a:p></p:txBody>
          </p:sp>
        </mc:Fallback>
      </mc:AlternateContent>

      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Inline"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="1000000"/><a:ext cx="4000000" cy="800000"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/>
          <a:p>
            <a:r><a:t>Area is </a:t></a:r>
            <a14:m><m:oMath><m:r><m:t>A</m:t></m:r></m:oMath></a14:m>
          </a:p>
        </p:txBody>
      </p:sp>

      <p:sp>
        <p:nvSpPr><p:cNvPr id="4" name="Matrix"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="2000000"/><a:ext cx="4000000" cy="800000"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/>
          <a:p><a14:m><m:oMath>
            <m:m>
              <m:mr><m:e><m:r><m:t>1</m:t></m:r></m:e><m:e><m:r><m:t>0</m:t></m:r></m:e></m:mr>
            </m:m>
          </m:oMath></a14:m></a:p>
        </p:txBody>
      </p:sp>

      <p:sp>
        <p:nvSpPr><p:cNvPr id="6" name="Split"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="4000000"/><a:ext cx="4000000" cy="800000"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/>
          <a:p>
            <a14:m><m:oMath><m:r><m:t>F=</m:t></m:r></m:oMath></a14:m>
            <a:r><a:t> </a:t></a:r>
            <a14:m><m:oMath><m:r><m:t>ma</m:t></m:r></m:oMath></a14:m>
          </a:p>
        </p:txBody>
      </p:sp>

      <p:sp>
        <p:nvSpPr><p:cNvPr id="5" name="Empty"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="3000000"/><a:ext cx="4000000" cy="800000"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/>
          <a:p><a14:m><m:oMath/></a14:m></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>
"""


class MathExtractionTests(unittest.TestCase):
    def setUp(self) -> None:
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        source = build_test_pptx(Path(temporary.name) / "math.pptx", MATH_SLIDE)
        with pptx2markdown.parse(source) as presentation:
            self.document = presentation.document
        self.blocks = {block.shape_id: block for block in self.document.slides[0].blocks}

    def test_equation_wrapped_in_a14_m_becomes_math_block(self) -> None:
        block = self.blocks["2"]
        self.assertIsInstance(block, MathBlock)
        self.assertEqual(block.latex, "{x}^{2}")

    def test_inline_equation_stays_inside_text(self) -> None:
        block = self.blocks["3"]
        self.assertIsInstance(block, TextBlock)
        self.assertEqual(block.text, "Area is $A$")

    def test_matrix_is_not_mistaken_for_the_a14_wrapper(self) -> None:
        block = self.blocks["4"]
        self.assertIsInstance(block, MathBlock)
        self.assertEqual(block.latex, r"\begin{matrix} 1 & 0 \end{matrix}")

    def test_equation_split_by_whitespace_runs_is_one_math_block(self) -> None:
        block = self.blocks["6"]
        self.assertIsInstance(block, MathBlock)
        self.assertEqual(block.latex, "F= ma")

    def test_unconvertible_equation_is_reported_not_dropped_silently(self) -> None:
        self.assertNotIn("5", self.blocks)
        warnings = [w for w in self.document.warnings if w.code == "math-unconverted"]
        self.assertEqual(len(warnings), 1)
        self.assertEqual((warnings[0].slide, warnings[0].shape_id), (1, "5"))


class OmmlToLatexTests(unittest.TestCase):
    def convert(self, body: str) -> str:
        return omml_to_latex(ET.fromstring(f'<m:oMath xmlns:m="{M_NS}">{body}</m:oMath>'))

    def test_math_italic_letters_and_greek_are_normalized(self) -> None:
        self.assertEqual(self.convert("<m:r><m:t>𝑥+π</m:t></m:r>"), r"x + \pi")

    def test_bar_less_fraction_is_not_rendered_as_a_fraction(self) -> None:
        latex = self.convert(
            '<m:d><m:e><m:f><m:fPr><m:type m:val="noBar"/></m:fPr>'
            "<m:num><m:r><m:t>n</m:t></m:r></m:num><m:den><m:r><m:t>k</m:t></m:r></m:den>"
            "</m:f></m:e></m:d>"
        )
        self.assertNotIn(r"\frac", latex)
        self.assertIn(r"\genfrac{}{}{0pt}{}{n}{k}", latex)

    def test_delimiter_characters_are_preserved(self) -> None:
        latex = self.convert(
            '<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr>'
            "<m:e><m:r><m:t>a</m:t></m:r></m:e></m:d>"
        )
        self.assertEqual(latex, r"\left[ a \right]")

    def test_empty_equation_converts_to_empty_string(self) -> None:
        self.assertEqual(self.convert(""), "")


if __name__ == "__main__":
    unittest.main()
