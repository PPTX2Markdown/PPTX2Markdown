from __future__ import annotations

import io
import json
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path

import pptx2markdown
from pptx2markdown import cli, mcp_server
from pptx2markdown.audit import normalize

from .pptx_factory import SLIDE, build_test_pptx

FOOTER = """<p:sp><p:nvSpPr><p:cNvPr id="9" name="Footer"/><p:cNvSpPr/>
<p:nvPr><p:ph type="ftr" idx="11"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="0" y="6000000"/><a:ext cx="3000000" cy="300000"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Confidential footer</a:t></a:r></a:p>
</p:txBody></p:sp>"""


class AuditTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.directory = tempfile.TemporaryDirectory()
        cls.source = build_test_pptx(Path(cls.directory.name) / "deck.pptx")

    @classmethod
    def tearDownClass(cls) -> None:
        cls.directory.cleanup()

    def test_converted_deck_covers_its_text_links_included(self) -> None:
        with pptx2markdown.parse(self.source) as presentation:
            report = presentation.audit()
        self.assertTrue(report.ok, report.losses)
        self.assertEqual(report.coverage, 1.0)
        self.assertEqual((report.slides_in_package, report.slides_audited), (1, 1))
        self.assertGreater(report.chunks, 4)

    def test_text_missing_from_the_document_is_reported_with_its_shape(self) -> None:
        with pptx2markdown.parse(self.source) as presentation:
            document = presentation.document
        slide = document.slides[0]
        kept = tuple(block for block in slide.blocks if block.shape_id != "3")
        trimmed = document.model_copy(
            update={"slides": (slide.model_copy(update={"blocks": kept}),)}
        )
        report = pptx2markdown.audit(self.source, trimmed)
        self.assertFalse(report.ok)
        self.assertEqual(
            [(loss.slide, loss.shape_id, loss.text) for loss in report.losses],
            [(1, "3", "OpenAI reference"), (1, "3", "Revenue grew"), (1, "3", "APAC led")],
        )
        self.assertLess(report.coverage, 1.0)

    def test_a_slide_missing_from_the_document_is_reported(self) -> None:
        with pptx2markdown.parse(self.source) as presentation:
            empty = presentation.document.model_copy(update={"slides": ()})
        report = pptx2markdown.audit(self.source.read_bytes(), empty)
        self.assertEqual(report.missing_slides, (1,))
        self.assertEqual(report.slides_audited, 0)

    def test_footers_and_cells_hidden_by_a_merge_are_not_audited(self) -> None:
        slide = SLIDE.replace("<p:grpSpPr/>", "<p:grpSpPr/>" + FOOTER, 1).replace(
            "<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Revenue</a:t>",
            '<a:tc hMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Hidden</a:t>',
            1,
        )
        source = build_test_pptx(Path(self.directory.name) / "footer.pptx", slide)
        with pptx2markdown.parse(source) as presentation:
            report = presentation.audit()
        self.assertTrue(report.ok, report.losses)

    def test_comparison_ignores_escaping_spacing_and_entities(self) -> None:
        self.assertEqual(normalize(r"**A\_b** &amp; <br>C-d 3"), "abbrcd3")
        self.assertEqual(normalize("구\n분"), "구분")

    def test_cli_prints_the_report_and_writes_nothing(self) -> None:
        output = io.StringIO()
        out_dir = Path(self.directory.name) / "out"
        with redirect_stdout(output):
            code = cli.main([str(self.source), "--audit", "-o", str(out_dir)])
        self.assertEqual(code, 0)
        self.assertFalse(out_dir.exists())
        report = json.loads(output.getvalue())
        self.assertEqual((report[0]["source"], report[0]["ok"]), ("deck.pptx", True))

    def test_mcp_tool_and_bytes_input(self) -> None:
        self.assertTrue(mcp_server.audit_presentation(str(self.source))["ok"])
        with pptx2markdown.parse(self.source.read_bytes()) as presentation:
            with self.assertRaisesRegex(ValueError, "pass its bytes"):
                presentation.audit()


if __name__ == "__main__":
    unittest.main()
