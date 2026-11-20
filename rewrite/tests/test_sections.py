from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path

import pptx2markdown
from pptx2markdown import Recipe, slide_chunks
from pptx2markdown.agent import document_view
from pptx2markdown.recipe import Selector

from .pptx_factory import PRESENTATION_RELS, build_test_pptx

PRESENTATION = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
  <p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">
    <p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
      <p14:section name="Results" id="{A}"><p14:sldIdLst><p14:sldId id="256"/></p14:sldIdLst>
      </p14:section>
      <p14:section name="Appendix" id="{B}"><p14:sldIdLst><p14:sldId id="257"/></p14:sldIdLst>
      </p14:section>
    </p14:sectionLst>
  </p:ext></p:extLst>
</p:presentation>"""


def _two_sections(path: Path) -> Path:
    """The factory deck with its slide twice, in sections "Results" and "Appendix"."""

    single = build_test_pptx(path.with_name("single.pptx"))
    relationships = PRESENTATION_RELS.replace(
        "</Relationships>",
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/'
        '2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>',
    )
    with zipfile.ZipFile(single) as source, zipfile.ZipFile(path, "w") as target:
        for info in source.infolist():
            data = source.read(info)
            if info.filename == "ppt/presentation.xml":
                data = PRESENTATION.encode()
            elif info.filename == "ppt/_rels/presentation.xml.rels":
                data = relationships.encode()
            target.writestr(info, data)
            if info.filename.startswith("ppt/slides/") and "slide1.xml" in info.filename:
                target.writestr(info.filename.replace("slide1.xml", "slide2.xml"), data)
    return path


class SectionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.directory = tempfile.TemporaryDirectory()
        cls.source = _two_sections(Path(cls.directory.name) / "deck.pptx")

    @classmethod
    def tearDownClass(cls) -> None:
        cls.directory.cleanup()

    def test_each_slide_knows_its_section(self) -> None:
        with pptx2markdown.parse(self.source) as presentation:
            document = presentation.document
        self.assertEqual([s.section for s in document.slides], ["Results", "Appendix"])
        self.assertEqual(
            [chunk.metadata["section"] for chunk in slide_chunks(document)],
            ["Results", "Appendix"],
        )
        self.assertIn("<!-- section: Appendix -->", document_view(document))

    def test_a_recipe_can_select_a_section(self) -> None:
        recipe = Recipe.model_validate(
            {"rules": [{"target": "section[appendix]", "exclude": True}]}
        )
        with pptx2markdown.parse(self.source, recipe=recipe) as presentation:
            self.assertEqual([s.page for s in presentation.document.slides], [1])
        tables = Selector.parse("section[Results]/table")
        self.assertEqual((tables.section, tables.kind), ("results", "table"))

    def test_plain_markdown_is_unchanged(self) -> None:
        with pptx2markdown.parse(self.source) as presentation:
            self.assertNotIn("section", presentation.render_markdown(assets="omit"))


if __name__ == "__main__":
    unittest.main()
