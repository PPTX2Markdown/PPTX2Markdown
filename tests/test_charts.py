from __future__ import annotations

import io
import tempfile
import unittest
import zipfile
from pathlib import Path

import pptx2markdown
from pptx2markdown.errors import PackageSecurityError
from pptx2markdown.model import ChartBlock, UnsupportedBlock
from pptx2markdown.ooxml import PackageLimits, validate_embedded_package

from .pptx_factory import build_test_pptx

_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_CHARTEX_REL = "http://schemas.microsoft.com/office/2014/relationships/chartEx"


def _frame(shape_id: int, y: int, graphic_data: str) -> str:
    return f"""
      <p:graphicFrame>
        <p:nvGraphicFramePr>
          <p:cNvPr id="{shape_id}" name="Chart {shape_id}"/><p:cNvGraphicFramePr/><p:nvPr/>
        </p:nvGraphicFramePr>
        <p:xfrm><a:off x="0" y="{y}"/><a:ext cx="4000000" cy="2000000"/></p:xfrm>
        <a:graphic>{graphic_data}</a:graphic>
      </p:graphicFrame>"""


# A classic chart, and a chartEx chart the way PowerPoint saves it: inside
# mc:AlternateContent with a picture fallback.
CHART_SLIDE = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="{_REL}"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      {
    _frame(
        2,
        0,
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
        '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
        ' r:id="rIdChart"/></a:graphicData>',
    )
}
      <mc:AlternateContent>
        <mc:Choice xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex"
          Requires="cx4">
          {
    _frame(
        3,
        3000000,
        '<a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex">'
        '<cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"'
        ' r:id="rIdChartEx"/></a:graphicData>',
    )
}
        </mc:Choice>
        <mc:Fallback>
          <p:pic>
            <p:nvPicPr><p:cNvPr id="3" name="Chart 3"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
            <p:blipFill><a:blip r:embed="rIdImage"/></p:blipFill>
            <p:spPr><a:xfrm><a:off x="0" y="3000000"/><a:ext cx="4000000" cy="2000000"/></a:xfrm>
            </p:spPr>
          </p:pic>
        </mc:Fallback>
      </mc:AlternateContent>
      {
    _frame(
        4,
        6000000,
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
        '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
        ' r:id="rIdMissing"/></a:graphicData>',
    )
}
    </p:spTree>
  </p:cSld>
</p:sld>
"""

CHART_RELS = (
    f'<Relationship Id="rIdChart" Type="{_REL}/chart" Target="../charts/chart1.xml"/>'
    f'<Relationship Id="rIdChartEx" Type="{_CHARTEX_REL}" Target="../charts/chartEx1.xml"/>'
)

CHART_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:plotArea><c:barChart><c:barDir val="col"/>
    <c:ser><c:idx val="0"/><c:order val="0"/>
      <c:tx><c:strRef><c:strCache><c:ptCount val="1"/>
        <c:pt idx="0"><c:v>Sales</c:v></c:pt></c:strCache></c:strRef></c:tx>
      <c:cat><c:strRef><c:strCache><c:ptCount val="2"/>
        <c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt>
      </c:strCache></c:strRef></c:cat>
      <c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/>
        <c:pt idx="0"><c:v>4.4000000000000004</c:v></c:pt><c:pt idx="1"><c:v>2.5</c:v></c:pt>
      </c:numCache></c:numRef></c:val>
    </c:ser>
  </c:barChart></c:plotArea></c:chart>
</c:chartSpace>
"""

CHARTEX_XML = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
  xmlns:r="{_REL}">
  <cx:chartData>
    <cx:externalData r:id="rIdBook" cx:autoUpdate="0"/>
    <cx:data id="0">
      <cx:strDim type="cat"><cx:lvl ptCount="2">
        <cx:pt idx="0">Cached A</cx:pt><cx:pt idx="1">Cached B</cx:pt></cx:lvl></cx:strDim>
      <cx:numDim type="val"><cx:lvl ptCount="2">
        <cx:pt idx="0">10</cx:pt><cx:pt idx="1">20</cx:pt></cx:lvl></cx:numDim>
    </cx:data>
  </cx:chartData>
  <cx:chart><cx:plotArea><cx:plotAreaRegion>
    <cx:series layoutId="funnel"><cx:tx><cx:txData><cx:v>Pipeline</cx:v></cx:txData></cx:tx>
      <cx:dataId val="0"/></cx:series>
  </cx:plotAreaRegion></cx:plotArea></cx:chart>
</cx:chartSpace>
"""

CHARTEX_RELS = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdBook" Type="{_REL}/package"
    Target="../embeddings/Microsoft_Excel_Worksheet1.xlsx"/>
</Relationships>
"""

_SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _workbook(rows: list[tuple[str, int]], *, padding: int = 0, doctype: bool = False) -> bytes:
    cells = "".join(
        f'<row r="{index}"><c r="A{index}" t="inlineStr"><is><t>{label}</t></is></c>'
        f'<c r="B{index}"><v>{value}</v></c></row>'
        for index, (label, value) in enumerate(rows, 1)
    )
    prolog = '<!DOCTYPE worksheet [<!ENTITY x "y">]>' if doctype else ""
    sheet = (
        f'<?xml version="1.0"?>{prolog}<worksheet xmlns="{_SHEET_NS}">'
        f"<sheetData>{cells}</sheetData>{' ' * padding}</worksheet>"
    )
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as book:
        book.writestr(
            "xl/workbook.xml",
            f'<workbook xmlns="{_SHEET_NS}" xmlns:r="{_REL}"><sheets>'
            '<sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
        )
        book.writestr(
            "xl/_rels/workbook.xml.rels",
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{_REL}/worksheet" Target="worksheets/sheet1.xml"/>'
            "</Relationships>",
        )
        book.writestr("xl/worksheets/sheet1.xml", sheet)
    return buffer.getvalue()


WORKBOOK_ROWS = [("Stage", 0), ("Workbook A", 30), ("Workbook B", 40)]


def _parse(workbook: bytes):
    temporary = tempfile.TemporaryDirectory()
    source = build_test_pptx(
        Path(temporary.name) / "charts.pptx",
        CHART_SLIDE,
        extra_slide_relationships=CHART_RELS,
        extra_parts={
            "ppt/charts/chart1.xml": CHART_XML,
            "ppt/charts/chartEx1.xml": CHARTEX_XML,
            "ppt/charts/_rels/chartEx1.xml.rels": CHARTEX_RELS,
            "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx": workbook,
        },
    )
    with pptx2markdown.parse(source) as presentation:
        document = presentation.document
    temporary.cleanup()
    return document, {block.shape_id: block for block in document.slides[0].blocks}


class ChartTests(unittest.TestCase):
    def test_classic_chart_has_type_line_and_rounded_values(self) -> None:
        _, blocks = _parse(_workbook(WORKBOOK_ROWS))
        block = blocks["2"]
        self.assertIsInstance(block, ChartBlock)
        self.assertTrue(block.markdown.startswith("**Chart type:** Column Chart"))
        self.assertIn("| Q1 | 4.4   |", block.markdown)
        self.assertNotIn("4.4000000000000004", block.markdown)

    def test_chartex_inside_alternate_content_prefers_embedded_workbook(self) -> None:
        document, blocks = _parse(_workbook(WORKBOOK_ROWS))
        block = blocks["3"]
        self.assertIsInstance(block, ChartBlock)
        self.assertIn("**Chart type:** Funnel", block.markdown)
        self.assertIn("Workbook A", block.markdown)
        self.assertNotIn("Cached A", block.markdown)
        self.assertFalse([w for w in document.warnings if w.code == "chart-workbook-rejected"])

    def test_missing_chart_part_is_reported_as_unsupported(self) -> None:
        document, blocks = _parse(_workbook(WORKBOOK_ROWS))
        self.assertIsInstance(blocks["4"], UnsupportedBlock)
        codes = {(w.code, w.shape_id) for w in document.warnings}
        self.assertIn(("chart-data-unavailable", "4"), codes)

    def test_compression_bomb_workbook_falls_back_to_cached_values(self) -> None:
        document, blocks = _parse(_workbook(WORKBOOK_ROWS, padding=4 * 1024 * 1024))
        block = blocks["3"]
        self.assertIsInstance(block, ChartBlock)
        self.assertIn("Cached A", block.markdown)
        self.assertNotIn("Workbook A", block.markdown)
        rejected = [w for w in document.warnings if w.code == "chart-workbook-rejected"]
        self.assertEqual([(w.slide, w.shape_id) for w in rejected], [(1, "3")])
        self.assertIn("compression ratio", rejected[0].message)

    def test_workbook_with_dtd_falls_back_to_cached_values(self) -> None:
        document, blocks = _parse(_workbook(WORKBOOK_ROWS, doctype=True))
        self.assertIn("Cached A", blocks["3"].markdown)
        rejected = [w for w in document.warnings if w.code == "chart-workbook-rejected"]
        self.assertIn("DTD", rejected[0].message)


class EmbeddedPackageValidationTests(unittest.TestCase):
    def test_valid_workbook_passes(self) -> None:
        validate_embedded_package(_workbook(WORKBOOK_ROWS), PackageLimits(), "book.xlsx")

    def test_non_zip_payload_is_rejected(self) -> None:
        with self.assertRaises(PackageSecurityError):
            validate_embedded_package(b"not a zip", PackageLimits(), "book.xlsx")

    def test_entry_limit_applies_to_embedded_packages(self) -> None:
        with self.assertRaises(PackageSecurityError):
            validate_embedded_package(
                _workbook(WORKBOOK_ROWS), PackageLimits(max_entries=2), "book.xlsx"
            )


if __name__ == "__main__":
    unittest.main()
