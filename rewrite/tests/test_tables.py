from __future__ import annotations

import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path

import pptx2markdown
from pptx2markdown.model import ImageBlock, TableBlock, TableCell
from pptx2markdown.tables import caption_rows, grid_rows, rows_markdown, table_grid

from .pptx_factory import build_test_pptx

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _p(text: str, properties: str = "") -> str:
    return f'<a:p>{properties}<a:r><a:t xml:space="preserve">{text}</a:t></a:r></a:p>'


def _tc(body: str = "", attributes: str = "", cell_properties: str = "") -> str:
    paragraphs = body or "<a:p/>"
    return (
        f"<a:tc {attributes}><a:txBody><a:bodyPr/><a:lstStyle/>{paragraphs}</a:txBody>"
        f"<a:tcPr>{cell_properties}</a:tcPr></a:tc>"
    )


def _table(rows: list[list[str]], *, columns: int) -> str:
    grid = "".join('<a:gridCol w="1000000"/>' for _ in range(columns))
    body = "".join(f'<a:tr h="500000">{"".join(row)}</a:tr>' for row in rows)
    return f"<a:tbl><a:tblPr/><a:tblGrid>{grid}</a:tblGrid>{body}</a:tbl>"


def _frame(shape_id: int, x: int, y: int, table: str, columns: int, rows: int) -> str:
    return f"""
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="{shape_id}" name="Table {shape_id}"/>
          <p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{columns * 1000000}" cy="{rows * 500000}"/>
        </p:xfrm>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
          {table}
        </a:graphicData></a:graphic>
      </p:graphicFrame>"""


def _picture(shape_id: int, x: int, y: int, cx: int, cy: int, descr: str) -> str:
    return f"""
      <p:pic>
        <p:nvPicPr><p:cNvPr id="{shape_id}" name="Picture {shape_id}" descr="{descr}"/>
          <p:cNvPicPr/><p:nvPr/></p:nvPicPr>
        <p:blipFill><a:blip r:embed="rIdImage"/></p:blipFill>
        <p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm></p:spPr>
      </p:pic>"""


BULLET = '<a:pPr><a:buChar char="•"/></a:pPr>'
NUMBERED = '<a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr>'
NESTED_NUMBERED = '<a:pPr lvl="1"><a:buAutoNum type="arabicPeriod"/></a:pPr>'
FILL = '<a:blipFill><a:blip r:embed="rIdImage"/></a:blipFill>'

# Table 2: header merged across two columns, an empty header, a vertical merge in the
# body, lists, a line break, a pipe and a hyperlink.
MERGED = _table(
    [
        [_tc(_p("Group"), 'gridSpan="2"'), _tc(attributes='hMerge="1"'), _tc()],
        [
            _tc(_p("Region"), 'rowSpan="2"'),
            _tc(_p("One", BULLET) + _p("Two", BULLET) + _p("Plain")),
            _tc(_p("a|b")),
        ],
        [
            _tc(attributes='vMerge="1"'),
            _tc(_p("First", NUMBERED) + _p("Inner", NESTED_NUMBERED) + _p("Second", NUMBERED)),
            _tc(
                "<a:p><a:r><a:t>Narrow</a:t></a:r><a:br/><a:r><a:t>label</a:t></a:r></a:p>"
                '<a:p><a:r><a:rPr><a:hlinkClick r:id="rIdLink"/></a:rPr>'
                "<a:t>site</a:t></a:r></a:p>"
            ),
        ],
    ],
    columns=3,
)

# Table 3 sits at y=3,000,000 with a picture over cell (1, 1), a picture mostly outside
# the table, and a cell whose background is a picture.
PICTURES = _table(
    [
        [_tc(_p("A")), _tc(_p("B"))],
        [_tc(_p("left"), cell_properties=FILL), _tc(_p("right"))],
    ],
    columns=2,
)

SLIDE = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="{A}" xmlns:r="{R}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      {_frame(2, 0, 0, MERGED, columns=3, rows=3)}
      {_frame(3, 0, 3000000, PICTURES, columns=2, rows=2)}
      {_picture(4, 1100000, 3550000, 800000, 400000, "Cell photo")}
      {_picture(5, 1600000, 3600000, 1000000, 300000, "Overhanging photo")}
    </p:spTree>
  </p:cSld>
</p:sld>
"""


class TableTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = build_test_pptx(Path(directory) / "tables.pptx", SLIDE)
            with pptx2markdown.parse(source) as presentation:
                cls.document = presentation.document
                cls.markdown = presentation.render_markdown(
                    asset_url=lambda asset_id: f"/assets/{asset_id}"
                )
        cls.blocks = {block.shape_id: block for block in cls.document.slides[0].blocks}

    def test_merges_lists_breaks_and_links(self) -> None:
        block = self.blocks["2"]
        self.assertIsInstance(block, TableBlock)
        self.assertEqual(
            block.markdown.splitlines(),
            [
                "| Group | Group | col_3 |",
                "| --- | --- | --- |",
                r"| Region | - One<br>- Two<br>Plain | a\|b |",
                "| Region | 1. First<br>  1. Inner<br>2. Second | Narrow label<br>[site](https://openai.com) |",
            ],
        )

    def test_rows_keep_the_grid_with_spans_and_merge_origins(self) -> None:
        rows = self.blocks["2"].rows
        self.assertEqual(rows[0][0], TableCell(text="Group", col_span=2))
        self.assertEqual(rows[0][1], TableCell(merged_from=(0, 0)))
        self.assertEqual(rows[1][0], TableCell(text="Region", row_span=2))
        self.assertEqual(rows[2][0], TableCell(merged_from=(1, 0)))
        self.assertEqual(rows[1][1].text, "- One\n- Two\nPlain")

    def test_picture_on_a_cell_is_written_into_that_cell(self) -> None:
        block = self.blocks["3"]
        self.assertIsInstance(block, TableBlock)
        rows = block.markdown.splitlines()
        asset_id = block.asset_ids[0]
        self.assertEqual(block.asset_ids, (asset_id,))
        self.assertEqual(
            rows[2],
            f"| left<br>![image](asset://{asset_id}) | right<br>![Cell photo](asset://{asset_id}) |",
        )
        self.assertNotIn("4", self.blocks)
        self.assertIn(f"![Cell photo](/assets/{asset_id})", self.markdown)

    def test_picture_mostly_outside_the_table_stays_an_image_block(self) -> None:
        self.assertIsInstance(self.blocks["5"], ImageBlock)


class GridTests(unittest.TestCase):
    def test_rows_that_omit_span_placeholders_stay_aligned(self) -> None:
        table = ET.fromstring(
            f'<a:tbl xmlns:a="{A}"><a:tblGrid><a:gridCol w="1"/><a:gridCol w="1"/>'
            '<a:gridCol w="1"/></a:tblGrid>'
            '<a:tr><a:tc rowSpan="2"><a:txBody><a:p><a:r><a:t>Tall</a:t></a:r></a:p></a:txBody>'
            "</a:tc><a:tc/><a:tc/></a:tr>"
            "<a:tr/></a:tbl>"
        )

        def lines(cell: ET.Element) -> list[str]:
            text = "".join(node.text or "" for node in cell.iter(f"{{{A}}}t"))
            return [text] if text else []

        grid = table_grid(table, lines)
        self.assertEqual([cell.merged for cell in grid[1]], [True, False, False])
        self.assertEqual(grid[1][0].origin, (0, 0))
        self.assertEqual(
            rows_markdown(grid_rows(grid)).splitlines()[0], "| Tall | col_2 | col_3 |"
        )


def _cell(text: str = "", **spans: int) -> TableCell:
    return TableCell(text=text, **spans)


def _covered(row: int, column: int) -> TableCell:
    return TableCell(merged_from=(row, column))


# A 2x2 block merge in the body, next to a vertical merge, under a two-row header.
BLOCK = (
    (_cell("Year", row_span=2), _cell("2024", col_span=2), _covered(0, 1), _cell("Note")),
    (_covered(0, 0), _cell("Q1"), _cell("Q2"), _cell()),
    (_cell("Sales"), _cell("Flat", row_span=2, col_span=2), _covered(2, 1), _cell("x")),
    (_cell("Cost"), _covered(2, 1), _covered(2, 1), _covered(2, 3)),
)


class MergedCellPolicyTests(unittest.TestCase):
    def test_merge_origins_must_point_at_an_earlier_merge_start(self) -> None:
        for broken in (_covered(5, 5), _covered(0, 1), _covered(1, 1)):
            with self.subTest(origin=broken.merged_from), self.assertRaises(ValueError):
                TableBlock(markdown="x", rows=((_cell("a"), _covered(0, 0)), (_cell(), broken)))

    def test_fill_repeats_merged_values_in_every_covered_cell(self) -> None:
        self.assertEqual(
            rows_markdown(BLOCK, header_rows=1, merged_cells="fill").splitlines(),
            [
                "| Year | 2024 | 2024 | Note |",
                "| --- | --- | --- | --- |",
                "| Year | Q1 | Q2 |  |",
                "| Sales | Flat | Flat | x |",
                "| Cost | Flat | Flat | x |",
            ],
        )

    def test_header_repeats_only_in_the_header_row(self) -> None:
        self.assertEqual(
            rows_markdown(BLOCK, merged_cells="header").splitlines()[2:],
            ["|  | Q1 | Q2 |  |", "| Sales | Flat |  | x |", "| Cost |  |  |  |"],
        )

    def test_blank_leaves_every_covered_cell_empty(self) -> None:
        self.assertEqual(
            rows_markdown(BLOCK, merged_cells="blank").splitlines()[0],
            "| Year | 2024 | col_3 | Note |",
        )

    def test_header_rows_are_combined_per_column(self) -> None:
        self.assertEqual(
            rows_markdown(BLOCK, header_rows=2).splitlines()[:3],
            [
                "| Year | 2024 / Q1 | 2024 / Q2 | Note |",
                "| --- | --- | --- | --- |",
                "| Sales | Flat | Flat | x |",
            ],
        )

    def test_without_a_header_row_columns_are_numbered(self) -> None:
        self.assertEqual(
            rows_markdown(BLOCK, header_rows=0).splitlines()[:3],
            [
                "| col_1 | col_2 | col_3 | col_4 |",
                "| --- | --- | --- | --- |",
                "| Year | 2024 | 2024 | Note |",
            ],
        )


if __name__ == "__main__":
    unittest.main()


TITLED = (
    (_cell("**2024 results**", col_span=3), _covered(0, 0), _covered(0, 0)),
    (_cell("Region"), _cell("Q1"), _cell("Q2")),
    (_cell("Domestic", col_span=3), _covered(2, 0), _covered(2, 0)),
    (_cell("Seoul"), _cell("10"), _cell("12")),
)


class CaptionTests(unittest.TestCase):
    def test_a_full_width_first_row_is_a_title(self) -> None:
        self.assertEqual(caption_rows(TITLED), 1)
        self.assertEqual(caption_rows(TITLED[:2]), 0)
        self.assertEqual(caption_rows(BLOCK), 0)

    def test_title_is_written_once_above_the_table_and_sections_are_not_repeated(self) -> None:
        self.assertEqual(
            rows_markdown(TITLED, caption_rows=1).splitlines(),
            [
                "**2024 results**",
                "",
                "| Region | Q1 | Q2 |",
                "| --- | --- | --- |",
                "| Domestic |  |  |",
                "| Seoul | 10 | 12 |",
            ],
        )

    def test_without_a_caption_the_title_names_the_columns(self) -> None:
        self.assertEqual(
            rows_markdown(TITLED).splitlines()[0],
            "| **2024 results** | **2024 results** | **2024 results** |",
        )
