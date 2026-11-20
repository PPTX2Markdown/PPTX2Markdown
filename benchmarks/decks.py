"""Build one small PowerPoint deck per feature, with python-pptx.

Each case pairs a deck with a check on the Markdown a converter produced from
it. Decks are generated, not stored, so anyone can rebuild and inspect them.
"""

from __future__ import annotations

import copy
import io
import re
import struct
import zlib
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


@dataclass(frozen=True)
class Case:
    name: str
    feature: str
    build: Callable[[Path], None]
    check: Callable[[str], bool]


def _png() -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        body = kind + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body))

    raw = b"\x00\xff\x00\x00" * 1
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )


def _blank(prs: Presentation):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _box(slide, text: str, x: float, y: float, w: float = 4, h: float = 1):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    shape.text_frame.text = text
    return shape


def _lines(markdown: str) -> list[str]:
    return [line.strip() for line in markdown.splitlines() if line.strip()]


def _table_rows(markdown: str) -> list[list[str]]:
    rows = []
    for line in _lines(markdown):
        if line.startswith("|") and not re.fullmatch(r"\|[\s:|-]+\|", line):
            cells = re.split(r"(?<!\\)\|", line.strip("|"))
            rows.append([cell.strip() for cell in cells])
    return rows


# --- cases -----------------------------------------------------------------


def build_slide_boundaries(path: Path) -> None:
    prs = Presentation()
    for word in ("alpha", "beta", "gamma"):
        _box(_blank(prs), f"Topic {word}", 1, 1)
    prs.save(path)


def check_slide_boundaries(markdown: str) -> bool:
    markers = re.findall(
        r"(?im)^(?:\[Page_\d+\]|<!--\s*slide number:\s*\d+\s*-->|#+\s*slide\s*\d+)", markdown
    )
    return len(markers) >= 3


def build_title(path: Path) -> None:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Quarterly Review"
    _box(slide, "Body text", 1, 3)
    prs.save(path)


def check_title(markdown: str) -> bool:
    return any(re.match(r"^#{1,6}\s+Quarterly Review", line) for line in _lines(markdown))


def build_two_columns(path: Path) -> None:
    prs = Presentation()
    slide = _blank(prs)
    # Two text columns whose paragraphs are staggered, so no line runs across both
    # columns; drawn out of reading order on purpose.
    _box(slide, "Right first", 5.2, 1.5)
    _box(slide, "Left second", 0.5, 2.6)
    _box(slide, "Left first", 0.5, 1.0)
    _box(slide, "Right second", 5.2, 2.3)
    prs.save(path)


def check_two_columns(markdown: str) -> bool:
    order = ["Left first", "Left second", "Right first", "Right second"]
    positions = [markdown.find(text) for text in order]
    return all(p >= 0 for p in positions) and positions == sorted(positions)


def build_nested_list(path: Path) -> None:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Plan"
    body = slide.placeholders[1].text_frame
    body.text = "Phase one"
    for text, level in (("Design", 1), ("Build", 1), ("Phase two", 0)):
        paragraph = body.add_paragraph()
        paragraph.text = text
        paragraph.level = level
    prs.save(path)


def check_nested_list(markdown: str) -> bool:
    raw = markdown.splitlines()
    top = next((line for line in raw if re.match(r"^\s*[-*+]\s+Phase one", line)), None)
    child = next((line for line in raw if re.match(r"^\s*[-*+]\s+Design", line)), None)
    if top is None or child is None:
        return False
    indent = lambda line: len(line) - len(line.lstrip())  # noqa: E731
    return indent(child) > indent(top)


def build_hyperlink(path: Path) -> None:
    prs = Presentation()
    paragraph = _box(_blank(prs), "", 1, 1).text_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = "documentation"
    run.hyperlink.address = "https://example.com/docs"
    prs.save(path)


def check_hyperlink(markdown: str) -> bool:
    return "[documentation](https://example.com/docs)" in markdown


def build_bold(path: Path) -> None:
    prs = Presentation()
    paragraph = _box(_blank(prs), "", 1, 1).text_frame.paragraphs[0]
    plain = paragraph.add_run()
    plain.text = "Revenue grew "
    bold = paragraph.add_run()
    bold.text = "12 percent"
    bold.font.bold = True
    prs.save(path)


def check_bold(markdown: str) -> bool:
    return bool(re.search(r"(\*\*|__)12 percent\1", markdown))


def build_escaping(path: Path) -> None:
    prs = Presentation()
    slide = _blank(prs)
    _box(slide, "# not a heading", 1, 1)
    _box(slide, "<script>alert(1)</script>", 1, 2.5)
    prs.save(path)


def check_escaping(markdown: str) -> bool:
    heading = any(re.match(r"^#\s+not a heading", line) for line in markdown.splitlines())
    raw_tag = re.search(r"(?<!\\)<script>", markdown)
    return "not a heading" in markdown and "alert(1)" in markdown and not heading and not raw_tag


def build_table(path: Path) -> None:
    prs = Presentation()
    slide = _blank(prs)
    table = slide.shapes.add_table(3, 3, Inches(1), Inches(1), Inches(6), Inches(1.5)).table
    values = [["Region", "Q1", "Q2"], ["Seoul", "10", "12"], ["Busan", "7", "a|b"]]
    for r, row in enumerate(values):
        for c, value in enumerate(row):
            table.cell(r, c).text = value
    prs.save(path)


def check_table(markdown: str) -> bool:
    rows = _table_rows(markdown)
    return (
        len(rows) == 3
        and all(len(row) == 3 for row in rows)
        and rows[0] == ["Region", "Q1", "Q2"]
        and rows[2][2] in {r"a\|b", "a|b"}
    )


def build_merged(path: Path) -> None:
    prs = Presentation()
    slide = _blank(prs)
    table = slide.shapes.add_table(4, 3, Inches(1), Inches(1), Inches(6), Inches(2)).table
    values = [
        ["Region", "Sales", ""],
        ["Seoul", "10", "12"],
        ["", "11", "13"],
        ["Busan", "7", "8"],
    ]
    for r, row in enumerate(values):
        for c, value in enumerate(row):
            table.cell(r, c).text = value
    table.cell(0, 1).merge(table.cell(0, 2))
    table.cell(1, 0).merge(table.cell(2, 0))
    prs.save(path)


def check_merged(markdown: str) -> bool:
    rows = _table_rows(markdown)
    if len(rows) != 4 or any(len(row) != 3 for row in rows):
        return False
    # The second Seoul row must still say which region it belongs to.
    return rows[2][0].strip("*") == "Seoul" and rows[2][1:] == ["11", "13"]


def _math_paragraph() -> etree._Element:
    """A paragraph holding the equation E = mc^2, as PowerPoint writes it."""

    paragraph = etree.Element(f"{{{A}}}p", nsmap={"a": A, "a14": A14, "m": M})
    math = etree.SubElement(
        etree.SubElement(etree.SubElement(paragraph, f"{{{A14}}}m"), f"{{{M}}}oMathPara"),
        f"{{{M}}}oMath",
    )
    run = etree.SubElement(math, f"{{{M}}}r")
    etree.SubElement(run, f"{{{M}}}t").text = "E=m"
    power = etree.SubElement(math, f"{{{M}}}sSup")
    for tag, text in (("e", "c"), ("sup", "2")):
        part_run = etree.SubElement(etree.SubElement(power, f"{{{M}}}{tag}"), f"{{{M}}}r")
        etree.SubElement(part_run, f"{{{M}}}t").text = text
    return paragraph


def build_math(path: Path) -> None:
    """PowerPoint wraps a shape with an equation in mc:AlternateContent: the a14
    choice has the equation, the fallback a copy without it (really a picture)."""

    prs = Presentation()
    shape = _box(_blank(prs), "Energy", 1, 1)
    element = shape._element
    fallback = copy.deepcopy(element)
    element.txBody.append(_math_paragraph())
    tree = element.getparent()
    index = tree.index(element)
    alternate = etree.Element(f"{{{MC}}}AlternateContent", nsmap={"mc": MC, "a14": A14})
    etree.SubElement(alternate, f"{{{MC}}}Choice", Requires="a14").append(element)
    etree.SubElement(alternate, f"{{{MC}}}Fallback").append(fallback)
    tree.insert(index, alternate)
    prs.save(path)


def check_math(markdown: str) -> bool:
    # Equivalent LaTeX spellings ("mc^{2}", "m{c}^{2}", "m c^2") compare equal.
    return "E=mc^2" in re.sub(r"[\s{}]", "", markdown)


def build_notes(path: Path) -> None:
    prs = Presentation()
    slide = _blank(prs)
    _box(slide, "Visible text", 1, 1)
    slide.notes_slide.notes_text_frame.text = "Remember the Seoul office"
    prs.save(path)


def check_notes(markdown: str) -> bool:
    return "Remember the Seoul office" in markdown


def build_hidden(path: Path) -> None:
    prs = Presentation()
    _box(_blank(prs), "Shown slide", 1, 1)
    hidden = _blank(prs)
    _box(hidden, "Backup material", 1, 1)
    hidden._element.set("show", "0")
    prs.save(path)


def check_hidden(markdown: str) -> bool:
    index = markdown.find("Backup material")
    if index < 0:
        return False
    window = markdown[max(0, index - 200) : index].casefold()
    return "hidden" in window


def build_image_alt(path: Path) -> None:
    prs = Presentation()
    picture = _blank(prs).shapes.add_picture(io.BytesIO(_png()), Inches(1), Inches(1))
    picture._element.nvPicPr.cNvPr.set("descr", "Revenue by quarter chart")
    prs.save(path)


def check_image_alt(markdown: str) -> bool:
    return "Revenue by quarter chart" in markdown


def build_chart(path: Path) -> None:
    prs = Presentation()
    data = CategoryChartData()
    data.categories = ["North", "South"]
    data.add_series("Units", (14, 23))
    _blank(prs).shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1), Inches(6), Inches(4), data
    )
    prs.save(path)


def check_chart(markdown: str) -> bool:
    return all(text in markdown for text in ("North", "South", "14", "23"))


def build_group(path: Path) -> None:
    prs = Presentation()
    group = _blank(prs).shapes.add_group_shape()
    box = group.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
    box.text_frame.text = "Inside a group"
    prs.save(path)


def check_group(markdown: str) -> bool:
    return "Inside a group" in markdown


def build_empty(path: Path) -> None:
    Presentation().save(path)


def check_empty(markdown: str) -> bool:
    return True  # converting without an error is the check


def build_font_size_heading(path: Path) -> None:
    prs = Presentation()
    slide = _blank(prs)
    title = _box(slide, "Market Overview", 0.5, 0.3, 8, 1)
    title.text_frame.paragraphs[0].runs[0].font.size = Pt(36)
    body = _box(slide, "Demand rose in every region this year.", 0.5, 2, 8, 1)
    body.text_frame.paragraphs[0].runs[0].font.size = Pt(16)
    prs.save(path)


def check_font_size_heading(markdown: str) -> bool:
    return any(re.match(r"^#{1,6}\s+Market Overview", line) for line in _lines(markdown))


CASES: tuple[Case, ...] = (
    Case(
        "slide-boundaries",
        "Each slide is delimited",
        build_slide_boundaries,
        check_slide_boundaries,
    ),
    Case("title", "Title placeholder becomes a heading", build_title, check_title),
    Case(
        "font-size-heading",
        "Large first text box becomes a heading",
        build_font_size_heading,
        check_font_size_heading,
    ),
    Case("two-columns", "Two-column reading order", build_two_columns, check_two_columns),
    Case("nested-list", "Nested bullet levels", build_nested_list, check_nested_list),
    Case("hyperlink", "Hyperlinks", build_hyperlink, check_hyperlink),
    Case("bold", "Bold runs", build_bold, check_bold),
    Case("escaping", "Typed Markdown/HTML stays literal", build_escaping, check_escaping),
    Case("table", "Tables (incl. a pipe in a cell)", build_table, check_table),
    Case("merged-cells", "Merged cells keep row meaning", build_merged, check_merged),
    Case("math", "Equations (OMML) as LaTeX", build_math, check_math),
    Case("notes", "Speaker notes", build_notes, check_notes),
    Case("hidden", "Hidden slides are marked", build_hidden, check_hidden),
    Case("image-alt", "Image alt text", build_image_alt, check_image_alt),
    Case("chart", "Chart data", build_chart, check_chart),
    Case("group", "Text inside grouped shapes", build_group, check_group),
    Case("empty", "Deck without slides", build_empty, check_empty),
)


def build_large(path: Path, slides: int = 300) -> None:
    """A deck for timing: text, a table and a chart on every slide."""

    source = Presentation()
    layout = source.slide_layouts[5]
    for index in range(slides):
        slide = source.slides.add_slide(layout)
        slide.shapes.title.text = f"Section {index + 1}"
        _box(slide, f"Notes for section {index + 1}\nSecond line", 0.5, 1.5, 4, 1)
        table = slide.shapes.add_table(3, 3, Inches(5), Inches(1.5), Inches(4), Inches(1)).table
        for r in range(3):
            for c in range(3):
                table.cell(r, c).text = f"{r}-{c}"
    source.save(path)


__all__ = ["CASES", "Case", "build_large"]
