"""Native tables as a merged-cell-aware grid rendered to a Markdown table."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field

from .model import MergedCells, TableCell

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"a": A_NS}

Span = tuple[int, int]
Bounds = tuple[list[Span], list[Span]]


@dataclass
class GridCell:
    """One grid position. A merged cell points at the cell its merge starts from."""

    origin: tuple[int, int]
    lines: list[str] = field(default_factory=list)
    merged: bool = False


def _int(element: ET.Element, name: str, default: int) -> int:
    try:
        return int(element.attrib.get(name, default))
    except ValueError:
        return default


def _flag(element: ET.Element, name: str) -> bool:
    return element.attrib.get(name) in {"1", "true"}


def table_grid(
    table: ET.Element, cell_lines: Callable[[ET.Element], list[str]]
) -> list[list[GridCell]]:
    """Map every ``a:tc`` onto the table grid, resolving horizontal and vertical merges.

    ``cell_lines`` turns a cell into its display lines.
    """

    rows: list[list[GridCell]] = []
    spans: dict[tuple[int, int], tuple[int, int]] = {}
    for row_index, row in enumerate(table.findall("./a:tr", NS)):
        cells: list[GridCell] = []
        for column, element in enumerate(row.findall("./a:tc", NS)):
            if _flag(element, "hMerge") and column > 0:
                cells.append(GridCell(origin=cells[column - 1].origin, merged=True))
            elif _flag(element, "vMerge") and row_index > 0 and column < len(rows[-1]):
                cells.append(GridCell(origin=rows[-1][column].origin, merged=True))
            else:
                cells.append(GridCell(origin=(row_index, column), lines=cell_lines(element)))
                spans[(row_index, column)] = (
                    max(1, _int(element, "rowSpan", 1)),
                    max(1, _int(element, "gridSpan", 1)),
                )
        rows.append(cells)

    width = max([len(table.findall("./a:tblGrid/a:gridCol", NS)), *(len(row) for row in rows)])
    for row_index, cells in enumerate(rows):
        for column in range(len(cells), width):
            # A row that omits the placeholder cells of a span still sits under it.
            origin = next(
                (
                    start
                    for start, (row_span, grid_span) in spans.items()
                    if start[0] <= row_index < start[0] + row_span
                    and start[1] <= column < start[1] + grid_span
                ),
                None,
            )
            if origin is None:
                cells.append(GridCell(origin=(row_index, column)))
            else:
                cells.append(GridCell(origin=origin, merged=True))
    return rows


def _split(sizes: list[int], start: int, end: int) -> list[Span]:
    total = sum(sizes)
    spans: list[Span] = []
    cursor, used = start, 0
    for index, size in enumerate(sizes):
        used += size
        edge = end if index == len(sizes) - 1 else start + (end - start) * used // total
        spans.append((cursor, edge))
        cursor = edge
    return spans


def cell_bounds(table: ET.Element, frame: tuple[int, int, int, int]) -> Bounds | None:
    """Column and row extents in slide coordinates, scaled to the frame like PowerPoint."""

    widths = [
        max(0, _int(column, "w", 0)) for column in table.findall("./a:tblGrid/a:gridCol", NS)
    ]
    heights = [max(0, _int(row, "h", 0)) for row in table.findall("./a:tr", NS)]
    left, top, right, bottom = frame
    if sum(widths) <= 0 or sum(heights) <= 0 or right <= left or bottom <= top:
        return None
    return _split(widths, left, right), _split(heights, top, bottom)


def cell_at(
    grid: Sequence[Sequence[GridCell]], bounds: Bounds, point: tuple[int, int]
) -> GridCell | None:
    """The cell (the start of its merge) under a slide point, or None outside the table."""

    columns, rows = bounds
    x, y = point
    row = next((i for i, (start, end) in enumerate(rows) if start <= y <= end), None)
    column = next((i for i, (start, end) in enumerate(columns) if start <= x <= end), None)
    return cell_in(grid, row, column)


def cell_in(grid: Sequence[Sequence[GridCell]], row: int | None, column: int | None):
    """The cell (the start of its merge) at a grid position, or None."""

    if row is None or column is None or row >= len(grid) or column >= len(grid[row]):
        return None
    origin_row, origin_column = grid[row][column].origin
    return grid[origin_row][origin_column]


MERGED_CELL_POLICIES: tuple[MergedCells, ...] = ("fill", "header", "blank")


def grid_rows(grid: Sequence[Sequence[GridCell]]) -> tuple[tuple[TableCell, ...], ...]:
    """The grid as model cells: merge starts carry text and spans, covered cells point back."""

    spans: dict[tuple[int, int], list[int]] = {}
    for row_index, row in enumerate(grid):
        for column, cell in enumerate(row):
            origin = cell.origin
            span = spans.setdefault(origin, [1, 1])
            span[0] = max(span[0], row_index - origin[0] + 1)
            span[1] = max(span[1], column - origin[1] + 1)
    return tuple(
        tuple(
            TableCell(merged_from=cell.origin)
            if cell.origin != (row_index, column)
            else TableCell(
                text="\n".join(cell.lines),
                row_span=spans[cell.origin][0],
                col_span=spans[cell.origin][1],
            )
            for column, cell in enumerate(row)
        )
        for row_index, row in enumerate(grid)
    )


def _cell_markdown(text: str) -> str:
    return "<br>".join(line for line in text.split("\n") if line).replace("|", r"\|")


def caption_rows(rows: Sequence[Sequence[TableCell]]) -> int:
    """1 when the first row is one cell merged across the whole table, i.e. a title.

    A title row above the column labels ("2024 results") would otherwise become
    the name of every column. A two-row table keeps it, since it is the only header.
    """

    if len(rows) < 3:
        return 0
    width = max(len(row) for row in rows)
    first = rows[0][0] if rows[0] else None
    if (
        width < 2
        or first is None
        or first.merged_from is not None
        or first.col_span < width
        or first.row_span != 1
        or not first.text.strip()
    ):
        return 0
    return 1


def rows_markdown(
    rows: Sequence[Sequence[TableCell]],
    *,
    header_rows: int = 1,
    merged_cells: MergedCells = "fill",
    caption_rows: int = 0,
) -> str:
    """Render model rows as a Markdown table.

    ``merged_cells`` decides what a position covered by a merge shows: ``fill``
    repeats the merged text everywhere (each row stays self-contained, which suits
    retrieval), ``header`` repeats it only in the header so every column keeps a
    name, and ``blank`` leaves every covered position empty. A merge across the
    whole width in the body is a section label and is never repeated across the
    row. Several header rows are combined per column as ``top / bottom``; an empty
    header is ``col_N``. The first ``caption_rows`` rows are written as a bold line
    above the table.
    """

    if not rows:
        return ""
    width = max(len(row) for row in rows)

    def origin_of(row: int, column: int) -> TableCell | None:
        if column >= len(rows[row]):
            return None
        cell = rows[row][column]
        if cell.merged_from is None:
            return cell
        origin_row, origin_column = cell.merged_from
        return rows[origin_row][origin_column]

    def text(row: int, column: int, *, in_header: bool) -> str:
        if column >= len(rows[row]):
            return ""
        cell = rows[row][column]
        if cell.merged_from is None:
            return _cell_markdown(cell.text)
        origin = origin_of(row, column)
        assert origin is not None
        if not in_header and origin.col_span >= width:
            return ""
        repeat = merged_cells == "fill" or (merged_cells == "header" and in_header)
        return _cell_markdown(origin.text) if repeat else ""

    caption_count = max(0, min(caption_rows, len(rows)))
    captions: list[str] = []
    for row in range(caption_count):
        for column in range(width):
            origin = origin_of(row, column)
            value = " ".join(origin.text.split("\n")).strip() if origin else ""
            if value and value not in captions:
                captions.append(value)
    body_rows = list(range(caption_count, len(rows)))
    header_count = max(0, min(header_rows, len(body_rows)))
    header: list[str] = []
    for column in range(width):
        parts: list[str] = []
        for row in body_rows[:header_count]:
            value = text(row, column, in_header=True)
            if value and (not parts or parts[-1] != value):
                parts.append(value)
        header.append(" / ".join(parts) or f"col_{column + 1}")

    lines = [
        caption if caption.startswith("**") and caption.endswith("**") else f"**{caption}**"
        for caption in captions
    ]
    if lines:
        lines.append("")
    lines += [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join("---" for _ in header) + " |",
    ]
    lines.extend(
        "| " + " | ".join(text(row, column, in_header=False) for column in range(width)) + " |"
        for row in body_rows[header_count:]
    )
    return "\n".join(lines)
