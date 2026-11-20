"""Native tables as a merged-cell-aware grid rendered to a Markdown table."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field

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


def _escape(text: str) -> str:
    return text.replace("|", r"\|")


def grid_markdown(grid: Sequence[Sequence[GridCell]]) -> str:
    """Render the grid with its first row as the header.

    A header cell covered by a merge repeats the merged text, so every column has a
    name, and an empty header is named ``col_N``. Covered body cells stay empty, so a
    merged value is not repeated in the data.
    """

    if not grid:
        return ""

    def text(cell: GridCell) -> str:
        row, column = cell.origin
        return _escape("<br>".join(grid[row][column].lines))

    header = [text(cell) or f"col_{column + 1}" for column, cell in enumerate(grid[0])]
    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join("---" for _ in header) + " |",
    ]
    lines.extend(
        "| " + " | ".join("" if cell.merged else text(cell) for cell in row) + " |"
        for row in grid[1:]
    )
    return "\n".join(lines)
