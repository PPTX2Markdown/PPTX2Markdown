"""Render a parsed OpenXML table as Markdown."""

from __future__ import annotations

import re
from typing import Any

FILL_H = "horizontal"
FILL_V = "vertical"
FILL_BOTH = "both"
FILL_HEADER = "header"
FILL_NONE = "none"

# Public merged-cell policies: fill repeats a merged value in every cell it
# covers, so each row reads on its own; blank leaves covered cells empty.
MERGED_CELL_POLICIES = ("fill", "blank")


def fill_mode(merged_cells: str) -> str:
    """The renderer fill mode for a public merged-cell policy."""
    if merged_cells == "fill":
        return FILL_BOTH
    if merged_cells == "blank":
        return FILL_NONE
    raise ValueError(f"merged_cells must be one of {', '.join(MERGED_CELL_POLICIES)}")


def _normalize_text(text: str) -> str:
    normalized: list[str] = []
    for line in text.splitlines():
        leading_spaces = len(line) - len(line.lstrip(" "))
        body = re.sub(r"\s+", " ", line.lstrip()).strip()
        if body:
            normalized.append((" " * leading_spaces) + body)
    return "\n".join(normalized)


def _fill_allowed(
    fill_merged: str,
    cell_type: str,
    *,
    row_index: int,
    header_rows: int,
) -> bool:
    if fill_merged == FILL_HEADER:
        return row_index < header_rows and cell_type in {"hMerge", "vMerge"}
    if fill_merged == FILL_BOTH:
        return cell_type in {"hMerge", "vMerge"}
    if fill_merged == FILL_H:
        return cell_type == "hMerge"
    if fill_merged == FILL_V:
        return cell_type == "vMerge"
    return False


def _dense_grid_from_parsed_table(
    payload: dict[str, Any],
    fill_merged: str,
    *,
    header_rows: int,
) -> list[list[str]]:
    rows = payload.get("rows")
    n_rows = int(payload.get("n_rows", len(rows or [])))
    n_cols = int(payload.get("n_cols", 0))
    if not isinstance(rows, list):
        raise ValueError("Invalid parsed-table JSON: rows missing.")
    if n_cols <= 0 and rows:
        n_cols = max((len(r) for r in rows if isinstance(r, list)), default=0)
    dense = [["" for _ in range(n_cols)] for _ in range(n_rows)]
    origin_texts: dict[tuple[int, int], str] = {}
    pointers: dict[tuple[int, int], tuple[int, int]] = {}
    for r_idx, row in enumerate(rows[:n_rows]):
        if not isinstance(row, list):
            continue
        for c_idx, cell in enumerate(row[:n_cols]):
            origin = cell.get("origin") if isinstance(cell, dict) else None
            if (
                isinstance(cell, dict)
                and cell.get("type") in {"hMerge", "vMerge"}
                and isinstance(origin, list)
                and len(origin) == 2
                and all(isinstance(value, int) for value in origin)
            ):
                pointers[(r_idx, c_idx)] = (origin[0], origin[1])

    def merge_start(position: tuple[int, int]) -> tuple[int, int]:
        # A covered cell may point at another covered cell (inside a 2D block);
        # follow the chain to the cell the merge starts from.
        seen = {position}
        while position in pointers and pointers[position] not in seen:
            position = pointers[position]
            seen.add(position)
        return position

    for r_idx, row in enumerate(rows[:n_rows]):
        if not isinstance(row, list):
            continue
        for c_idx, cell in enumerate(row[:n_cols]):
            if not isinstance(cell, dict):
                continue
            ctype = cell.get("type", "")
            if ctype == "origin":
                text = _normalize_text(str(cell.get("text", "")))
                dense[r_idx][c_idx] = text
                origin_texts[(r_idx, c_idx)] = text
            elif ctype in {"hMerge", "vMerge"} and _fill_allowed(
                fill_merged,
                ctype,
                row_index=r_idx,
                header_rows=header_rows,
            ):
                if (r_idx, c_idx) in pointers:
                    dense[r_idx][c_idx] = origin_texts.get(merge_start((r_idx, c_idx)), "")
    return dense


def _make_header_names(dense: list[list[str]], header_rows: int) -> list[str]:
    if not dense:
        return []
    n_cols = max(len(r) for r in dense)
    names: list[str] = []
    for c_idx in range(n_cols):
        parts: list[str] = []
        for r_idx in range(min(header_rows, len(dense))):
            row = dense[r_idx]
            token = _normalize_text(row[c_idx]) if c_idx < len(row) else ""
            if token:
                if not parts or parts[-1] != token:
                    parts.append(token)
        names.append(" - ".join(parts) if parts else f"col_{c_idx + 1}")
    return names


def _escape_md(text: str) -> str:
    return text.replace("|", r"\|").replace("\n", "<br>")


def _render_markdown_flat(dense: list[list[str]], header_rows: int, use_header_rows: bool) -> str:
    if not dense:
        return ""
    n_cols = max(len(r) for r in dense)
    dense = [r + [""] * (n_cols - len(r)) for r in dense]

    if use_header_rows and header_rows > 0:
        headers = _make_header_names(dense, header_rows)
        body = dense[header_rows:]
    else:
        headers = [f"col_{i + 1}" for i in range(n_cols)]
        body = dense

    lines = []
    lines.append("| " + " | ".join(_escape_md(h) for h in headers) + " |")
    lines.append("|" + "|".join("---" for _ in headers) + "|")
    for row in body:
        lines.append("| " + " | ".join(_escape_md(v) for v in row) + " |")
    return "\n".join(lines) + "\n"


def render_parsed_table_to_markdown(
    parsed_table: dict[str, Any],
    header_rows: int = 1,
    fill_merged: str = FILL_BOTH,
) -> str:
    """Public API for rendering parsed-table payload into markdown."""
    normalized_header_rows = max(0, int(header_rows))
    dense = _dense_grid_from_parsed_table(
        parsed_table,
        fill_merged=fill_merged,
        header_rows=normalized_header_rows,
    )
    return _render_markdown_flat(
        dense=dense,
        header_rows=normalized_header_rows,
        use_header_rows=header_rows > 0,
    )
