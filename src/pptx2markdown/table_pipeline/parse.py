"""Parse an OpenXML table element into a structured grid."""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"a": A_NS}


def _int_attr(elem: ET.Element, name: str, default: int = 0) -> int:
    val = elem.get(name)
    if val is None:
        return default
    try:
        return int(val)
    except ValueError:
        return default


def _cell_text(tc: ET.Element) -> str:
    paragraphs: list[str] = []
    for p in tc.findall(".//a:txBody/a:p", NS):
        runs = p.findall(".//a:t", NS)
        if not runs:
            continue
        text = "".join(t.text or "" for t in runs)
        text = re.sub(r"\s+", " ", text).strip()
        if text:
            paragraphs.append(text)
    return "\n".join(paragraphs)


def _find_table(root: ET.Element) -> ET.Element | None:
    if root.tag == f"{{{A_NS}}}tbl":
        return root
    return root.find(".//a:tbl", NS)


def parse_table_root(root: ET.Element, source: str = "<in-memory>") -> dict[str, object]:
    table = _find_table(root)
    if table is None:
        raise ValueError("No <a:tbl> found in input XML.")

    col_elems = table.findall("./a:tblGrid/a:gridCol", NS)
    col_widths = [_int_attr(c, "w", 0) for c in col_elems]
    n_cols = len(col_widths)

    tr_elems = table.findall("./a:tr", NS)
    n_rows = len(tr_elems)

    # Active vertical merges by column.
    active_v: dict[int, dict[str, object]] = {}
    rows_out: list[list[dict[str, object]]] = []
    origin_cells: list[dict[str, object]] = []

    for r_idx, tr in enumerate(tr_elems):
        tc_elems = tr.findall("./a:tc", NS)
        row_cells: list[dict[str, object]] = []
        current_origin_col: int | None = None
        touched_v_cols: set[int] = set()

        for c_idx, tc in enumerate(tc_elems):
            grid_span = _int_attr(tc, "gridSpan", 1)
            row_span = _int_attr(tc, "rowSpan", 1)
            is_h_merge = tc.get("hMerge") == "1"
            is_v_merge = tc.get("vMerge") == "1"
            text = _cell_text(tc)

            if is_h_merge:
                origin_col = current_origin_col if current_origin_col is not None else c_idx - 1
                row_cells.append(
                    {
                        "type": "hMerge",
                        "origin": [r_idx, max(origin_col, 0)],
                        "text": text,
                    }
                )
                current_origin_col = max(origin_col, 0)
                continue

            if is_v_merge:
                if c_idx in active_v:
                    origin = active_v[c_idx]["origin"]
                    remaining = int(active_v[c_idx]["remaining"]) - 1
                    touched_v_cols.add(c_idx)
                    if remaining <= 0:
                        del active_v[c_idx]
                    else:
                        active_v[c_idx]["remaining"] = remaining
                else:
                    origin = [max(r_idx - 1, 0), c_idx]
                row_cells.append(
                    {
                        "type": "vMerge",
                        "origin": origin,
                        "text": text,
                    }
                )
                current_origin_col = None
                continue

            # Origin cell.
            cell = {
                "type": "origin",
                "origin": [r_idx, c_idx],
                "text": text,
                "rowspan": row_span,
                "colspan": grid_span,
            }
            row_cells.append(cell)
            origin_cells.append(
                {
                    "row": r_idx,
                    "col": c_idx,
                    "text": text,
                    "rowspan": row_span,
                    "colspan": grid_span,
                }
            )
            current_origin_col = c_idx

            if row_span > 1:
                for covered_col in range(c_idx, c_idx + grid_span):
                    active_v[covered_col] = {
                        "origin": [r_idx, c_idx],
                        "remaining": row_span - 1,
                    }

        # If XML omitted vMerge placeholders, backfill from active spans.
        if n_cols and len(row_cells) < n_cols:
            for c_idx in range(len(row_cells), n_cols):
                if c_idx in active_v and c_idx not in touched_v_cols:
                    origin = active_v[c_idx]["origin"]
                    remaining = int(active_v[c_idx]["remaining"]) - 1
                    if remaining <= 0:
                        del active_v[c_idx]
                    else:
                        active_v[c_idx]["remaining"] = remaining
                    row_cells.append({"type": "vMerge", "origin": origin, "text": ""})
                else:
                    row_cells.append({"type": "empty", "text": ""})

        rows_out.append(row_cells)

    return {
        "source": source,
        "n_rows": n_rows,
        "n_cols": n_cols,
        "column_widths": col_widths,
        "rows": rows_out,
        "origin_cells": origin_cells,
    }


def parse_table_element(table_elem: ET.Element, source: str = "<in-memory>") -> dict[str, object]:
    return parse_table_root(table_elem, source=source)
