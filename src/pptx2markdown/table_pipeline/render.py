#!/usr/bin/env python3
"""Render table JSON into markdown/html/csv.

Supports:
1) Parsed table JSON produced by parse.py
2) Generic dense JSON rows (list[list[str]])

Batch mode:
- If no input path is provided, converts all ./artifacts/parsing_results/*.json into
  markdown tables and writes them to ./artifacts/tables/*.md.
"""

from __future__ import annotations

import argparse
import csv
import html
import io
import json
from pathlib import Path
import re
import sys
from typing import Any


MODE_MARKDOWN = "markdown-flat"
MODE_HTML = "html"
MODE_CSV = "csv"
FILL_NONE = "none"
FILL_H = "horizontal"
FILL_V = "vertical"
FILL_BOTH = "both"
HEADER_AUTO = "auto"


def _normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _fill_allowed(fill_merged: str, cell_type: str) -> bool:
    if fill_merged == FILL_BOTH:
        return cell_type in {"hMerge", "vMerge"}
    if fill_merged == FILL_H:
        return cell_type == "hMerge"
    if fill_merged == FILL_V:
        return cell_type == "vMerge"
    return False


def _load_json(path: Path) -> dict[str, Any] | list[Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _is_parsed_table_format(payload: Any) -> bool:
    return isinstance(payload, dict) and isinstance(payload.get("rows"), list)


def _dense_grid_from_generic_rows(payload: Any) -> list[list[str]]:
    if isinstance(payload, dict):
        rows = payload.get("rows")
    else:
        rows = payload
    if not isinstance(rows, list):
        raise ValueError("Unsupported JSON structure: expected list rows or dict with rows.")
    dense: list[list[str]] = []
    for row in rows:
        if not isinstance(row, list):
            raise ValueError("Unsupported row structure: expected list.")
        dense.append([_normalize_text(str(cell)) if cell is not None else "" for cell in row])
    return dense


def _dense_grid_from_parsed_table(payload: dict[str, Any], fill_merged: str) -> list[list[str]]:
    rows = payload.get("rows")
    n_rows = int(payload.get("n_rows", len(rows or [])))
    n_cols = int(payload.get("n_cols", 0))
    if not isinstance(rows, list):
        raise ValueError("Invalid parsed-table JSON: rows missing.")
    if n_cols <= 0 and rows:
        n_cols = max((len(r) for r in rows if isinstance(r, list)), default=0)
    dense = [["" for _ in range(n_cols)] for _ in range(n_rows)]
    origin_texts: dict[tuple[int, int], str] = {}

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
            elif ctype in {"hMerge", "vMerge"} and _fill_allowed(fill_merged, ctype):
                origin = cell.get("origin")
                if (
                    isinstance(origin, list)
                    and len(origin) == 2
                    and isinstance(origin[0], int)
                    and isinstance(origin[1], int)
                ):
                    dense[r_idx][c_idx] = origin_texts.get((origin[0], origin[1]), "")
    return dense


def _origin_row_with_spans(payload: dict[str, Any]) -> list[list[dict[str, Any]]]:
    rows = payload.get("rows")
    if not isinstance(rows, list):
        raise ValueError("Invalid parsed-table JSON: rows missing.")
    rendered: list[list[dict[str, Any]]] = []
    for row in rows:
        if not isinstance(row, list):
            continue
        out_row: list[dict[str, Any]] = []
        for cell in row:
            if not isinstance(cell, dict):
                continue
            if cell.get("type") != "origin":
                continue
            out_row.append(
                {
                    "text": _normalize_text(str(cell.get("text", ""))),
                    "rowspan": int(cell.get("rowspan", 1)),
                    "colspan": int(cell.get("colspan", 1)),
                }
            )
        rendered.append(out_row)
    return rendered


def _auto_header_rows(payload: dict[str, Any] | list[Any]) -> int:
    # Conservative default: top one row only.
    return 1


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


def _render_markdown_flat(
    dense: list[list[str]], header_rows: int, use_header_rows: bool
) -> str:
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
    dense = _dense_grid_from_parsed_table(parsed_table, fill_merged=fill_merged)
    return _render_markdown_flat(
        dense=dense,
        header_rows=max(0, int(header_rows)),
        use_header_rows=header_rows > 0,
    )


def _render_html(
    payload: dict[str, Any] | list[Any], dense: list[list[str]], header_rows: int, use_header_rows: bool
) -> str:
    if isinstance(payload, dict) and _is_parsed_table_format(payload):
        rows = _origin_row_with_spans(payload)
        lines = ["<table>"]
        for r_idx, row in enumerate(rows):
            lines.append("  <tr>")
            tag = "th" if use_header_rows and r_idx < header_rows else "td"
            for cell in row:
                attrs = []
                if cell["rowspan"] > 1:
                    attrs.append(f'rowspan="{cell["rowspan"]}"')
                if cell["colspan"] > 1:
                    attrs.append(f'colspan="{cell["colspan"]}"')
                attr = (" " + " ".join(attrs)) if attrs else ""
                text = html.escape(cell["text"]).replace("\n", "<br>")
                lines.append(f"    <{tag}{attr}>{text}</{tag}>")
            lines.append("  </tr>")
        lines.append("</table>")
        return "\n".join(lines) + "\n"

    if not dense:
        return "<table></table>\n"
    n_cols = max(len(r) for r in dense)
    dense = [r + [""] * (n_cols - len(r)) for r in dense]
    lines = ["<table>"]
    for r_idx, row in enumerate(dense):
        lines.append("  <tr>")
        tag = "th" if use_header_rows and r_idx < header_rows else "td"
        for v in row:
            lines.append(f"    <{tag}>{html.escape(v).replace(chr(10), '<br>')}</{tag}>")
        lines.append("  </tr>")
    lines.append("</table>")
    return "\n".join(lines) + "\n"


def _render_csv(dense: list[list[str]], header_rows: int, use_header_rows: bool) -> str:
    if not dense:
        return ""
    n_cols = max(len(r) for r in dense)
    dense = [r + [""] * (n_cols - len(r)) for r in dense]

    output = io.StringIO()
    writer = csv.writer(output)
    if use_header_rows and header_rows > 0:
        writer.writerow(_make_header_names(dense, header_rows))
        rows = dense[header_rows:]
    else:
        rows = dense
    writer.writerows(rows)
    return output.getvalue()


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="Render table JSON into markdown/html/csv.")
    parser.add_argument("input_json", nargs="?", help="Path to input table JSON.")
    parser.add_argument(
        "--mode",
        choices=[MODE_MARKDOWN, MODE_HTML, MODE_CSV],
        default=MODE_MARKDOWN,
        help="Output mode. Default: markdown-flat",
    )
    parser.add_argument(
        "--header-rows",
        default=HEADER_AUTO,
        help='Number of top rows treated as header (e.g. 1,2) or "auto". Default: auto',
    )
    parser.add_argument(
        "--fill-merged",
        choices=[FILL_NONE, FILL_H, FILL_V, FILL_BOTH],
        default=FILL_BOTH,
        help="How merged cells are filled in flattened output. Default: both",
    )
    parser.add_argument(
        "-o",
        "--output",
        help="Optional output file path. If omitted, prints to stdout.",
    )
    args = parser.parse_args(argv[1:])

    def _resolve_header_rows(payload: dict[str, Any] | list[Any]) -> tuple[int, bool] | None:
        if args.header_rows == HEADER_AUTO:
            return _auto_header_rows(payload if isinstance(payload, (dict, list)) else {}), True
        try:
            header_rows_i = max(0, int(args.header_rows))
            return header_rows_i, header_rows_i > 0
        except ValueError:
            print('[ERROR] --header-rows must be an integer or "auto".', file=sys.stderr)
            return None

    # Batch mode: no input argument.
    if not args.input_json:
        base_dir = Path.cwd()
        parsing_dir = base_dir / "artifacts" / "parsing_results"
        tables_dir = base_dir / "artifacts" / "tables"
        parsing_dir.mkdir(parents=True, exist_ok=True)
        tables_dir.mkdir(parents=True, exist_ok=True)

        json_files = sorted(
            p for p in parsing_dir.glob("*.json") if p.is_file() and p.name != "manifest.json"
        )
        if not json_files:
            print("[ERROR] no .json files found in ./artifacts/parsing_results.", file=sys.stderr)
            return 1

        if args.mode != MODE_MARKDOWN:
            print("[INFO] no-input mode forces markdown-flat output into ./artifacts/tables.")

        ok_count = 0
        fail_count = 0
        skip_count = 0
        for input_path in json_files:
            try:
                payload = _load_json(input_path)
                if isinstance(payload, dict) and "rows" not in payload:
                    print(f"[INFO] skipped non-table json: {input_path.name}")
                    skip_count += 1
                    continue
                if _is_parsed_table_format(payload):
                    dense = _dense_grid_from_parsed_table(payload, fill_merged=args.fill_merged)
                else:
                    dense = _dense_grid_from_generic_rows(payload)

                header_info = _resolve_header_rows(payload)
                if header_info is None:
                    return 1
                header_rows, use_header_rows = header_info

                rendered = _render_markdown_flat(
                    dense=dense,
                    header_rows=header_rows,
                    use_header_rows=use_header_rows,
                )
                out_path = tables_dir / f"{input_path.stem}.md"
                out_path.write_text(rendered, encoding="utf-8")
                print(f"[OK] wrote table: {out_path}")
                ok_count += 1
            except Exception as exc:
                print(f"[ERROR] failed to convert: {input_path} ({exc})", file=sys.stderr)
                fail_count += 1

        print(f"[OK] batch completed: {ok_count} converted, {skip_count} skipped, {fail_count} failed")
        return 1 if fail_count > 0 else 0

    # Single input mode.
    input_path = Path(args.input_json).expanduser().resolve()
    if not input_path.exists() or not input_path.is_file():
        print(f"[ERROR] input file not found: {input_path}", file=sys.stderr)
        return 1

    payload = _load_json(input_path)
    if _is_parsed_table_format(payload):
        dense = _dense_grid_from_parsed_table(payload, fill_merged=args.fill_merged)
    else:
        dense = _dense_grid_from_generic_rows(payload)

    header_info = _resolve_header_rows(payload)
    if header_info is None:
        return 1
    header_rows, use_header_rows = header_info

    if args.mode == MODE_MARKDOWN:
        rendered = _render_markdown_flat(dense=dense, header_rows=header_rows, use_header_rows=use_header_rows)
    elif args.mode == MODE_HTML:
        rendered = _render_html(
            payload=payload,
            dense=dense,
            header_rows=header_rows,
            use_header_rows=use_header_rows,
        )
    else:
        rendered = _render_csv(dense=dense, header_rows=header_rows, use_header_rows=use_header_rows)

    if args.output:
        out_path = Path(args.output).expanduser().resolve()
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(rendered, encoding="utf-8")
        print(f"[OK] wrote output: {out_path}")
    else:
        print(rendered, end="")
    return 0


if __name__ == "__main__":
    print(
        "[INFO] render.py is an internal pipeline module.\n"
        "Use `python run.py` from table_pipeline/ as the entrypoint."
    )
    raise SystemExit(1)
