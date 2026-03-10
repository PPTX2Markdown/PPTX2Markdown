#!/usr/bin/env python3
"""Image-table pipeline service: PaddleOCR layout gate + Surya table parsing.

This module classifies an input image first. Only images predicted as "table"
are sent to Surya table recognition. Output markdown is rendered through the
existing parsed-table format used by table_parser.
"""

from __future__ import annotations

from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Sequence, Tuple


MODEL_NAME = "PP-DocLayout_plus-L"

_LAYOUT_MODEL: Any = None
_SURYA_MODELS: Optional[Tuple[Any, Any, Any, Any]] = None

_CLASSIFY_CACHE: Dict[str, Dict[str, Any]] = {}
_RESULT_CACHE: Dict[str, Dict[str, Any]] = {}
_SUPPORTED_LAYOUT_SUFFIXES = {
    ".png",
    ".jpg",
    ".jpeg",
    ".bmp",
    ".webp",
    ".tif",
    ".tiff",
}
_DEFAULT_BG_GRAY = 192


def _to_builtin(obj: Any) -> Any:
    if hasattr(obj, "model_dump"):
        return _to_builtin(obj.model_dump())
    if isinstance(obj, dict):
        return {k: _to_builtin(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_to_builtin(v) for v in obj]
    if hasattr(obj, "tolist"):
        try:
            return obj.tolist()
        except Exception:
            pass
    if hasattr(obj, "__dict__"):
        return {k: _to_builtin(v) for k, v in vars(obj).items()}
    return obj


def _normalize_boxes(page_result: Any) -> List[Dict[str, Any]]:
    if isinstance(page_result, dict):
        if isinstance(page_result.get("res"), list):
            return page_result["res"]
        if isinstance(page_result.get("boxes"), list):
            return page_result["boxes"]
        if isinstance(page_result.get("layout"), list):
            return page_result["layout"]

    if hasattr(page_result, "json"):
        try:
            j = page_result.json()
            if isinstance(j, dict):
                return _normalize_boxes(j)
        except Exception:
            pass

    if hasattr(page_result, "__dict__"):
        return _normalize_boxes(vars(page_result))
    return []


def _get_label(item: Dict[str, Any]) -> Optional[str]:
    for key in ("label", "category_name", "cls_name", "type"):
        value = item.get(key)
        if isinstance(value, str):
            return value.strip().lower()
    return None


def _get_bbox(item: Dict[str, Any]) -> Optional[List[float]]:
    for key in ("bbox", "box", "coordinate"):
        value = item.get(key)
        if isinstance(value, (list, tuple)) and len(value) == 4:
            return [float(value[0]), float(value[1]), float(value[2]), float(value[3])]
    return None


def _classify_page(boxes: Sequence[Dict[str, Any]]) -> Tuple[str, Dict[str, float]]:
    area_by_label: Dict[str, float] = {}
    for item in boxes:
        label = _get_label(item)
        bbox = _get_bbox(item)
        if not label or not bbox:
            continue
        area = max(0.0, bbox[2] - bbox[0]) * max(0.0, bbox[3] - bbox[1])
        area_by_label[label] = area_by_label.get(label, 0.0) + area

    score = {
        "table": area_by_label.get("table", 0.0),
        "chart": area_by_label.get("chart", 0.0),
        "text": (
            area_by_label.get("text", 0.0)
            + area_by_label.get("document title", 0.0)
            + area_by_label.get("paragraph title", 0.0)
            + area_by_label.get("abstract", 0.0)
            + area_by_label.get("header", 0.0)
            + area_by_label.get("footer", 0.0)
            + area_by_label.get("references", 0.0)
            + area_by_label.get("footnote", 0.0)
            + area_by_label.get("sidebar text", 0.0)
            + area_by_label.get("algorithm", 0.0)
            + area_by_label.get("formula", 0.0)
            + area_by_label.get("formula number", 0.0)
            + area_by_label.get("figure_table title", 0.0)
        ),
        "other": area_by_label.get("image", 0.0) + area_by_label.get("seal", 0.0),
    }

    if not boxes or all(v <= 0 for v in score.values()):
        return "other", score
    return max(score, key=score.get), score


def _load_layout_model() -> Any:
    global _LAYOUT_MODEL
    if _LAYOUT_MODEL is None:
        from paddleocr import LayoutDetection  # type: ignore

        _LAYOUT_MODEL = LayoutDetection(model_name=MODEL_NAME)
    return _LAYOUT_MODEL


def classify_image(image_path: Path) -> Dict[str, Any]:
    resolved = str(image_path.expanduser().resolve())
    cached = _CLASSIFY_CACHE.get(resolved)
    if cached is not None:
        return cached

    model = _load_layout_model()
    results = model.predict(resolved)
    page = results[0] if isinstance(results, list) and results else results
    raw_page = _to_builtin(page)
    boxes = _normalize_boxes(raw_page)
    pred, score = _classify_page(boxes)

    out = {
        "file": resolved,
        "predicted_class": pred,
        "is_table": pred == "table",
        "score_by_group": score,
        "boxes_count": len(boxes),
    }
    _CLASSIFY_CACHE[resolved] = out
    return out


def _normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _int(value: Any, default: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _cell_text(cell: Dict[str, Any]) -> str:
    direct = cell.get("text")
    if isinstance(direct, str) and direct.strip():
        return _normalize_text(direct)

    lines = cell.get("text_lines")
    if not isinstance(lines, list):
        return ""

    parts: List[str] = []
    for line in lines:
        if isinstance(line, str):
            token = _normalize_text(line)
            if token:
                parts.append(token)
            continue
        if isinstance(line, dict):
            token = line.get("text")
            if isinstance(token, str):
                token = _normalize_text(token)
                if token:
                    parts.append(token)
    return "\n".join(parts)


def _to_bbox(item: Dict[str, Any]) -> Optional[List[float]]:
    bbox = item.get("bbox")
    if isinstance(bbox, list) and len(bbox) == 4:
        return [float(bbox[0]), float(bbox[1]), float(bbox[2]), float(bbox[3])]

    polygon = item.get("polygon")
    if isinstance(polygon, list) and len(polygon) >= 4:
        xs: List[float] = []
        ys: List[float] = []
        for pt in polygon:
            if isinstance(pt, list) and len(pt) == 2:
                xs.append(float(pt[0]))
                ys.append(float(pt[1]))
        if xs and ys:
            return [min(xs), min(ys), max(xs), max(ys)]
    return None


def _bbox_area(b: List[float]) -> float:
    return max(0.0, b[2] - b[0]) * max(0.0, b[3] - b[1])


def _intersection_area(a: List[float], b: List[float]) -> float:
    x1 = max(a[0], b[0])
    y1 = max(a[1], b[1])
    x2 = min(a[2], b[2])
    y2 = min(a[3], b[3])
    return max(0.0, x2 - x1) * max(0.0, y2 - y1)


def _contains_point(b: List[float], x: float, y: float) -> bool:
    return b[0] <= x <= b[2] and b[1] <= y <= b[3]


def _line_text(line: Dict[str, Any]) -> str:
    text = line.get("text")
    if not isinstance(text, str):
        return ""
    return _normalize_text(text)


def _inject_ocr_text_into_table(
    pred_table: Dict[str, Any],
    ocr_result: Dict[str, Any],
    min_overlap_ratio: float = 0.2,
) -> None:
    cells = pred_table.get("cells")
    if not isinstance(cells, list):
        return

    text_lines = ocr_result.get("text_lines")
    if not isinstance(text_lines, list):
        return

    cell_infos: List[Dict[str, Any]] = []
    for idx, cell in enumerate(cells):
        if not isinstance(cell, dict):
            continue
        bbox = _to_bbox(cell)
        if bbox is None:
            continue
        cell_infos.append({"idx": idx, "bbox": bbox, "lines": []})
    if not cell_infos:
        return

    line_infos: List[Dict[str, Any]] = []
    for line in text_lines:
        if not isinstance(line, dict):
            continue
        text = _line_text(line)
        if not text:
            continue
        bbox = _to_bbox(line)
        if bbox is None:
            continue
        line_infos.append({"text": text, "bbox": bbox})

    for line in line_infos:
        line_bbox = line["bbox"]
        line_area = _bbox_area(line_bbox)
        if line_area <= 0:
            continue

        cx = (line_bbox[0] + line_bbox[2]) / 2.0
        cy = (line_bbox[1] + line_bbox[3]) / 2.0
        best_cell: Optional[Dict[str, Any]] = None
        best_score = 0.0

        for cell_info in cell_infos:
            cell_bbox = cell_info["bbox"]
            overlap = _intersection_area(cell_bbox, line_bbox)
            overlap_ratio = overlap / line_area
            center_inside = _contains_point(cell_bbox, cx, cy)
            score = overlap_ratio + (0.5 if center_inside else 0.0)
            if score > best_score:
                best_score = score
                best_cell = cell_info

        if best_cell is None:
            continue
        if best_score < min_overlap_ratio and not _contains_point(best_cell["bbox"], cx, cy):
            continue
        best_cell["lines"].append(line)

    for cell_info in cell_infos:
        cell = cells[cell_info["idx"]]
        assigned = cell_info["lines"]
        assigned.sort(key=lambda it: (it["bbox"][1], it["bbox"][0]))
        cell["text_lines"] = [{"text": it["text"], "bbox": it["bbox"]} for it in assigned]
        cell["text"] = "\n".join(it["text"] for it in assigned)


def _column_widths(pred_table: Dict[str, Any], n_cols: int) -> List[int]:
    cols = pred_table.get("cols")
    if not isinstance(cols, list):
        return [0] * n_cols

    widths = [0] * n_cols
    for col in cols:
        if not isinstance(col, dict):
            continue
        col_id = _int(col.get("col_id"), -1)
        if not (0 <= col_id < n_cols):
            continue
        bbox = col.get("bbox")
        if isinstance(bbox, list) and len(bbox) == 4:
            width = int(round(float(bbox[2]) - float(bbox[0])))
            widths[col_id] = max(0, width)
    return widths


def _surya_table_to_parsed(pred_table: Dict[str, Any], source: str) -> Dict[str, Any]:
    cells = pred_table.get("cells")
    if not isinstance(cells, list):
        raise ValueError("invalid surya table: 'cells' is missing")

    normalized_cells: List[Dict[str, Any]] = []
    max_row = 0
    max_col = 0
    for raw in cells:
        if not isinstance(raw, dict):
            continue
        row = _int(raw.get("row_id"), 0)
        col = _int(raw.get("col_id"), 0)
        rowspan = max(1, _int(raw.get("rowspan"), 1))
        colspan = max(1, _int(raw.get("colspan"), 1))
        normalized_cells.append(
            {
                "row": row,
                "col": col,
                "rowspan": rowspan,
                "colspan": colspan,
                "text": _cell_text(raw),
            }
        )
        max_row = max(max_row, row + rowspan)
        max_col = max(max_col, col + colspan)

    n_rows = max_row
    n_cols = max_col
    if n_rows <= 0 or n_cols <= 0:
        return {
            "source": source,
            "n_rows": 0,
            "n_cols": 0,
            "column_widths": [],
            "rows": [],
            "origin_cells": [],
        }

    grid: List[List[Optional[Dict[str, Any]]]] = [[None for _ in range(n_cols)] for _ in range(n_rows)]
    origins: List[Dict[str, Any]] = []

    normalized_cells.sort(key=lambda c: (c["row"], c["col"]))
    for cell in normalized_cells:
        row = cell["row"]
        col = cell["col"]
        rowspan = cell["rowspan"]
        colspan = cell["colspan"]
        text = cell["text"]

        if grid[row][col] is None:
            origin = {
                "type": "origin",
                "origin": [row, col],
                "text": text,
                "rowspan": rowspan,
                "colspan": colspan,
            }
            grid[row][col] = origin
            origins.append(
                {
                    "row": row,
                    "col": col,
                    "text": text,
                    "rowspan": rowspan,
                    "colspan": colspan,
                }
            )

        for dr in range(rowspan):
            for dc in range(colspan):
                if dr == 0 and dc == 0:
                    continue
                rr = row + dr
                cc = col + dc
                if rr >= n_rows or cc >= n_cols:
                    continue
                if grid[rr][cc] is not None:
                    continue
                ctype = "hMerge" if dr == 0 else "vMerge"
                grid[rr][cc] = {
                    "type": ctype,
                    "origin": [row, col],
                    "text": "",
                }

    rows_out: List[List[Dict[str, Any]]] = []
    for row in grid:
        rows_out.append([cell if cell is not None else {"type": "empty", "text": ""} for cell in row])

    return {
        "source": source,
        "n_rows": n_rows,
        "n_cols": n_cols,
        "column_widths": _column_widths(pred_table, n_cols),
        "rows": rows_out,
        "origin_cells": origins,
    }


def _load_surya_models() -> Tuple[Any, Any, Any, Any]:
    global _SURYA_MODELS
    if _SURYA_MODELS is None:
        from surya.common.surya.schema import TaskNames  # type: ignore
        from surya.detection import DetectionPredictor  # type: ignore
        from surya.foundation import FoundationPredictor  # type: ignore
        from surya.recognition import RecognitionPredictor  # type: ignore
        from surya.table_rec import TableRecPredictor  # type: ignore

        table_predictor = TableRecPredictor()
        foundation_predictor = FoundationPredictor()
        det_predictor = DetectionPredictor()
        rec_predictor = RecognitionPredictor(foundation_predictor)
        _SURYA_MODELS = (table_predictor, det_predictor, rec_predictor, TaskNames)
    return _SURYA_MODELS


def _prepare_image_for_surya(image_path: Path, bg_gray: int = _DEFAULT_BG_GRAY) -> Any:
    from PIL import Image  # type: ignore

    img = Image.open(image_path)
    has_alpha = (
        "A" in img.getbands()
        or img.mode in {"RGBA", "LA"}
        or (img.mode == "P" and "transparency" in img.info)
    )
    if not has_alpha:
        return img.convert("RGB")

    rgba = img.convert("RGBA")
    bg_gray = max(0, min(255, int(bg_gray)))
    bg = Image.new("RGBA", rgba.size, (bg_gray, bg_gray, bg_gray, 255))
    composited = Image.alpha_composite(bg, rgba)
    return composited.convert("RGB")


def _run_surya_parsed_tables(image_path: Path) -> List[Dict[str, Any]]:
    table_predictor, det_predictor, rec_predictor, task_names = _load_surya_models()

    image = _prepare_image_for_surya(image_path)
    table_preds = table_predictor([image])
    ocr_preds = rec_predictor(
        [image],
        task_names=[task_names.ocr_with_boxes],
        det_predictor=det_predictor,
        highres_images=[image],
        math_mode=False,
    )

    raw_tables = _to_builtin(table_preds)
    raw_ocr = _to_builtin(ocr_preds)
    if isinstance(raw_tables, dict):
        raw_tables = [raw_tables]
    if isinstance(raw_ocr, dict):
        raw_ocr = [raw_ocr]

    if isinstance(raw_tables, list):
        for t_idx, pred_table in enumerate(raw_tables):
            if not isinstance(pred_table, dict):
                continue
            ocr_page = raw_ocr[t_idx] if isinstance(raw_ocr, list) and t_idx < len(raw_ocr) else {}
            if isinstance(ocr_page, dict):
                _inject_ocr_text_into_table(pred_table, ocr_page)

    out: List[Dict[str, Any]] = []
    if isinstance(raw_tables, list):
        for i, item in enumerate(raw_tables, start=1):
            if not isinstance(item, dict):
                continue
            out.append(_surya_table_to_parsed(item, source=f"{image_path.resolve()}#table{i}"))
    return out


def _render_markdown_from_parsed_tables(parsed_tables: Sequence[Dict[str, Any]], header_rows: int = 1) -> str:
    try:
        import tableMaker  # type: ignore
    except Exception:
        from table_parser import tableMaker  # type: ignore

    blocks: List[str] = []
    for table in parsed_tables:
        dense = tableMaker._dense_grid_from_parsed_table(table, fill_merged=tableMaker.FILL_BOTH)
        md = tableMaker._render_markdown_flat(
            dense=dense,
            header_rows=max(0, int(header_rows)),
            use_header_rows=header_rows > 0,
        )
        if md.strip():
            blocks.append(md.strip())
    if not blocks:
        return ""
    return "\n\n".join(blocks).rstrip() + "\n"


def _table_quality_ok(parsed_tables: Sequence[Dict[str, Any]]) -> bool:
    for table in parsed_tables:
        n_rows = _int(table.get("n_rows"), 0)
        n_cols = _int(table.get("n_cols"), 0)
        if n_rows < 2 or n_cols < 2:
            continue

        origins = table.get("origin_cells")
        if not isinstance(origins, list):
            continue

        non_empty_total = 0
        non_empty_body = 0
        for cell in origins:
            if not isinstance(cell, dict):
                continue
            row = _int(cell.get("row"), 0)
            text = str(cell.get("text", "")).strip()
            if text:
                non_empty_total += 1
                if row > 0:
                    non_empty_body += 1

        # Require at least one non-empty body cell to avoid "header-only empty tables".
        if non_empty_body >= 1:
            return True
        # Allow single-row-like cases only when text density is clearly meaningful.
        if n_rows <= 2 and non_empty_total >= max(2, n_cols - 1):
            return True
    return False


def _should_try_surya_fallback(classification: Dict[str, Any], image_path: Path) -> bool:
    score = classification.get("score_by_group")
    if not isinstance(score, dict):
        score = {}
    table_score = float(score.get("table", 0.0) or 0.0)
    chart_score = float(score.get("chart", 0.0) or 0.0)
    text_score = float(score.get("text", 0.0) or 0.0)
    other_score = float(score.get("other", 0.0) or 0.0)
    max_non_table = max(chart_score, text_score, other_score, 0.0)

    # If table signal exists at all, try once.
    if table_score > 0:
        return True

    # Wide images are often pasted tables/screenshots.
    try:
        from PIL import Image  # type: ignore

        with Image.open(image_path) as im:
            w, h = im.size
    except Exception:
        w, h = 0, 0

    if h > 0 and (w / h) >= 2.2:
        return True

    # Near-ambiguous layout score.
    if max_non_table > 0 and table_score >= max_non_table * 0.25:
        return True
    return False


def extract_table_markdown_from_image(image_path: Path, header_rows: int = 1) -> Dict[str, Any]:
    """Return a status dict for one image.

    Status:
      - {"status": "table", "markdown": "...", ...}
      - {"status": "not_table", ...}
      - {"status": "error", "error": "...", ...}
    """
    resolved_path = image_path.expanduser().resolve()
    key = str(resolved_path)
    cached = _RESULT_CACHE.get(key)
    if cached is not None:
        return cached

    if not resolved_path.exists() or not resolved_path.is_file():
        out = {"status": "error", "file": key, "error": f"file not found: {resolved_path}"}
        _RESULT_CACHE[key] = out
        return out

    suffix = resolved_path.suffix.lower()
    if suffix not in _SUPPORTED_LAYOUT_SUFFIXES:
        out = {
            "status": "not_table",
            "file": key,
            "reason": f"unsupported image extension for layout model: {suffix or '(none)'}",
        }
        _RESULT_CACHE[key] = out
        return out

    try:
        cls = classify_image(resolved_path)
    except Exception as exc:
        out = {
            "status": "error",
            "file": key,
            "error": f"paddle layout classification failed: {type(exc).__name__}: {exc}",
        }
        _RESULT_CACHE[key] = out
        return out

    if not cls.get("is_table"):
        if _should_try_surya_fallback(cls, resolved_path):
            try:
                parsed_tables = _run_surya_parsed_tables(resolved_path)
            except Exception:
                parsed_tables = []

            if parsed_tables and _table_quality_ok(parsed_tables):
                markdown = _render_markdown_from_parsed_tables(parsed_tables, header_rows=header_rows)
                if markdown.strip():
                    out = {
                        "status": "table",
                        "file": key,
                        "predicted_class": cls.get("predicted_class"),
                        "score_by_group": cls.get("score_by_group"),
                        "boxes_count": cls.get("boxes_count"),
                        "table_count": len(parsed_tables),
                        "markdown": markdown,
                        "used_surya_fallback": True,
                    }
                    _RESULT_CACHE[key] = out
                    return out

        out = {
            "status": "not_table",
            "file": key,
            "predicted_class": cls.get("predicted_class"),
            "score_by_group": cls.get("score_by_group"),
            "boxes_count": cls.get("boxes_count"),
        }
        _RESULT_CACHE[key] = out
        return out

    try:
        parsed_tables = _run_surya_parsed_tables(resolved_path)
    except Exception as exc:
        out = {
            "status": "error",
            "file": key,
            "predicted_class": cls.get("predicted_class"),
            "score_by_group": cls.get("score_by_group"),
            "error": f"surya table extraction failed: {type(exc).__name__}: {exc}",
        }
        _RESULT_CACHE[key] = out
        return out

    if not parsed_tables:
        out = {
            "status": "error",
            "file": key,
            "predicted_class": cls.get("predicted_class"),
            "score_by_group": cls.get("score_by_group"),
            "error": "surya did not produce any table payload",
        }
        _RESULT_CACHE[key] = out
        return out

    if not _table_quality_ok(parsed_tables):
        out = {
            "status": "not_table",
            "file": key,
            "predicted_class": cls.get("predicted_class"),
            "score_by_group": cls.get("score_by_group"),
            "boxes_count": cls.get("boxes_count"),
            "reason": "low_text_density_after_ocr",
        }
        _RESULT_CACHE[key] = out
        return out

    markdown = _render_markdown_from_parsed_tables(parsed_tables, header_rows=header_rows)
    if not markdown.strip():
        out = {
            "status": "error",
            "file": key,
            "predicted_class": cls.get("predicted_class"),
            "score_by_group": cls.get("score_by_group"),
            "error": "table markdown rendering returned empty output",
        }
        _RESULT_CACHE[key] = out
        return out

    out = {
        "status": "table",
        "file": key,
        "predicted_class": cls.get("predicted_class"),
        "score_by_group": cls.get("score_by_group"),
        "boxes_count": cls.get("boxes_count"),
        "table_count": len(parsed_tables),
        "markdown": markdown,
    }
    _RESULT_CACHE[key] = out
    return out
