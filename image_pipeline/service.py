#!/usr/bin/env python3
"""Surya bbox-based table classifier + markdown extraction.

This module treats image-table detection as a binary classification problem:
  - ``table``
  - ``non-table``

Classification is based on Surya table-recognition outputs instead of OpenCV
heuristics. The main signals are:
  - detected table union area / visible image area
  - IoU(detected table union bbox, visible image bbox)
  - normalized center offset between the two boxes
  - detected row / col / cell counts

The final decision uses a double-threshold policy:
  - high threshold: accept as table
  - low threshold: reject as non-table
  - mid band: accept only when structure/geometric gates both pass
"""

from __future__ import annotations

import contextlib
import io
import os
import re
import shutil
import subprocess
import tempfile
import warnings
from pathlib import Path
from typing import Any, Dict, Iterator, List, Optional, Sequence, Tuple


TABLE_HIGH_THRESHOLD = 0.64
TABLE_LOW_THRESHOLD = 0.36
MID_BAND_STRUCTURE_THRESHOLD = 0.58
MID_BAND_GEOMETRY_THRESHOLD = 0.24
MIN_ROW_COUNT = 2
MIN_COL_COUNT = 2
MIN_CELL_COUNT = 4
MID_BAND_MIN_CELL_COUNT = 6

_SURYA_MODELS: Optional[Tuple[Any, Any, Any, Any]] = None
_CLASSIFY_CACHE: Dict[str, Dict[str, Any]] = {}
_RESULT_CACHE: Dict[str, Dict[str, Any]] = {}
_SURYA_TABLE_CACHE: Dict[str, Dict[str, Any]] = {}
_SUPPORTED_IMAGE_SUFFIXES = {
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".bmp",
    ".webp",
    ".tif",
    ".tiff",
}
_VECTOR_IMAGE_SUFFIXES = {
    ".emf",
    ".wmf",
}
_DEFAULT_BG_GRAY = 192
_HIDE_SURYA_LOGS = os.getenv("IMAGE_TABLE_HIDE_SURYA_LOGS", "").strip().lower() in {"1", "true", "yes", "on"}
_DISABLE_SURYA = os.getenv("IMAGE_TABLE_DISABLE_SURYA", "").strip().lower() in {"1", "true", "yes", "on"}

warnings.filterwarnings(
    "ignore",
    message=r"Using `TRANSFORMERS_CACHE` is deprecated and will be removed in v5 of Transformers\..*",
    category=FutureWarning,
)


def _clamp(value: float, lower: float = 0.0, upper: float = 1.0) -> float:
    return max(lower, min(upper, float(value)))


def _safe_ratio(numerator: float, denominator: float) -> float:
    if denominator <= 0:
        return 0.0
    return float(numerator) / float(denominator)


def _normalize_score(value: float, full_score_at: float) -> float:
    if full_score_at <= 0:
        return 0.0
    return _clamp(value / full_score_at)


def _is_supported_image_suffix(suffix: str) -> bool:
    return suffix in _SUPPORTED_IMAGE_SUFFIXES or suffix in _VECTOR_IMAGE_SUFFIXES


def _resolve_vector_converter() -> Optional[str]:
    for candidate in ("soffice", "libreoffice"):
        found = shutil.which(candidate)
        if found:
            return found
    return None


def _rasterize_vector_image(image_path: Path) -> Path:
    converter = _resolve_vector_converter()
    if not converter:
        raise RuntimeError("LibreOffice is required to rasterize vector images such as EMF/WMF")

    with tempfile.TemporaryDirectory(prefix="vector_raster_") as tmpdir:
        tmp_root = Path(tmpdir)
        staged_input = tmp_root / image_path.name
        shutil.copy2(image_path, staged_input)
        proc = subprocess.run(
            [
                converter,
                "--headless",
                "--convert-to",
                "png",
                "--outdir",
                str(tmp_root),
                str(staged_input),
            ],
            text=True,
            capture_output=True,
            check=False,
        )
        if proc.returncode != 0:
            raise RuntimeError(
                "vector image rasterization failed\n"
                f"stdout={proc.stdout.strip()}\n"
                f"stderr={proc.stderr.strip()}"
            )

        output_path = tmp_root / f"{image_path.stem}.png"
        if not output_path.exists():
            pngs = sorted(tmp_root.glob("*.png"))
            if not pngs:
                raise RuntimeError(f"vector image rasterization produced no PNG output: {image_path.name}")
            output_path = pngs[0]
        persisted_output = Path(tempfile.mkdtemp(prefix="vector_raster_png_")) / output_path.name
        shutil.copy2(output_path, persisted_output)
        return persisted_output


def _visible_region_from_rgba(rgba: Any) -> Tuple[List[float], float]:
    alpha = rgba.getchannel("A")
    bbox = alpha.getbbox()
    histogram = alpha.histogram()
    visible_area = float(sum(histogram[1:])) if histogram else 0.0
    width, height = rgba.size
    if bbox is None or visible_area <= 0:
        return [0.0, 0.0, float(width), float(height)], float(width * height)
    return [float(bbox[0]), float(bbox[1]), float(bbox[2]), float(bbox[3])], visible_area


def _load_image_bundle(image_path: Path, bg_gray: int = _DEFAULT_BG_GRAY) -> Dict[str, Any]:
    from PIL import Image, ImageOps  # type: ignore

    suffix = image_path.suffix.lower()
    if not _is_supported_image_suffix(suffix):
        raise ValueError(f"unsupported image extension: {suffix or '(none)'}")

    bg_gray = max(0, min(255, int(bg_gray)))
    raster_path = image_path
    cleanup_dir: Optional[Path] = None

    if suffix in _VECTOR_IMAGE_SUFFIXES:
        raster_path = _rasterize_vector_image(image_path)
        cleanup_dir = raster_path.parent

    try:
        with Image.open(raster_path) as loaded:
            image = ImageOps.exif_transpose(loaded)
            try:
                image.seek(0)
            except Exception:
                pass

            width, height = image.size
            canvas_bbox = [0.0, 0.0, float(width), float(height)]
            canvas_area = float(width * height)

            if image.mode in {"RGBA", "LA"} or "transparency" in image.info:
                rgba = image.convert("RGBA")
                visible_bbox, visible_area = _visible_region_from_rgba(rgba)
                bg = Image.new("RGBA", rgba.size, (bg_gray, bg_gray, bg_gray, 255))
                composited = Image.alpha_composite(bg, rgba).convert("RGB")
                return {
                    "image": composited,
                    "image_size": {"width": width, "height": height},
                    "visible_image": {
                        "bbox": visible_bbox,
                        "area": visible_area,
                        "canvas_bbox": canvas_bbox,
                        "canvas_area": canvas_area,
                        "alpha_cropped": visible_bbox != canvas_bbox,
                    },
                }

            return {
                "image": image.convert("RGB"),
                "image_size": {"width": width, "height": height},
                "visible_image": {
                    "bbox": canvas_bbox,
                    "area": canvas_area,
                    "canvas_bbox": canvas_bbox,
                    "canvas_area": canvas_area,
                    "alpha_cropped": False,
                },
            }
    finally:
        if cleanup_dir is not None:
            shutil.rmtree(cleanup_dir, ignore_errors=True)


@contextlib.contextmanager
def _suppress_external_output() -> Iterator[None]:
    sink_out = io.StringIO()
    sink_err = io.StringIO()
    with contextlib.redirect_stdout(sink_out), contextlib.redirect_stderr(sink_err):
        yield


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


def _normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _int(value: Any, default: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _to_bbox(item: Dict[str, Any]) -> Optional[List[float]]:
    bbox = item.get("bbox")
    if isinstance(bbox, list) and len(bbox) == 4:
        return [
            float(min(bbox[0], bbox[2])),
            float(min(bbox[1], bbox[3])),
            float(max(bbox[0], bbox[2])),
            float(max(bbox[1], bbox[3])),
        ]

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


def _bbox_area(bbox: Optional[List[float]]) -> float:
    if bbox is None:
        return 0.0
    return max(0.0, bbox[2] - bbox[0]) * max(0.0, bbox[3] - bbox[1])


def _merge_bboxes(bboxes: Sequence[List[float]]) -> Optional[List[float]]:
    if not bboxes:
        return None
    xs1 = [bbox[0] for bbox in bboxes]
    ys1 = [bbox[1] for bbox in bboxes]
    xs2 = [bbox[2] for bbox in bboxes]
    ys2 = [bbox[3] for bbox in bboxes]
    return [min(xs1), min(ys1), max(xs2), max(ys2)]


def _intersection_area(a: Optional[List[float]], b: Optional[List[float]]) -> float:
    if a is None or b is None:
        return 0.0
    x1 = max(a[0], b[0])
    y1 = max(a[1], b[1])
    x2 = min(a[2], b[2])
    y2 = min(a[3], b[3])
    return max(0.0, x2 - x1) * max(0.0, y2 - y1)


def _bbox_iou(a: Optional[List[float]], b: Optional[List[float]]) -> float:
    inter = _intersection_area(a, b)
    if inter <= 0:
        return 0.0
    union = _bbox_area(a) + _bbox_area(b) - inter
    return _safe_ratio(inter, union)


def _bbox_center(bbox: Optional[List[float]]) -> Optional[Tuple[float, float]]:
    if bbox is None:
        return None
    return ((bbox[0] + bbox[2]) / 2.0, (bbox[1] + bbox[3]) / 2.0)


def _normalized_center_offset(a: Optional[List[float]], b: Optional[List[float]]) -> float:
    center_a = _bbox_center(a)
    center_b = _bbox_center(b)
    if center_a is None or center_b is None or b is None:
        return 1.0
    ref_w = max(1.0, b[2] - b[0])
    ref_h = max(1.0, b[3] - b[1])
    ref_diag = (ref_w**2 + ref_h**2) ** 0.5
    distance = ((center_a[0] - center_b[0]) ** 2 + (center_a[1] - center_b[1]) ** 2) ** 0.5
    return _clamp(distance / ref_diag)


def _rectangles_union_area(rectangles: Sequence[List[float]]) -> float:
    events: List[Tuple[float, int, float, float]] = []
    for rect in rectangles:
        if len(rect) != 4:
            continue
        x1, y1, x2, y2 = rect
        if x2 <= x1 or y2 <= y1:
            continue
        events.append((x1, 1, y1, y2))
        events.append((x2, -1, y1, y2))
    if not events:
        return 0.0

    events.sort(key=lambda item: item[0])
    area = 0.0
    active: List[Tuple[float, float]] = []
    prev_x = events[0][0]
    idx = 0

    while idx < len(events):
        x = events[idx][0]
        dx = x - prev_x
        if dx > 0 and active:
            merged_y = 0.0
            ordered = sorted(active)
            start, end = ordered[0]
            for cur_start, cur_end in ordered[1:]:
                if cur_start <= end:
                    end = max(end, cur_end)
                else:
                    merged_y += end - start
                    start, end = cur_start, cur_end
            merged_y += end - start
            area += dx * merged_y

        while idx < len(events) and events[idx][0] == x:
            _, event_type, y1, y2 = events[idx]
            interval = (y1, y2)
            if event_type == 1:
                active.append(interval)
            else:
                try:
                    active.remove(interval)
                except ValueError:
                    pass
            idx += 1
        prev_x = x

    return area


def _count_valid_boxes(items: Any) -> int:
    if not isinstance(items, list):
        return 0
    count = 0
    for item in items:
        if isinstance(item, dict) and _to_bbox(item) is not None:
            count += 1
    return count


def _table_bbox_from_prediction(pred_table: Dict[str, Any]) -> Optional[List[float]]:
    bbox = _to_bbox(pred_table)
    if bbox is not None:
        return bbox

    bboxes: List[List[float]] = []
    for key in ("cells", "rows", "cols"):
        items = pred_table.get(key)
        if not isinstance(items, list):
            continue
        for item in items:
            if not isinstance(item, dict):
                continue
            item_bbox = _to_bbox(item)
            if item_bbox is not None:
                bboxes.append(item_bbox)
    return _merge_bboxes(bboxes)


def _summarize_table_predictions(
    raw_tables: Sequence[Dict[str, Any]],
    visible_bbox: List[float],
    visible_area: float,
) -> Dict[str, Any]:
    table_summaries: List[Dict[str, Any]] = []
    table_bboxes: List[List[float]] = []
    total_rows = 0
    total_cols = 0
    total_cells = 0
    max_rows = 0
    max_cols = 0
    max_cells = 0

    for idx, pred_table in enumerate(raw_tables, start=1):
        if not isinstance(pred_table, dict):
            continue
        bbox = _table_bbox_from_prediction(pred_table)
        row_count = _count_valid_boxes(pred_table.get("rows"))
        col_count = _count_valid_boxes(pred_table.get("cols"))
        cell_count = _count_valid_boxes(pred_table.get("cells"))
        if bbox is None and row_count == 0 and col_count == 0 and cell_count == 0:
            continue

        if bbox is not None:
            table_bboxes.append(bbox)

        total_rows += row_count
        total_cols += col_count
        total_cells += cell_count
        max_rows = max(max_rows, row_count)
        max_cols = max(max_cols, col_count)
        max_cells = max(max_cells, cell_count)

        table_summaries.append(
            {
                "table_idx": idx,
                "bbox": bbox,
                "bbox_area": _bbox_area(bbox),
                "bbox_area_ratio": _safe_ratio(_bbox_area(bbox), visible_area),
                "row_count": row_count,
                "col_count": col_count,
                "cell_count": cell_count,
                "box_count": row_count + col_count + cell_count,
            }
        )

    union_bbox = _merge_bboxes(table_bboxes)
    union_area = _rectangles_union_area(table_bboxes)
    union_bbox_area = _bbox_area(union_bbox)
    area_ratio = _safe_ratio(union_area, visible_area)
    union_bbox_area_ratio = _safe_ratio(union_bbox_area, visible_area)
    bbox_iou = _bbox_iou(union_bbox, visible_bbox)
    center_offset = _normalized_center_offset(union_bbox, visible_bbox)

    area_score = _normalize_score(area_ratio, 0.18)
    bbox_area_score = _normalize_score(union_bbox_area_ratio, 0.22)
    iou_score = _normalize_score(bbox_iou, 0.75)
    center_score = 1.0 - _clamp(center_offset / 0.45)
    row_score = _normalize_score(max_rows, 4.0)
    col_score = _normalize_score(max_cols, 3.0)
    cell_score = _normalize_score(max(total_cells, max_cells), 12.0)
    box_count = total_rows + total_cols + total_cells
    box_count_score = _normalize_score(box_count, 18.0)
    table_count_score = _normalize_score(len(table_summaries), 2.0)

    geometry_score = _clamp(
        (area_score * 0.34)
        + (bbox_area_score * 0.16)
        + (iou_score * 0.30)
        + (center_score * 0.20)
    )
    structure_score = _clamp(
        (row_score * 0.18)
        + (col_score * 0.16)
        + (cell_score * 0.38)
        + (box_count_score * 0.18)
        + (table_count_score * 0.10)
    )
    final_score = _clamp((geometry_score * 0.55) + (structure_score * 0.45))

    has_detection = bool(table_summaries and union_bbox is not None)
    has_min_structure = max_rows >= MIN_ROW_COUNT and max_cols >= MIN_COL_COUNT and total_cells >= MIN_CELL_COUNT
    mid_band_structure_ok = (
        structure_score >= MID_BAND_STRUCTURE_THRESHOLD
        and max_rows >= MIN_ROW_COUNT
        and max_cols >= MIN_COL_COUNT
        and total_cells >= MID_BAND_MIN_CELL_COUNT
    )
    mid_band_geometry_ok = (
        geometry_score >= MID_BAND_GEOMETRY_THRESHOLD
        and center_offset <= 0.38
        and (area_ratio >= 0.04 or union_bbox_area_ratio >= 0.06 or bbox_iou >= 0.16)
    )
    high_confidence_table = has_min_structure and final_score >= TABLE_HIGH_THRESHOLD
    mid_band = TABLE_LOW_THRESHOLD <= final_score < TABLE_HIGH_THRESHOLD
    mid_band_accept = mid_band and has_detection and mid_band_structure_ok and mid_band_geometry_ok
    predicted_table = high_confidence_table or mid_band_accept
    predicted_class = "table" if predicted_table else "non-table"
    low_confidence = mid_band

    return {
        "predicted_class": predicted_class,
        "score": final_score,
        "low_confidence": low_confidence,
        "feature_values": {
            "detected_table_count": len(table_summaries),
            "detected_row_count": total_rows,
            "detected_col_count": total_cols,
            "detected_cell_count": total_cells,
            "detected_box_count": box_count,
            "max_row_count": max_rows,
            "max_col_count": max_cols,
            "max_cell_count": max_cells,
            "detected_table_union_area": union_area,
            "detected_table_union_bbox_area": union_bbox_area,
            "visible_image_area": visible_area,
            "table_union_area_ratio": area_ratio,
            "table_union_bbox_area_ratio": union_bbox_area_ratio,
            "table_union_bbox_iou": bbox_iou,
            "table_union_center_offset": center_offset,
        },
        "feature_details": {
            "visible_image_bbox": visible_bbox,
            "detected_table_union_bbox": union_bbox,
            "detected_tables": table_summaries,
        },
        "score_breakdown": {
            "geometry_score": geometry_score,
            "structure_score": structure_score,
            "area_score": area_score,
            "bbox_area_score": bbox_area_score,
            "iou_score": iou_score,
            "center_score": center_score,
            "row_score": row_score,
            "col_score": col_score,
            "cell_score": cell_score,
            "box_count_score": box_count_score,
            "table_count_score": table_count_score,
            "final_score": final_score,
        },
        "decision_flags": {
            "has_detection": has_detection,
            "has_min_structure": has_min_structure,
            "mid_band": mid_band,
            "mid_band_structure_ok": mid_band_structure_ok,
            "mid_band_geometry_ok": mid_band_geometry_ok,
            "mid_band_accept": mid_band_accept,
            "high_confidence_table": high_confidence_table,
        },
    }


def _compute_features(surya_bundle: Dict[str, Any]) -> Dict[str, Any]:
    visible_image = surya_bundle.get("visible_image")
    if not isinstance(visible_image, dict):
        raise ValueError("surya bundle missing visible_image metadata")

    visible_bbox = visible_image.get("bbox")
    visible_area = float(visible_image.get("area", 0.0))
    if not isinstance(visible_bbox, list) or len(visible_bbox) != 4:
        raise ValueError("surya bundle missing visible image bbox")

    raw_tables = surya_bundle.get("raw_tables")
    if not isinstance(raw_tables, list):
        raw_tables = []

    summary = _summarize_table_predictions(raw_tables, visible_bbox=visible_bbox, visible_area=visible_area)
    return {
        "image_size": surya_bundle.get("image_size"),
        "visible_image": visible_image,
        "feature_values": summary["feature_values"],
        "feature_details": summary["feature_details"],
        "score_breakdown": summary["score_breakdown"],
        "decision_flags": summary["decision_flags"],
    }


def classify_image(image_path: Path) -> Dict[str, Any]:
    resolved_path = image_path.expanduser().resolve()
    key = str(resolved_path)
    cached = _CLASSIFY_CACHE.get(key)
    if cached is not None:
        return cached

    if not resolved_path.exists() or not resolved_path.is_file():
        raise FileNotFoundError(f"file not found: {resolved_path}")

    suffix = resolved_path.suffix.lower()
    if not _is_supported_image_suffix(suffix):
        raise ValueError(f"unsupported image extension: {suffix or '(none)'}")
    if _DISABLE_SURYA:
        raise RuntimeError("surya table detection is disabled by IMAGE_TABLE_DISABLE_SURYA")

    surya_bundle = _get_surya_table_bundle(resolved_path)
    features = _compute_features(surya_bundle)
    score = float(features["score_breakdown"]["final_score"])
    flags = features["decision_flags"]
    predicted_class = "table" if bool(flags.get("high_confidence_table") or flags.get("mid_band_accept")) else "non-table"
    out = {
        "file": key,
        "predicted_class": predicted_class,
        "is_table": predicted_class == "table",
        "score": score,
        "low_confidence": TABLE_LOW_THRESHOLD <= score < TABLE_HIGH_THRESHOLD,
        "thresholds": {
            "table_high": TABLE_HIGH_THRESHOLD,
            "table_low": TABLE_LOW_THRESHOLD,
            "mid_band_structure": MID_BAND_STRUCTURE_THRESHOLD,
            "mid_band_geometry": MID_BAND_GEOMETRY_THRESHOLD,
        },
        **features,
    }
    _CLASSIFY_CACHE[key] = out
    return out


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
        bbox = _to_bbox(col)
        if bbox is not None:
            width = int(round(bbox[2] - bbox[0]))
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


def _get_surya_table_bundle(image_path: Path) -> Dict[str, Any]:
    resolved_path = image_path.expanduser().resolve()
    key = str(resolved_path)
    cached = _SURYA_TABLE_CACHE.get(key)
    if cached is not None:
        return cached
    if _DISABLE_SURYA:
        raise RuntimeError("surya table detection is disabled by IMAGE_TABLE_DISABLE_SURYA")

    image_bundle = _load_image_bundle(resolved_path)
    image = image_bundle["image"]

    run_ctx = _suppress_external_output() if _HIDE_SURYA_LOGS else contextlib.nullcontext()
    with run_ctx:
        table_predictor, det_predictor, rec_predictor, task_names = _load_surya_models()
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

    normalized_tables: List[Dict[str, Any]] = []
    if isinstance(raw_tables, list):
        for table_idx, pred_table in enumerate(raw_tables):
            if not isinstance(pred_table, dict):
                continue
            ocr_page = raw_ocr[table_idx] if isinstance(raw_ocr, list) and table_idx < len(raw_ocr) else {}
            if isinstance(ocr_page, dict):
                _inject_ocr_text_into_table(pred_table, ocr_page)
            normalized_tables.append(pred_table)

    parsed_tables: List[Dict[str, Any]] = []
    for idx, pred_table in enumerate(normalized_tables, start=1):
        parsed_tables.append(_surya_table_to_parsed(pred_table, source=f"{resolved_path}#table{idx}"))

    out = {
        "file": key,
        "image_size": image_bundle["image_size"],
        "visible_image": image_bundle["visible_image"],
        "raw_tables": normalized_tables,
        "parsed_tables": parsed_tables,
    }
    _SURYA_TABLE_CACHE[key] = out
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

        if non_empty_body >= 1:
            return True
        if n_rows <= 2 and non_empty_total >= max(2, n_cols - 1):
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
    if not _is_supported_image_suffix(suffix):
        out = {
            "status": "not_table",
            "file": key,
            "predicted_class": "non-table",
            "reason": f"unsupported image extension: {suffix or '(none)'}",
        }
        _RESULT_CACHE[key] = out
        return out

    if _DISABLE_SURYA:
        out = {
            "status": "table_skipped",
            "file": key,
            "predicted_class": None,
            "surya_attempted": False,
            "reason": "surya_disabled_by_env",
        }
        _RESULT_CACHE[key] = out
        return out

    try:
        cls = classify_image(resolved_path)
    except Exception as exc:  # noqa: BLE001
        out = {
            "status": "error",
            "file": key,
            "error": f"bbox-based table classification failed: {type(exc).__name__}: {exc}",
        }
        _RESULT_CACHE[key] = out
        return out

    if not cls.get("is_table"):
        out = {
            "status": "not_table",
            "file": key,
            "predicted_class": cls.get("predicted_class"),
            "score": cls.get("score"),
            "low_confidence": cls.get("low_confidence"),
            "surya_attempted": True,
            "classification": cls,
        }
        _RESULT_CACHE[key] = out
        return out

    try:
        surya_bundle = _get_surya_table_bundle(resolved_path)
    except Exception as exc:  # noqa: BLE001
        out = {
            "status": "error",
            "file": key,
            "predicted_class": cls.get("predicted_class"),
            "score": cls.get("score"),
            "low_confidence": cls.get("low_confidence"),
            "surya_attempted": True,
            "classification": cls,
            "error": f"surya table extraction failed: {type(exc).__name__}: {exc}",
        }
        _RESULT_CACHE[key] = out
        return out

    parsed_tables = surya_bundle.get("parsed_tables")
    if not isinstance(parsed_tables, list) or not parsed_tables:
        out = {
            "status": "error",
            "file": key,
            "predicted_class": cls.get("predicted_class"),
            "score": cls.get("score"),
            "low_confidence": cls.get("low_confidence"),
            "surya_attempted": True,
            "classification": cls,
            "error": "surya did not produce any table payload",
        }
        _RESULT_CACHE[key] = out
        return out

    markdown = _render_markdown_from_parsed_tables(parsed_tables, header_rows=header_rows)
    if not markdown.strip():
        out = {
            "status": "error",
            "file": key,
            "predicted_class": cls.get("predicted_class"),
            "score": cls.get("score"),
            "low_confidence": cls.get("low_confidence"),
            "surya_attempted": True,
            "classification": cls,
            "error": "table markdown rendering returned empty output",
        }
        _RESULT_CACHE[key] = out
        return out

    quality_ok = _table_quality_ok(parsed_tables)
    out = {
        "status": "table",
        "file": key,
        "predicted_class": cls.get("predicted_class"),
        "score": cls.get("score"),
        "low_confidence": cls.get("low_confidence"),
        "surya_attempted": True,
        "surya_quality_ok": quality_ok,
        "table_count": len(parsed_tables),
        "markdown": markdown,
        "classification": cls,
    }
    if not quality_ok:
        out["reason"] = "low_text_density_after_ocr"
    _RESULT_CACHE[key] = out
    return out
