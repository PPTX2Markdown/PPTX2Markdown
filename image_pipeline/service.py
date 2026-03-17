#!/usr/bin/env python3
"""OpenCV heuristic table classifier + Surya table extraction.

This module treats image-table detection as a binary classification problem:
  - ``table``
  - ``non-table``

Classification is done with OpenCV heuristics only. When an image is classified
as ``table``, Surya table recognition is used as the downstream extractor.
"""

from __future__ import annotations

import contextlib
import io
import math
import os
import re
import shutil
import subprocess
import tempfile
import warnings
from pathlib import Path
from typing import Any, Dict, Iterator, List, Optional, Sequence, Tuple


TABLE_THRESHOLD = 0.59
LOW_CONFIDENCE_THRESHOLD = 0.34
MAX_IMAGE_EDGE = 1600

_SURYA_MODELS: Optional[Tuple[Any, Any, Any, Any]] = None
_CLASSIFY_CACHE: Dict[str, Dict[str, Any]] = {}
_RESULT_CACHE: Dict[str, Dict[str, Any]] = {}
_SUPPORTED_IMAGE_SUFFIXES = {
    ".png",
    ".jpg",
    ".jpeg",
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


def _cluster_positions(values: Sequence[float], tolerance: float) -> List[List[float]]:
    if not values:
        return []
    tol = max(1.0, float(tolerance))
    ordered = sorted(float(v) for v in values)
    groups: List[List[float]] = [[ordered[0]]]
    for value in ordered[1:]:
        if abs(value - groups[-1][-1]) <= tol:
            groups[-1].append(value)
        else:
            groups.append([value])
    return groups


def _median(values: Sequence[float]) -> float:
    if not values:
        return 0.0
    ordered = sorted(float(v) for v in values)
    mid = len(ordered) // 2
    if len(ordered) % 2:
        return ordered[mid]
    return (ordered[mid - 1] + ordered[mid]) / 2.0


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


def _rasterize_vector_image_with_cv(image_path: Path, cv2: Any, imread_flag: int) -> Any:
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

        image = cv2.imread(str(output_path), imread_flag)
        if image is None:
            raise RuntimeError(f"failed to read rasterized vector image: {output_path}")
        return image


def _read_image_with_cv(image_path: Path, cv2: Any, imread_flag: int) -> Any:
    suffix = image_path.suffix.lower()
    if suffix in _SUPPORTED_IMAGE_SUFFIXES:
        image = cv2.imread(str(image_path), imread_flag)
        if image is None:
            raise ValueError(f"failed to read image: {image_path}")
        return image
    if suffix in _VECTOR_IMAGE_SUFFIXES:
        return _rasterize_vector_image_with_cv(image_path, cv2, imread_flag)
    raise ValueError(f"unsupported image extension: {suffix or '(none)'}")


@contextlib.contextmanager
def _suppress_external_output() -> Iterator[None]:
    sink_out = io.StringIO()
    sink_err = io.StringIO()
    with contextlib.redirect_stdout(sink_out), contextlib.redirect_stderr(sink_err):
        yield


def _resize_for_analysis(gray: Any, cv2: Any) -> Any:
    height, width = gray.shape[:2]
    longest = max(height, width)
    if longest <= MAX_IMAGE_EDGE:
        return gray
    scale = MAX_IMAGE_EDGE / float(longest)
    resized_w = max(1, int(round(width * scale)))
    resized_h = max(1, int(round(height * scale)))
    return cv2.resize(gray, (resized_w, resized_h), interpolation=cv2.INTER_AREA)


def _prepare_binary(gray: Any, cv2: Any) -> Tuple[Any, Any]:
    gray = _resize_for_analysis(gray, cv2)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    binary = cv2.adaptiveThreshold(
        blurred,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY_INV,
        31,
        11,
    )
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    binary = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel, iterations=1)
    binary = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, kernel, iterations=1)
    return gray, binary


def _extract_line_mask(binary: Any, axis: str, cv2: Any) -> Any:
    height, width = binary.shape[:2]
    if axis == "horizontal":
        kernel_len = max(12, width // 18)
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (kernel_len, 1))
    else:
        kernel_len = max(12, height // 18)
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, kernel_len))
    return cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel, iterations=1)


def _count_long_lines(mask: Any, axis: str, image_shape: Tuple[int, int], cv2: Any) -> int:
    height, width = image_shape
    contours, _ = cv2.findContours(mask, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    count = 0
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        if axis == "horizontal":
            if w >= width * 0.30 and h <= max(8, int(height * 0.04)):
                count += 1
        else:
            if h >= height * 0.30 and w <= max(8, int(width * 0.04)):
                count += 1
    return count


def _count_intersections(horizontal_mask: Any, vertical_mask: Any, cv2: Any) -> int:
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    h = cv2.dilate(horizontal_mask, kernel, iterations=1)
    v = cv2.dilate(vertical_mask, kernel, iterations=1)
    intersections = cv2.bitwise_and(h, v)
    count, _, stats, _ = cv2.connectedComponentsWithStats(intersections, connectivity=8)
    hits = 0
    for idx in range(1, count):
        area = int(stats[idx, cv2.CC_STAT_AREA])
        if area >= 4:
            hits += 1
    return hits


def _rectangularity(contour: Any, cv2: Any) -> float:
    area = float(cv2.contourArea(contour))
    if area <= 0:
        return 0.0
    x, y, w, h = cv2.boundingRect(contour)
    bbox_area = float(w * h)
    if bbox_area <= 0:
        return 0.0
    return _clamp(area / bbox_area)


def _collect_rectangles(combined_mask: Any, image_shape: Tuple[int, int], cv2: Any) -> List[Dict[str, float]]:
    height, width = image_shape
    work = cv2.dilate(combined_mask, cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3)), iterations=1)
    contours, _ = cv2.findContours(work, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    rectangles: List[Dict[str, float]] = []
    min_area = max(24.0, float(height * width) * 0.00008)
    max_area = float(height * width) * 0.45
    for contour in contours:
        area = float(cv2.contourArea(contour))
        if area < min_area or area > max_area:
            continue
        peri = cv2.arcLength(contour, True)
        approx = cv2.approxPolyDP(contour, 0.04 * peri, True)
        if len(approx) != 4 or not cv2.isContourConvex(approx):
            continue
        x, y, w, h = cv2.boundingRect(approx)
        if w < 8 or h < 8:
            continue
        rectangularity = _rectangularity(contour, cv2)
        if rectangularity < 0.60:
            continue
        rectangles.append(
            {
                "x": float(x),
                "y": float(y),
                "w": float(w),
                "h": float(h),
                "area": area,
                "rectangularity": rectangularity,
            }
        )
    rectangles.sort(key=lambda item: (item["y"], item["x"]))
    return rectangles


def _repeating_cell_structure(rectangles: Sequence[Dict[str, float]]) -> Tuple[float, Dict[str, Any]]:
    if len(rectangles) < 4:
        return 0.0, {"repeating_cells": False, "similar_rectangles": 0}

    widths = [rect["w"] for rect in rectangles]
    heights = [rect["h"] for rect in rectangles]
    median_w = _median(widths)
    median_h = _median(heights)
    if median_w <= 0 or median_h <= 0:
        return 0.0, {"repeating_cells": False, "similar_rectangles": 0}

    similar = [
        rect
        for rect in rectangles
        if abs(rect["w"] - median_w) <= median_w * 0.35 and abs(rect["h"] - median_h) <= median_h * 0.35
    ]
    score = _normalize_score(len(similar), 8.0)
    return score, {"repeating_cells": len(similar) >= 4, "similar_rectangles": len(similar)}


def _component_alignment(binary: Any, image_shape: Tuple[int, int], cv2: Any) -> Tuple[float, Dict[str, Any]]:
    height, width = image_shape
    count, _, stats, centroids = cv2.connectedComponentsWithStats(binary, connectivity=8)
    boxes: List[Dict[str, float]] = []
    min_area = max(6.0, float(height * width) * 0.000015)
    max_area = float(height * width) * 0.03
    for idx in range(1, count):
        area = float(stats[idx, cv2.CC_STAT_AREA])
        if area < min_area or area > max_area:
            continue
        w = float(stats[idx, cv2.CC_STAT_WIDTH])
        h = float(stats[idx, cv2.CC_STAT_HEIGHT])
        if w <= 1 or h <= 1:
            continue
        aspect = max(w / h, h / w)
        if aspect > 25:
            continue
        cx, cy = centroids[idx]
        boxes.append({"cx": float(cx), "cy": float(cy), "w": w, "h": h})

    if len(boxes) < 6:
        return 0.0, {"components_used": len(boxes), "row_clusters": 0, "col_clusters": 0}

    median_h = _median([box["h"] for box in boxes])
    median_w = _median([box["w"] for box in boxes])
    row_groups = _cluster_positions([box["cy"] for box in boxes], tolerance=max(6.0, median_h * 0.8))
    col_groups = _cluster_positions([box["cx"] for box in boxes], tolerance=max(6.0, median_w * 0.8))

    meaningful_rows = [group for group in row_groups if len(group) >= 2]
    meaningful_cols = [group for group in col_groups if len(group) >= 2]
    row_score = _normalize_score(len(meaningful_rows), 5.0)
    col_score = _normalize_score(len(meaningful_cols), 4.0)
    occupancy = _normalize_score(len(boxes), 40.0)
    score = (row_score * 0.45) + (col_score * 0.35) + (occupancy * 0.20)
    return score, {
        "components_used": len(boxes),
        "row_clusters": len(meaningful_rows),
        "col_clusters": len(meaningful_cols),
    }


def _gap_regularity(rectangles: Sequence[Dict[str, float]]) -> Tuple[float, Dict[str, Any]]:
    if len(rectangles) < 4:
        return 0.0, {"row_gap_cv": None, "col_gap_cv": None}

    row_groups = _cluster_positions(
        [rect["y"] + (rect["h"] / 2.0) for rect in rectangles],
        tolerance=_median([r["h"] for r in rectangles]) * 0.7,
    )
    col_groups = _cluster_positions(
        [rect["x"] + (rect["w"] / 2.0) for rect in rectangles],
        tolerance=_median([r["w"] for r in rectangles]) * 0.7,
    )

    def gap_cv(groups: Sequence[Sequence[float]]) -> Optional[float]:
        centers = [sum(group) / len(group) for group in groups if group]
        if len(centers) < 3:
            return None
        gaps = [centers[idx + 1] - centers[idx] for idx in range(len(centers) - 1)]
        mean_gap = sum(gaps) / len(gaps)
        if mean_gap <= 0:
            return None
        variance = sum((gap - mean_gap) ** 2 for gap in gaps) / len(gaps)
        return (variance ** 0.5) / mean_gap

    row_cv = gap_cv(row_groups)
    col_cv = gap_cv(col_groups)

    parts: List[float] = []
    if row_cv is not None:
        parts.append(1.0 - _clamp(row_cv))
    if col_cv is not None:
        parts.append(1.0 - _clamp(col_cv))
    score = sum(parts) / len(parts) if parts else 0.0
    return score, {"row_gap_cv": row_cv, "col_gap_cv": col_cv}


def _diagonal_line_ratio(gray: Any, cv2: Any) -> float:
    edges = cv2.Canny(gray, 50, 150, apertureSize=3)
    lines = cv2.HoughLinesP(edges, 1, math.pi / 180.0, threshold=40, minLineLength=25, maxLineGap=8)
    if lines is None:
        return 0.0
    total = 0
    diagonal = 0
    for item in lines:
        x1, y1, x2, y2 = item[0]
        dx = float(x2 - x1)
        dy = float(y2 - y1)
        if dx == 0 and dy == 0:
            continue
        total += 1
        angle = abs(math.degrees(math.atan2(dy, dx))) % 180.0
        if not (angle <= 12.0 or angle >= 168.0 or 78.0 <= angle <= 102.0):
            diagonal += 1
    return _safe_ratio(diagonal, total)


def _irregular_blob_ratio(binary: Any, line_mask: Any, cv2: Any) -> float:
    residual = cv2.bitwise_and(binary, cv2.bitwise_not(line_mask))
    contours, _ = cv2.findContours(residual, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    irregular_area = 0.0
    total_area = float(cv2.countNonZero(binary))
    if total_area <= 0:
        return 0.0
    for contour in contours:
        area = float(cv2.contourArea(contour))
        if area < 12.0:
            continue
        hull = cv2.convexHull(contour)
        hull_area = float(cv2.contourArea(hull))
        if hull_area <= 0:
            continue
        solidity = area / hull_area
        peri = cv2.arcLength(contour, True)
        approx = cv2.approxPolyDP(contour, 0.03 * peri, True)
        if solidity < 0.7 or len(approx) > 8:
            irregular_area += area
    return _clamp(irregular_area / total_area)


def _chart_like_score(
    long_horizontal_lines: int,
    long_vertical_lines: int,
    intersection_count: int,
    diagonal_ratio: float,
    irregular_blob_ratio: float,
    horizontal_ratio: float,
    vertical_ratio: float,
) -> float:
    axis_presence = 1.0 if long_horizontal_lines >= 1 and long_vertical_lines >= 1 else 0.0
    sparse_grid = 1.0 - _normalize_score(intersection_count, 8.0)
    line_presence = _clamp((horizontal_ratio + vertical_ratio) / 0.12)
    return _clamp(
        (axis_presence * 0.35)
        + (sparse_grid * 0.20)
        + (_clamp(diagonal_ratio / 0.35) * 0.25)
        + (_clamp(irregular_blob_ratio / 0.30) * 0.10)
        + (line_presence * 0.10)
    )


def _compute_features(image_path: Path) -> Dict[str, Any]:
    try:
        import cv2  # type: ignore
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(f"opencv heuristic classifier unavailable: {type(exc).__name__}: {exc}") from exc

    image = _read_image_with_cv(image_path, cv2, cv2.IMREAD_COLOR)

    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    gray, binary = _prepare_binary(gray, cv2)
    height, width = gray.shape[:2]
    image_area = float(height * width)
    if image_area <= 0:
        raise ValueError(f"invalid image size: {image_path}")

    horizontal_mask = _extract_line_mask(binary, "horizontal", cv2)
    vertical_mask = _extract_line_mask(binary, "vertical", cv2)
    line_mask = cv2.bitwise_or(horizontal_mask, vertical_mask)
    combined_mask = cv2.bitwise_or(binary, line_mask)

    horizontal_ratio = _safe_ratio(cv2.countNonZero(horizontal_mask), image_area)
    vertical_ratio = _safe_ratio(cv2.countNonZero(vertical_mask), image_area)
    long_horizontal_lines = _count_long_lines(horizontal_mask, "horizontal", (height, width), cv2)
    long_vertical_lines = _count_long_lines(vertical_mask, "vertical", (height, width), cv2)
    intersection_count = _count_intersections(horizontal_mask, vertical_mask, cv2)

    rectangles = _collect_rectangles(combined_mask, (height, width), cv2)
    rectangle_count = len(rectangles)
    repeating_cell_score, repeating_info = _repeating_cell_structure(rectangles)
    alignment_score, alignment_info = _component_alignment(binary, (height, width), cv2)
    gap_regularity_score, gap_info = _gap_regularity(rectangles)
    diagonal_ratio = _diagonal_line_ratio(gray, cv2)
    irregular_blob_ratio = _irregular_blob_ratio(binary, line_mask, cv2)
    chart_score = _chart_like_score(
        long_horizontal_lines=long_horizontal_lines,
        long_vertical_lines=long_vertical_lines,
        intersection_count=intersection_count,
        diagonal_ratio=diagonal_ratio,
        irregular_blob_ratio=irregular_blob_ratio,
        horizontal_ratio=horizontal_ratio,
        vertical_ratio=vertical_ratio,
    )

    horizontal_norm = _normalize_score(horizontal_ratio, 0.08)
    vertical_norm = _normalize_score(vertical_ratio, 0.08)
    long_lines_norm = _normalize_score(long_horizontal_lines + long_vertical_lines, 8.0)
    intersections_norm = _normalize_score(intersection_count, 20.0)
    rectangles_norm = _normalize_score(rectangle_count, 10.0)
    grid_support = min(horizontal_norm, vertical_norm)

    positive_score = (
        (horizontal_norm * 0.12)
        + (vertical_norm * 0.12)
        + (long_lines_norm * 0.10)
        + (intersections_norm * 0.16)
        + (rectangles_norm * 0.12)
        + (repeating_cell_score * 0.06)
        + (alignment_score * 0.05)
        + (gap_regularity_score * 0.06)
        + (grid_support * 0.10)
    )
    penalty_score = (
        (_clamp(diagonal_ratio / 0.35) * 0.16)
        + (_clamp(irregular_blob_ratio / 0.35) * 0.18)
        + (chart_score * 0.22)
    )
    raw_score = positive_score - penalty_score

    hard_table_signal = (
        intersection_count >= 6
        and rectangle_count >= 4
        and horizontal_ratio >= 0.01
        and vertical_ratio >= 0.01
    )
    hard_non_table_signal = (
        intersection_count == 0
        and rectangle_count < 2
        and (long_horizontal_lines + long_vertical_lines) < 2
    )
    chart_blob_veto = chart_score >= 0.45 and irregular_blob_ratio >= 0.10
    diagonal_blob_veto = diagonal_ratio >= 0.20 and irregular_blob_ratio >= 0.15
    sparse_rect_chart_veto = chart_score >= 0.45 and rectangle_count <= 10
    weak_grid_chart_veto = (
        gap_regularity_score < 0.40
        and chart_score >= 0.12
        and irregular_blob_ratio >= 0.10
        and rectangle_count <= 12
    )
    sparse_cell_grid_veto = (
        intersection_count >= 40
        and rectangle_count <= 8
        and repeating_cell_score < 0.40
        and gap_regularity_score < 0.55
    )

    score = _clamp(raw_score)
    if hard_table_signal:
        score = max(score, 0.60)
    if hard_non_table_signal and chart_score >= 0.45:
        score = min(score, 0.20)
    if chart_blob_veto:
        score = min(score, 0.20)
    if diagonal_blob_veto:
        score = min(score, 0.20)
    if sparse_rect_chart_veto:
        score = min(score, 0.25)
    if weak_grid_chart_veto:
        score = min(score, 0.30)
    if sparse_cell_grid_veto:
        score = min(score, 0.28)

    return {
        "image_size": {"width": width, "height": height},
        "feature_values": {
            "horizontal_line_ratio": horizontal_ratio,
            "vertical_line_ratio": vertical_ratio,
            "long_horizontal_line_count": long_horizontal_lines,
            "long_vertical_line_count": long_vertical_lines,
            "intersection_count": intersection_count,
            "rectangle_contour_count": rectangle_count,
            "repeating_cell_structure_score": repeating_cell_score,
            "connected_component_alignment_score": alignment_score,
            "gap_regularity_score": gap_regularity_score,
            "diagonal_curve_ratio": diagonal_ratio,
            "irregular_blob_ratio": irregular_blob_ratio,
            "chart_like_structure_score": chart_score,
            "grid_support_score": grid_support,
        },
        "feature_details": {
            "rectangles_used": rectangle_count,
            "repeating_cells": repeating_info,
            "component_alignment": alignment_info,
            "gap_regularity": gap_info,
        },
        "score_breakdown": {
            "positive_score": positive_score,
            "penalty_score": penalty_score,
            "raw_score": raw_score,
            "final_score": score,
        },
        "decision_flags": {
            "hard_table_signal": hard_table_signal,
            "hard_non_table_signal": hard_non_table_signal,
            "chart_blob_veto": chart_blob_veto,
            "diagonal_blob_veto": diagonal_blob_veto,
            "sparse_rect_chart_veto": sparse_rect_chart_veto,
            "weak_grid_chart_veto": weak_grid_chart_veto,
            "sparse_cell_grid_veto": sparse_cell_grid_veto,
        },
    }


def _predict_label(score: float) -> str:
    if score >= TABLE_THRESHOLD:
        return "table"
    return "non-table"


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

    features = _compute_features(resolved_path)
    score = float(features["score_breakdown"]["final_score"])
    predicted_class = _predict_label(score)
    out = {
        "file": key,
        "predicted_class": predicted_class,
        "is_table": predicted_class == "table",
        "score": score,
        "low_confidence": LOW_CONFIDENCE_THRESHOLD <= score < TABLE_THRESHOLD,
        "thresholds": {
            "table": TABLE_THRESHOLD,
            "low_confidence_floor": LOW_CONFIDENCE_THRESHOLD,
        },
        **features,
    }
    _CLASSIFY_CACHE[key] = out
    return out


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
    import cv2  # type: ignore
    import numpy as np  # type: ignore
    from PIL import Image  # type: ignore

    image = _read_image_with_cv(image_path, cv2, cv2.IMREAD_UNCHANGED)
    bg_gray = max(0, min(255, int(bg_gray)))

    if len(image.shape) == 2:
        return Image.fromarray(image).convert("RGB")

    if len(image.shape) == 3 and image.shape[2] == 4:
        alpha = image[:, :, 3:4].astype("float32") / 255.0
        rgb = image[:, :, :3][:, :, ::-1].astype("float32")
        bg = np.full(rgb.shape, bg_gray, dtype="float32")
        composited = (rgb * alpha) + (bg * (1.0 - alpha))
        return Image.fromarray(composited.clip(0, 255).astype("uint8"), mode="RGB")

    rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
    return Image.fromarray(rgb)


def _run_surya_parsed_tables(image_path: Path) -> List[Dict[str, Any]]:
    run_ctx = _suppress_external_output() if _HIDE_SURYA_LOGS else contextlib.nullcontext()
    with run_ctx:
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

    try:
        cls = classify_image(resolved_path)
    except Exception as exc:  # noqa: BLE001
        out = {
            "status": "error",
            "file": key,
            "error": f"heuristic table classification failed: {type(exc).__name__}: {exc}",
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
            "surya_attempted": False,
            "classification": cls,
        }
        _RESULT_CACHE[key] = out
        return out

    try:
        parsed_tables = _run_surya_parsed_tables(resolved_path)
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

    if not parsed_tables:
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
