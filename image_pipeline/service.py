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


def _mean(values: Sequence[float]) -> float:
    if not values:
        return 0.0
    return sum(float(v) for v in values) / float(len(values))


def _coefficient_of_variation(values: Sequence[float]) -> Optional[float]:
    if len(values) < 2:
        return None
    mean_value = _mean(values)
    if mean_value <= 0:
        return None
    variance = sum((float(value) - mean_value) ** 2 for value in values) / float(len(values))
    return (variance ** 0.5) / mean_value


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


def _load_normalized_rgb_image(image_path: Path, bg_gray: int = _DEFAULT_BG_GRAY) -> Any:
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

            if image.mode in {"RGBA", "LA"} or "transparency" in image.info:
                rgba = image.convert("RGBA")
                bg = Image.new("RGBA", rgba.size, (bg_gray, bg_gray, bg_gray, 255))
                composited = Image.alpha_composite(bg, rgba)
                return composited.convert("RGB")

            return image.convert("RGB")
    finally:
        if cleanup_dir is not None:
            shutil.rmtree(cleanup_dir, ignore_errors=True)


def _load_normalized_bgr_image(image_path: Path, bg_gray: int = _DEFAULT_BG_GRAY) -> Any:
    import cv2  # type: ignore
    import numpy as np  # type: ignore

    rgb_image = _load_normalized_rgb_image(image_path, bg_gray=bg_gray)
    rgb = np.array(rgb_image)
    return cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR)


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


def _collect_text_like_components(
    binary: Any,
    image_shape: Tuple[int, int],
    cv2: Any,
) -> Tuple[List[Dict[str, float]], Dict[str, Any]]:
    height, width = image_shape
    image_area = float(height * width)
    count, _, stats, centroids = cv2.connectedComponentsWithStats(binary, connectivity=8)

    components: List[Dict[str, float]] = []
    component_area_total = 0.0
    min_area = max(6.0, image_area * 0.000015)
    max_area = image_area * 0.03

    for idx in range(1, count):
        area = float(stats[idx, cv2.CC_STAT_AREA])
        if area < min_area or area > max_area:
            continue

        x = float(stats[idx, cv2.CC_STAT_LEFT])
        y = float(stats[idx, cv2.CC_STAT_TOP])
        w = float(stats[idx, cv2.CC_STAT_WIDTH])
        h = float(stats[idx, cv2.CC_STAT_HEIGHT])
        if w <= 1 or h <= 1:
            continue

        aspect = max(w / h, h / w)
        if aspect > 25:
            continue

        bbox_area = w * h
        if bbox_area <= 0:
            continue

        fill_ratio = area / bbox_area
        if fill_ratio < 0.08:
            continue

        cx, cy = centroids[idx]
        components.append(
            {
                "x": x,
                "y": y,
                "w": w,
                "h": h,
                "cx": float(cx),
                "cy": float(cy),
                "left": x,
                "center": float(cx),
                "right": x + w,
                "area": area,
                "fill_ratio": fill_ratio,
            }
        )
        component_area_total += area

    components.sort(key=lambda item: (item["y"], item["x"]))
    return components, {
        "components_used": len(components),
        "component_area_ratio": _safe_ratio(component_area_total, image_area),
    }


def _cluster_components_by_axis(
    components: Sequence[Dict[str, float]],
    axis_key: str,
    tolerance: float,
) -> List[List[Dict[str, float]]]:
    if not components:
        return []

    ordered = sorted(components, key=lambda item: float(item[axis_key]))
    groups: List[List[Dict[str, float]]] = [[ordered[0]]]
    for component in ordered[1:]:
        group = groups[-1]
        center = _mean([entry[axis_key] for entry in group])
        if abs(float(component[axis_key]) - center) <= max(1.0, float(tolerance)):
            group.append(component)
        else:
            groups.append([component])
    return groups


def _cluster_labeled_positions(
    pairs: Sequence[Tuple[float, int]],
    tolerance: float,
) -> List[Dict[str, Any]]:
    if not pairs:
        return []

    ordered = sorted((float(pos), int(label)) for pos, label in pairs)
    clusters: List[Dict[str, Any]] = [
        {
            "positions": [ordered[0][0]],
            "rows": {ordered[0][1]},
        }
    ]
    for position, row_idx in ordered[1:]:
        current = clusters[-1]
        center = _mean(current["positions"])
        if abs(position - center) <= max(1.0, float(tolerance)):
            current["positions"].append(position)
            current["rows"].add(row_idx)
        else:
            clusters.append({"positions": [position], "rows": {row_idx}})
    return clusters


def _smooth_sequence(values: Sequence[float], window: int) -> List[float]:
    if not values:
        return []
    half = max(0, int(window) // 2)
    out: List[float] = []
    for idx in range(len(values)):
        start = max(0, idx - half)
        end = min(len(values), idx + half + 1)
        out.append(_mean(values[start:end]))
    return out


def _gap_cv_from_centers(centers: Sequence[float]) -> Optional[float]:
    if len(centers) < 3:
        return None
    gaps = [float(centers[idx + 1]) - float(centers[idx]) for idx in range(len(centers) - 1)]
    return _coefficient_of_variation(gaps)


def _projection_peak_summary(values: Sequence[float], threshold: float) -> Dict[str, Any]:
    peaks: List[Dict[str, float]] = []
    start: Optional[int] = None
    total = 0.0
    for idx, value in enumerate(values):
        if value >= threshold:
            if start is None:
                start = idx
                total = 0.0
            total += float(value)
            continue
        if start is None:
            continue
        end = idx - 1
        peaks.append(
            {
                "start": float(start),
                "end": float(end),
                "center": float(start + end) / 2.0,
                "strength": total / float(max(1, end - start + 1)),
            }
        )
        start = None
        total = 0.0

    if start is not None:
        end = len(values) - 1
        peaks.append(
            {
                "start": float(start),
                "end": float(end),
                "center": float(start + end) / 2.0,
                "strength": total / float(max(1, end - start + 1)),
            }
        )

    centers = [peak["center"] for peak in peaks]
    gap_cv = _gap_cv_from_centers(centers)
    return {
        "peak_count": len(peaks),
        "gap_cv": gap_cv,
        "peak_strength_mean": _mean([peak["strength"] for peak in peaks]),
    }


def _best_anchor_alignment(
    meaningful_rows: Sequence[Sequence[Dict[str, float]]],
    tolerance: float,
) -> Dict[str, Any]:
    out: Dict[str, Any] = {
        "best_anchor": None,
        "best_anchor_score": 0.0,
        "best_anchor_cluster_count": 0,
        "best_anchor_row_coverage": 0.0,
        "left_anchor_clusters": 0,
        "center_anchor_clusters": 0,
        "right_anchor_clusters": 0,
    }
    row_count = len(meaningful_rows)
    if row_count < 2:
        return out

    min_rows = 3 if row_count >= 5 else 2
    best_score = 0.0

    for anchor_key in ("left", "center", "right"):
        labeled_positions: List[Tuple[float, int]] = []
        for row_idx, row in enumerate(meaningful_rows):
            for component in row:
                labeled_positions.append((float(component[anchor_key]), row_idx))

        recurring = [
            cluster
            for cluster in _cluster_labeled_positions(labeled_positions, tolerance=tolerance)
            if len(cluster["rows"]) >= min_rows
        ]
        cluster_count = len(recurring)
        if anchor_key == "left":
            out["left_anchor_clusters"] = cluster_count
        elif anchor_key == "center":
            out["center_anchor_clusters"] = cluster_count
        else:
            out["right_anchor_clusters"] = cluster_count

        if not recurring:
            continue

        row_coverages = [len(cluster["rows"]) / float(row_count) for cluster in recurring]
        cluster_score = _normalize_score(cluster_count, 4.0)
        coverage_score = max(row_coverages)
        score = _clamp((cluster_score * 0.55) + (coverage_score * 0.45))
        if score <= best_score:
            continue

        best_score = score
        out["best_anchor"] = anchor_key
        out["best_anchor_score"] = score
        out["best_anchor_cluster_count"] = cluster_count
        out["best_anchor_row_coverage"] = coverage_score

    return out


def _header_body_transition_score(meaningful_rows: Sequence[Sequence[Dict[str, float]]]) -> float:
    if len(meaningful_rows) < 4:
        return 0.0

    header = meaningful_rows[0]
    body = meaningful_rows[1:]
    header_count = float(len(header))
    body_counts = [float(len(row)) for row in body]
    body_count_median = _median(body_counts)

    header_width = _median([component["w"] for component in header])
    body_width = _median([component["w"] for row in body for component in row])
    if body_count_median <= 0 or body_width <= 0:
        return 0.0

    count_delta = abs(header_count - body_count_median) / body_count_median
    width_delta = max(0.0, (header_width / body_width) - 1.0)
    return _clamp((_clamp(count_delta / 0.60) * 0.55) + (_clamp(width_delta / 0.80) * 0.45))


def _analyze_text_layout(binary: Any, image_shape: Tuple[int, int], cv2: Any) -> Dict[str, Any]:
    height, width = image_shape
    components, component_info = _collect_text_like_components(binary, image_shape, cv2)
    component_area_ratio = float(component_info["component_area_ratio"])

    out: Dict[str, Any] = {
        **component_info,
        "meaningful_row_count": 0,
        "meaningful_col_count": 0,
        "rows_with_three_plus_ratio": 0.0,
        "median_components_per_row": 0.0,
        "row_component_count_cv": None,
        "row_component_stability_score": 0.0,
        "compact_component_ratio": 0.0,
        "header_body_transition_score": 0.0,
        "projection_row_peak_count": 0,
        "projection_col_peak_count": 0,
        "projection_row_peak_score": 0.0,
        "projection_col_peak_score": 0.0,
        "projection_row_regularity_score": 0.0,
        "projection_col_regularity_score": 0.0,
        "projection_row_gap_cv": None,
        "projection_col_gap_cv": None,
        "best_anchor": None,
        "best_anchor_score": 0.0,
        "best_anchor_cluster_count": 0,
        "best_anchor_row_coverage": 0.0,
        "left_anchor_clusters": 0,
        "center_anchor_clusters": 0,
        "right_anchor_clusters": 0,
        "row_groups_total": 0,
        "column_groups_total": 0,
    }

    if len(components) < 6:
        return out

    median_h = _median([component["h"] for component in components])
    median_w = _median([component["w"] for component in components])
    row_tolerance = max(6.0, median_h * 0.85)
    col_tolerance = max(6.0, median_w * 0.85)

    row_groups = _cluster_components_by_axis(components, "cy", tolerance=row_tolerance)
    col_groups = _cluster_components_by_axis(components, "cx", tolerance=col_tolerance)
    meaningful_rows = [sorted(group, key=lambda item: item["x"]) for group in row_groups if len(group) >= 2]
    meaningful_cols = [group for group in col_groups if len(group) >= 2]

    row_counts = [float(len(row)) for row in meaningful_rows]
    compact_threshold = _median([component["area"] for component in components]) * 1.8
    if compact_threshold > 0:
        compact_count = sum(1 for component in components if component["area"] <= compact_threshold)
        out["compact_component_ratio"] = compact_count / float(len(components))

    row_count_cv = _coefficient_of_variation(row_counts)
    row_component_stability_score = 0.0
    if row_count_cv is not None:
        row_component_stability_score = 1.0 - _clamp(row_count_cv / 0.65)

    anchor_info = _best_anchor_alignment(meaningful_rows, tolerance=max(6.0, median_w * 0.75))

    row_projection = [float(value) / (255.0 * float(width)) for value in binary.sum(axis=1)]
    col_projection = [float(value) / (255.0 * float(height)) for value in binary.sum(axis=0)]
    row_projection = _smooth_sequence(row_projection, window=5)
    col_projection = _smooth_sequence(col_projection, window=5)
    row_threshold = max(0.015, min(0.12, max(row_projection) * 0.30 if row_projection else 0.0))
    col_threshold = max(0.015, min(0.12, max(col_projection) * 0.30 if col_projection else 0.0))
    row_peak_info = _projection_peak_summary(row_projection, threshold=row_threshold)
    col_peak_info = _projection_peak_summary(col_projection, threshold=col_threshold)

    row_peak_regularity_score = 0.0
    if row_peak_info["gap_cv"] is not None:
        row_peak_regularity_score = 1.0 - _clamp(float(row_peak_info["gap_cv"]) / 0.65)
    col_peak_regularity_score = 0.0
    if col_peak_info["gap_cv"] is not None:
        col_peak_regularity_score = 1.0 - _clamp(float(col_peak_info["gap_cv"]) / 0.65)

    out.update(anchor_info)
    out.update(
        {
            "meaningful_row_count": len(meaningful_rows),
            "meaningful_col_count": len(meaningful_cols),
            "rows_with_three_plus_ratio": _safe_ratio(
                sum(1 for row in meaningful_rows if len(row) >= 3),
                len(meaningful_rows),
            ),
            "median_components_per_row": _median(row_counts),
            "row_component_count_cv": row_count_cv,
            "row_component_stability_score": row_component_stability_score,
            "header_body_transition_score": _header_body_transition_score(meaningful_rows),
            "projection_row_peak_count": int(row_peak_info["peak_count"]),
            "projection_col_peak_count": int(col_peak_info["peak_count"]),
            "projection_row_peak_score": _normalize_score(float(row_peak_info["peak_count"]), 6.0),
            "projection_col_peak_score": _normalize_score(float(col_peak_info["peak_count"]), 5.0),
            "projection_row_regularity_score": row_peak_regularity_score,
            "projection_col_regularity_score": col_peak_regularity_score,
            "projection_row_gap_cv": row_peak_info["gap_cv"],
            "projection_col_gap_cv": col_peak_info["gap_cv"],
            "row_groups_total": len(row_groups),
            "column_groups_total": len(col_groups),
        }
    )
    out["component_area_ratio"] = component_area_ratio
    return out


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


def _dominant_grid_bbox(
    horizontal_mask: Any,
    vertical_mask: Any,
    binary: Any,
    rectangles: Sequence[Dict[str, float]],
    image_shape: Tuple[int, int],
    cv2: Any,
) -> Tuple[float, Dict[str, Any]]:
    height, width = image_shape
    image_area = float(height * width)
    if image_area <= 0:
        return 0.0, {"found": False}

    grid_mask = cv2.bitwise_or(horizontal_mask, vertical_mask)
    grid_pixels = float(cv2.countNonZero(grid_mask))
    binary_pixels = float(cv2.countNonZero(binary))
    if grid_pixels <= 0:
        return 0.0, {"found": False, "grid_pixels": 0}

    bridge_kernel = cv2.getStructuringElement(
        cv2.MORPH_RECT,
        (max(5, width // 160), max(5, height // 160)),
    )
    merged = cv2.dilate(grid_mask, bridge_kernel, iterations=2)
    contours, _ = cv2.findContours(merged, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    best_score = 0.0
    best_info: Dict[str, Any] = {"found": False, "grid_pixels": int(grid_pixels)}
    min_bbox_area = image_area * 0.03

    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        bbox_area = float(w * h)
        if bbox_area < min_bbox_area:
            continue

        x2 = min(width, x + w)
        y2 = min(height, y + h)
        line_pixels_inside = float(cv2.countNonZero(grid_mask[y:y2, x:x2]))
        binary_pixels_inside = float(cv2.countNonZero(binary[y:y2, x:x2]))

        area_ratio = _clamp(bbox_area / image_area)
        line_density = _safe_ratio(line_pixels_inside, bbox_area)
        outside_noise_ratio = 0.0
        if binary_pixels > 0:
            outside_noise_ratio = _clamp((binary_pixels - binary_pixels_inside) / binary_pixels)

        rect_inside = 0
        for rect in rectangles:
            cx = rect["x"] + (rect["w"] / 2.0)
            cy = rect["y"] + (rect["h"] / 2.0)
            if x <= cx <= x2 and y <= cy <= y2:
                rect_inside += 1
        rect_coverage = _safe_ratio(rect_inside, len(rectangles)) if rectangles else 0.0

        score = _clamp(
            (_normalize_score(area_ratio, 0.45) * 0.45)
            + (_normalize_score(line_density, 0.08) * 0.15)
            + (rect_coverage * 0.20)
            + ((1.0 - outside_noise_ratio) * 0.20)
        )
        if score <= best_score:
            continue

        best_score = score
        best_info = {
            "found": True,
            "x": int(x),
            "y": int(y),
            "w": int(w),
            "h": int(h),
            "area_ratio": area_ratio,
            "line_density": line_density,
            "outside_noise_ratio": outside_noise_ratio,
            "rect_coverage_ratio": rect_coverage,
            "rectangles_inside": rect_inside,
            "grid_pixels": int(grid_pixels),
        }

    return best_score, best_info


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
    rectangle_count: int,
    repeating_cell_score: float,
    gap_regularity_score: float,
) -> float:
    axis_presence = 1.0 if long_horizontal_lines >= 1 and long_vertical_lines >= 1 else 0.0
    sparse_grid = 1.0 - _normalize_score(intersection_count, 28.0)
    line_presence = _clamp((horizontal_ratio + vertical_ratio) / 0.12) * sparse_grid
    weak_cell_structure = 1.0 - max(repeating_cell_score, _normalize_score(rectangle_count, 18.0))
    irregular_spacing = 1.0 - gap_regularity_score
    return _clamp(
        ((axis_presence * sparse_grid) * 0.18)
        + (sparse_grid * 0.18)
        + (_clamp(diagonal_ratio / 0.35) * 0.26)
        + (_clamp(irregular_blob_ratio / 0.30) * 0.12)
        + (line_presence * 0.10)
        + (weak_cell_structure * 0.08)
        + (irregular_spacing * 0.08)
    )


def _grid_table_detector(
    horizontal_ratio: float,
    vertical_ratio: float,
    long_horizontal_lines: int,
    long_vertical_lines: int,
    intersection_count: int,
    rectangle_count: int,
    repeating_cell_score: float,
    gap_regularity_score: float,
    dominant_grid_bbox_score: float,
) -> Dict[str, Any]:
    horizontal_norm = _normalize_score(horizontal_ratio, 0.08)
    vertical_norm = _normalize_score(vertical_ratio, 0.08)
    line_count_norm = _normalize_score(long_horizontal_lines + long_vertical_lines, 8.0)
    intersections_norm = _normalize_score(intersection_count, 20.0)
    rectangles_norm = _normalize_score(rectangle_count, 10.0)
    grid_support = min(horizontal_norm, vertical_norm)
    score = _clamp(
        (horizontal_norm * 0.12)
        + (vertical_norm * 0.12)
        + (line_count_norm * 0.08)
        + (intersections_norm * 0.20)
        + (rectangles_norm * 0.10)
        + (repeating_cell_score * 0.08)
        + (gap_regularity_score * 0.08)
        + (grid_support * 0.10)
        + (dominant_grid_bbox_score * 0.20)
    )
    return {
        "score": score,
        "components": {
            "horizontal_line_score": horizontal_norm,
            "vertical_line_score": vertical_norm,
            "line_count_score": line_count_norm,
            "intersection_score": intersections_norm,
            "rectangle_score": rectangles_norm,
            "repeating_cell_score": repeating_cell_score,
            "gap_regularity_score": gap_regularity_score,
            "grid_support_score": grid_support,
            "dominant_grid_bbox_score": dominant_grid_bbox_score,
        },
    }


def _alignment_table_detector(layout_info: Dict[str, Any]) -> Dict[str, Any]:
    row_sufficiency = _normalize_score(float(layout_info.get("meaningful_row_count", 0)), 4.0)
    row_gate = _clamp(0.10 + (row_sufficiency * 0.90))
    component_count_score = _normalize_score(float(layout_info.get("components_used", 0)), 36.0)
    component_area_score = _normalize_score(float(layout_info.get("component_area_ratio", 0.0)), 0.08)
    row_pattern_score = _clamp(
        (_normalize_score(float(layout_info.get("meaningful_row_count", 0)), 5.0) * 0.60)
        + (float(layout_info.get("row_component_stability_score", 0.0)) * 0.45)
    )
    column_pattern_score = _clamp(
        (_normalize_score(float(layout_info.get("best_anchor_cluster_count", 0)), 4.0) * 0.30)
        + (float(layout_info.get("best_anchor_row_coverage", 0.0)) * 0.20)
        + (_normalize_score(float(layout_info.get("meaningful_col_count", 0)), 4.0) * 0.20)
        + (float(layout_info.get("projection_col_regularity_score", 0.0)) * 0.15)
        + (float(layout_info.get("projection_col_peak_score", 0.0)) * 0.15)
    )
    projection_pattern_score = _clamp(
        (float(layout_info.get("projection_row_peak_score", 0.0)) * 0.35)
        + (float(layout_info.get("projection_col_peak_score", 0.0)) * 0.20)
        + (float(layout_info.get("projection_row_regularity_score", 0.0)) * 0.30)
        + (float(layout_info.get("projection_col_regularity_score", 0.0)) * 0.15)
    )
    occupancy_score = _clamp(
        (component_count_score * 0.60)
        + (min(component_count_score, component_area_score) * 0.40)
    )
    score = _clamp(
        (
            (row_pattern_score * 0.34)
            + (column_pattern_score * 0.28)
            + (projection_pattern_score * 0.20)
            + (occupancy_score * 0.18)
        )
        * row_gate
    )
    return {
        "score": score,
        "components": {
            "row_gate": row_gate,
            "row_pattern_score": row_pattern_score,
            "column_pattern_score": column_pattern_score,
            "projection_pattern_score": projection_pattern_score,
            "occupancy_score": occupancy_score,
        },
    }


def _dense_table_detector(layout_info: Dict[str, Any]) -> Dict[str, Any]:
    row_sufficiency = _normalize_score(float(layout_info.get("meaningful_row_count", 0)), 4.0)
    row_gate = _clamp(0.10 + (row_sufficiency * 0.90))
    component_count_score = _normalize_score(float(layout_info.get("components_used", 0)), 48.0)
    component_area_score = _normalize_score(float(layout_info.get("component_area_ratio", 0.0)), 0.11)
    density_score = _clamp(
        (min(component_area_score, component_count_score) * 0.70)
        + (component_count_score * 0.30)
    )
    compact_matrix_score = _clamp(
        (float(layout_info.get("compact_component_ratio", 0.0)) * 0.30)
        + (_normalize_score(float(layout_info.get("meaningful_row_count", 0)), 6.0) * 0.25)
        + (float(layout_info.get("rows_with_three_plus_ratio", 0.0)) * 0.25)
        + (_normalize_score(float(layout_info.get("median_components_per_row", 0.0)), 4.5) * 0.20)
    )
    partial_alignment_score = _clamp(
        (float(layout_info.get("best_anchor_score", 0.0)) * 0.40)
        + (float(layout_info.get("best_anchor_row_coverage", 0.0)) * 0.20)
        + (float(layout_info.get("projection_col_peak_score", 0.0)) * 0.20)
        + (float(layout_info.get("row_component_stability_score", 0.0)) * 0.20)
    )
    header_body_score = float(layout_info.get("header_body_transition_score", 0.0))
    score = _clamp(
        (
            (density_score * 0.28)
            + (compact_matrix_score * 0.30)
            + (partial_alignment_score * 0.26)
            + (header_body_score * 0.16)
        )
        * row_gate
    )
    return {
        "score": score,
        "components": {
            "row_gate": row_gate,
            "density_score": density_score,
            "compact_matrix_score": compact_matrix_score,
            "partial_alignment_score": partial_alignment_score,
            "header_body_score": header_body_score,
        },
    }


def _non_table_veto_detectors(
    chart_score: float,
    diagonal_ratio: float,
    irregular_blob_ratio: float,
    horizontal_ratio: float,
    vertical_ratio: float,
    layout_info: Dict[str, Any],
) -> Dict[str, Any]:
    structure_presence = max(
        float(layout_info.get("best_anchor_score", 0.0)),
        _normalize_score(float(layout_info.get("meaningful_row_count", 0)), 5.0),
        _normalize_score(float(layout_info.get("meaningful_col_count", 0)), 4.0),
    )
    line_presence = _clamp((horizontal_ratio + vertical_ratio) / 0.10)
    text_presence = _clamp(
        (_normalize_score(float(layout_info.get("components_used", 0)), 24.0) * 0.55)
        + (_normalize_score(float(layout_info.get("component_area_ratio", 0.0)), 0.06) * 0.45)
    )

    chart_veto_score = _clamp(
        (chart_score * 0.55)
        + (_clamp(diagonal_ratio / 0.25) * 0.10)
        + (_clamp(irregular_blob_ratio / 0.22) * 0.10)
        + ((1.0 - structure_presence) * 0.15)
        + ((1.0 - text_presence) * 0.10)
    )
    diagram_veto_score = _clamp(
        (_clamp(diagonal_ratio / 0.20) * 0.45)
        + (_clamp(irregular_blob_ratio / 0.24) * 0.25)
        + ((1.0 - structure_presence) * 0.20)
        + ((1.0 - line_presence) * 0.10)
    )
    photo_veto_score = _clamp(
        (_clamp(irregular_blob_ratio / 0.28) * 0.45)
        + ((1.0 - text_presence) * 0.30)
        + ((1.0 - line_presence) * 0.15)
        + ((1.0 - structure_presence) * 0.10)
    )
    strongest_name, strongest_score = max(
        (
            ("chart", chart_veto_score),
            ("diagram", diagram_veto_score),
            ("photo", photo_veto_score),
        ),
        key=lambda item: item[1],
    )
    return {
        "chart": {"score": chart_veto_score, "active": chart_veto_score >= 0.62},
        "diagram": {"score": diagram_veto_score, "active": diagram_veto_score >= 0.62},
        "photo": {"score": photo_veto_score, "active": photo_veto_score >= 0.62},
        "strongest_veto": strongest_name,
        "strongest_veto_score": strongest_score,
        "line_presence_score": line_presence,
        "structure_presence_score": structure_presence,
        "text_presence_score": text_presence,
    }


def _compute_features(image_path: Path) -> Dict[str, Any]:
    try:
        import cv2  # type: ignore
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(f"opencv heuristic classifier unavailable: {type(exc).__name__}: {exc}") from exc

    image = _load_normalized_bgr_image(image_path)

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
    layout_info = _analyze_text_layout(binary, (height, width), cv2)
    gap_regularity_score, gap_info = _gap_regularity(rectangles)
    dominant_grid_bbox_score, dominant_grid_bbox_info = _dominant_grid_bbox(
        horizontal_mask,
        vertical_mask,
        binary,
        rectangles,
        (height, width),
        cv2,
    )
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
        rectangle_count=rectangle_count,
        repeating_cell_score=repeating_cell_score,
        gap_regularity_score=gap_regularity_score,
    )

    dominant_bbox_area_raw = dominant_grid_bbox_info.get("area_ratio", 0.0)
    dominant_bbox_outside_raw = dominant_grid_bbox_info.get("outside_noise_ratio", 1.0)
    dominant_bbox_cover_raw = dominant_grid_bbox_info.get("rect_coverage_ratio", 0.0)
    dominant_bbox_area_ratio = float(dominant_bbox_area_raw if dominant_bbox_area_raw is not None else 0.0)
    dominant_bbox_outside_noise_ratio = float(
        dominant_bbox_outside_raw if dominant_bbox_outside_raw is not None else 1.0
    )
    dominant_bbox_rect_coverage_ratio = float(
        dominant_bbox_cover_raw if dominant_bbox_cover_raw is not None else 0.0
    )

    grid_detector = _grid_table_detector(
        horizontal_ratio=horizontal_ratio,
        vertical_ratio=vertical_ratio,
        long_horizontal_lines=long_horizontal_lines,
        long_vertical_lines=long_vertical_lines,
        intersection_count=intersection_count,
        rectangle_count=rectangle_count,
        repeating_cell_score=repeating_cell_score,
        gap_regularity_score=gap_regularity_score,
        dominant_grid_bbox_score=dominant_grid_bbox_score,
    )
    alignment_detector = _alignment_table_detector(layout_info)
    dense_detector = _dense_table_detector(layout_info)

    detector_scores = {
        "grid": grid_detector,
        "alignment": alignment_detector,
        "dense": dense_detector,
    }
    strongest_detector_name, strongest_detector = max(
        detector_scores.items(),
        key=lambda item: float(item[1]["score"]),
    )
    final_table_score = float(strongest_detector["score"])

    veto_breakdown = _non_table_veto_detectors(
        chart_score=chart_score,
        diagonal_ratio=diagonal_ratio,
        irregular_blob_ratio=irregular_blob_ratio,
        horizontal_ratio=horizontal_ratio,
        vertical_ratio=vertical_ratio,
        layout_info=layout_info,
    )
    strongest_veto_score = float(veto_breakdown["strongest_veto_score"])
    low_table_evidence = final_table_score < 0.55
    weak_non_grid_table = (
        strongest_detector_name != "grid"
        and float(grid_detector["score"]) < 0.40
        and strongest_veto_score >= 0.58
    )
    sparse_text_grid_veto = (
        float(grid_detector["score"]) >= 0.72
        and float(alignment_detector["score"]) < 0.40
        and float(dense_detector["score"]) < 0.55
        and int(layout_info.get("meaningful_row_count", 0)) <= 1
        and int(layout_info.get("components_used", 0)) <= 18
    )
    veto_applied = (
        (low_table_evidence and strongest_veto_score >= 0.62)
        or weak_non_grid_table
        or sparse_text_grid_veto
    )
    score = final_table_score
    if veto_applied:
        score = min(score, 0.24)

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
            "gap_regularity_score": gap_regularity_score,
            "text_component_count": layout_info.get("components_used"),
            "text_component_area_ratio": layout_info.get("component_area_ratio"),
            "meaningful_row_count": layout_info.get("meaningful_row_count"),
            "meaningful_col_count": layout_info.get("meaningful_col_count"),
            "row_component_stability_score": layout_info.get("row_component_stability_score"),
            "rows_with_three_plus_ratio": layout_info.get("rows_with_three_plus_ratio"),
            "median_components_per_row": layout_info.get("median_components_per_row"),
            "best_anchor_score": layout_info.get("best_anchor_score"),
            "best_anchor_row_coverage": layout_info.get("best_anchor_row_coverage"),
            "compact_component_ratio": layout_info.get("compact_component_ratio"),
            "projection_row_peak_count": layout_info.get("projection_row_peak_count"),
            "projection_col_peak_count": layout_info.get("projection_col_peak_count"),
            "projection_row_peak_score": layout_info.get("projection_row_peak_score"),
            "projection_col_peak_score": layout_info.get("projection_col_peak_score"),
            "projection_row_regularity_score": layout_info.get("projection_row_regularity_score"),
            "projection_col_regularity_score": layout_info.get("projection_col_regularity_score"),
            "header_body_transition_score": layout_info.get("header_body_transition_score"),
            "alignment_row_gate": alignment_detector["components"].get("row_gate"),
            "dense_row_gate": dense_detector["components"].get("row_gate"),
            "diagonal_curve_ratio": diagonal_ratio,
            "irregular_blob_ratio": irregular_blob_ratio,
            "chart_like_structure_score": chart_score,
            "dominant_grid_bbox_score": dominant_grid_bbox_score,
            "dominant_grid_bbox_area_ratio": dominant_bbox_area_ratio,
            "dominant_grid_bbox_outside_noise_ratio": dominant_bbox_outside_noise_ratio,
            "dominant_grid_bbox_rect_coverage_ratio": dominant_bbox_rect_coverage_ratio,
        },
        "feature_details": {
            "grid": {
                "rectangles_used": rectangle_count,
                "repeating_cells": repeating_info,
                "gap_regularity": gap_info,
                "dominant_grid_bbox": dominant_grid_bbox_info,
            },
            "alignment": layout_info,
            "dense": {
                "component_area_ratio": layout_info.get("component_area_ratio"),
                "compact_component_ratio": layout_info.get("compact_component_ratio"),
                "header_body_transition_score": layout_info.get("header_body_transition_score"),
                "rows_with_three_plus_ratio": layout_info.get("rows_with_three_plus_ratio"),
            },
            "veto": {
                "line_presence_score": veto_breakdown.get("line_presence_score"),
                "structure_presence_score": veto_breakdown.get("structure_presence_score"),
                "text_presence_score": veto_breakdown.get("text_presence_score"),
            },
        },
        "detector_scores": {
            name: {"score": detector["score"], "components": detector["components"]}
            for name, detector in detector_scores.items()
        },
        "veto_breakdown": veto_breakdown,
        "score_breakdown": {
            "grid_score": grid_detector["score"],
            "alignment_score": alignment_detector["score"],
            "dense_score": dense_detector["score"],
            "final_table_score": final_table_score,
            "strongest_veto_score": strongest_veto_score,
            "final_score": score,
        },
        "decision_flags": {
            "strongest_detector_is_grid": strongest_detector_name == "grid",
            "strongest_detector_is_alignment": strongest_detector_name == "alignment",
            "strongest_detector_is_dense": strongest_detector_name == "dense",
            "low_table_evidence": low_table_evidence,
            "weak_non_grid_table": weak_non_grid_table,
            "sparse_text_grid_veto": sparse_text_grid_veto,
            "veto_applied": veto_applied,
            "strong_chart_evidence": bool(veto_breakdown["chart"]["active"]),
            "strong_diagram_evidence": bool(veto_breakdown["diagram"]["active"]),
            "strong_photo_evidence": bool(veto_breakdown["photo"]["active"]),
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
    return _load_normalized_rgb_image(image_path, bg_gray=bg_gray)


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

    if _DISABLE_SURYA:
        out = {
            "status": "not_table",
            "file": key,
            "predicted_class": cls.get("predicted_class"),
            "score": cls.get("score"),
            "low_confidence": cls.get("low_confidence"),
            "surya_attempted": False,
            "classification": cls,
            "reason": "surya_disabled_by_env",
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
