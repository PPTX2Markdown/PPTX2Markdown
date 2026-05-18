from __future__ import annotations

import importlib
import json
import logging
import os
from pathlib import Path
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

_IMAGE_TABLE_PIPELINE_MODULE: Optional[object] = None
_IMAGE_TABLE_PIPELINE_IMPORT_ERROR: Optional[str] = None
_IMAGE_TABLE_DEBUG_JSON = os.getenv("IMAGE_TABLE_DEBUG_JSON", "").strip().lower() in {"1", "true", "yes", "on"}


def _load_image_table_pipeline() -> Tuple[Optional[object], Optional[str]]:
    global _IMAGE_TABLE_PIPELINE_MODULE, _IMAGE_TABLE_PIPELINE_IMPORT_ERROR
    if _IMAGE_TABLE_PIPELINE_MODULE is not None:
        return _IMAGE_TABLE_PIPELINE_MODULE, None
    if _IMAGE_TABLE_PIPELINE_IMPORT_ERROR is not None:
        return None, _IMAGE_TABLE_PIPELINE_IMPORT_ERROR

    import_errors: List[str] = []
    for module_name in ("image_pipeline.service", "image_table_pipeline"):
        try:
            _IMAGE_TABLE_PIPELINE_MODULE = importlib.import_module(module_name)
            return _IMAGE_TABLE_PIPELINE_MODULE, None
        except Exception as exc:  # noqa: BLE001
            import_errors.append(f"{module_name}: {type(exc).__name__}: {exc}")

    _IMAGE_TABLE_PIPELINE_IMPORT_ERROR = "; ".join(import_errors)
    return None, _IMAGE_TABLE_PIPELINE_IMPORT_ERROR


def convert_picture_to_table_markdown(
    image_path: str,
) -> Tuple[Optional[str], Optional[str], bool, Optional[Dict[str, object]]]:
    module, import_error = _load_image_table_pipeline()
    if module is None:
        return None, f"image-table pipeline unavailable: {import_error}", True, None
    try:
        result = module.extract_table_markdown_from_image(Path(image_path), header_rows=1)
    except Exception as exc:  # noqa: BLE001
        return (
            None,
            f"image-table pipeline failed on {Path(image_path).name}: {type(exc).__name__}: {exc}",
            False,
            None,
        )

    if not isinstance(result, dict):
        return None, f"image-table pipeline returned invalid payload: {type(result).__name__}", False, None

    _log_image_table_evaluation(image_path, result)

    status = str(result.get("status", "error"))
    if status == "table":
        markdown = result.get("markdown")
        if isinstance(markdown, str) and markdown.strip():
            return markdown, None, False, result
        return None, f"image-table pipeline rendered empty markdown: {Path(image_path).name}", False, result
    if status == "table_skipped":
        return None, None, False, result
    if status == "not_table":
        return None, None, False, result

    error = result.get("error")
    if not isinstance(error, str) or not error.strip():
        error = "unknown image-table pipeline error"
    return None, f"{Path(image_path).name}: {error}", False, result


def _fmt_eval_value(value: object, digits: int = 3) -> str:
    if isinstance(value, (int, float)):
        return f"{float(value):.{digits}f}"
    return "n/a"


def _to_jsonable_copy(value: object) -> object:
    return json.loads(json.dumps(value, ensure_ascii=False))


def _log_image_table_evaluation(image_path: str, result: Dict[str, object]) -> None:
    name = Path(image_path).name
    status = str(result.get("status", "error"))
    predicted = str(result.get("predicted_class", "unknown"))

    classification = result.get("classification")
    if not isinstance(classification, dict):
        classification = {}
    feature_values = classification.get("feature_values")
    if not isinstance(feature_values, dict):
        feature_values = None
    score_breakdown = classification.get("score_breakdown")
    if not isinstance(score_breakdown, dict):
        score_breakdown = None
    decision_flags = classification.get("decision_flags")
    if not isinstance(decision_flags, dict):
        decision_flags = None
    feature_details = classification.get("feature_details")
    if not isinstance(feature_details, dict):
        feature_details = None
    thresholds = classification.get("thresholds")
    if not isinstance(thresholds, dict):
        thresholds = None
    image_size = classification.get("image_size")
    if not isinstance(image_size, dict):
        image_size = None
    visible_image = classification.get("visible_image")
    if not isinstance(visible_image, dict):
        visible_image = None

    score = result.get(
        "score",
        classification.get("score", score_breakdown.get("final_score") if isinstance(score_breakdown, dict) else None),
    )
    table_count = result.get("table_count")
    reason = result.get("reason")
    error = result.get("error")
    surya_attempted = bool(result.get("surya_attempted"))
    surya_quality_ok = result.get("surya_quality_ok")

    geometry_score = score_breakdown.get("geometry_score") if isinstance(score_breakdown, dict) else None
    structure_score = score_breakdown.get("structure_score") if isinstance(score_breakdown, dict) else None
    width_score = score_breakdown.get("width_score") if isinstance(score_breakdown, dict) else None
    height_score = score_breakdown.get("height_score") if isinstance(score_breakdown, dict) else None
    iou_score = score_breakdown.get("iou_score") if isinstance(score_breakdown, dict) else None
    center_score = score_breakdown.get("center_score") if isinstance(score_breakdown, dict) else None
    box_count_score = score_breakdown.get("box_count_score") if isinstance(score_breakdown, dict) else None

    summary = (
        f"[image-table] {name} "
        f"status={status} "
        f"pred={predicted} "
        f"score={_fmt_eval_value(score)} "
        f"det(geom={_fmt_eval_value(geometry_score)},"
        f"struct={_fmt_eval_value(structure_score)},"
        f"w={_fmt_eval_value(width_score)},"
        f"h={_fmt_eval_value(height_score)},"
        f"iou={_fmt_eval_value(iou_score)},"
        f"center={_fmt_eval_value(center_score)},"
        f"boxes={_fmt_eval_value(box_count_score)}) "
        f"bbox(area_ratio={_fmt_eval_value(feature_values.get('table_union_area_ratio') if isinstance(feature_values, dict) else None)},"
        f"bbox_ratio={_fmt_eval_value(feature_values.get('table_union_bbox_area_ratio') if isinstance(feature_values, dict) else None)},"
        f"w_ratio={_fmt_eval_value(feature_values.get('table_union_bbox_width_ratio') if isinstance(feature_values, dict) else None)},"
        f"h_ratio={_fmt_eval_value(feature_values.get('table_union_bbox_height_ratio') if isinstance(feature_values, dict) else None)},"
        f"iou={_fmt_eval_value(feature_values.get('table_union_bbox_iou') if isinstance(feature_values, dict) else None)},"
        f"offset={_fmt_eval_value(feature_values.get('table_union_center_offset') if isinstance(feature_values, dict) else None)}) "
        f"count(tbl={_fmt_eval_value(feature_values.get('detected_table_count') if isinstance(feature_values, dict) else None, digits=0)},"
        f"rows={_fmt_eval_value(feature_values.get('detected_row_count') if isinstance(feature_values, dict) else None, digits=0)},"
        f"cols={_fmt_eval_value(feature_values.get('detected_col_count') if isinstance(feature_values, dict) else None, digits=0)},"
        f"cells={_fmt_eval_value(feature_values.get('detected_cell_count') if isinstance(feature_values, dict) else None, digits=0)},"
        f"box={_fmt_eval_value(feature_values.get('detected_box_count') if isinstance(feature_values, dict) else None, digits=0)})"
    )

    extras: List[str] = []
    extras.append(f"surya={'run' if surya_attempted else 'skip'}")
    if isinstance(table_count, int):
        extras.append(f"tables={table_count}")
    if isinstance(surya_quality_ok, bool):
        extras.append(f"surya_quality={'ok' if surya_quality_ok else 'low'}")
    if isinstance(reason, str) and reason.strip():
        extras.append(f"reason={reason}")
    if isinstance(error, str) and error.strip():
        extras.append(f"error={error}")
    active_flags = [name for name, active in decision_flags.items() if active] if isinstance(decision_flags, dict) else []
    if active_flags:
        extras.append("flags=" + ",".join(sorted(active_flags)))
    if extras:
        summary = f"{summary} " + " ".join(extras)

    logger.info(summary)
    if _IMAGE_TABLE_DEBUG_JSON:
        feature_details_payload = {
            "visible_image_bbox": feature_details.get("visible_image_bbox") if isinstance(feature_details, dict) else None,
            "detected_table_union_bbox": (
                feature_details.get("detected_table_union_bbox") if isinstance(feature_details, dict) else None
            ),
            "detected_tables": feature_details.get("detected_tables") if isinstance(feature_details, dict) else None,
        }
        if isinstance(feature_details, dict):
            feature_details_payload.update(feature_details)

        debug_payload = {
            "kind": "image_table_debug",
            "file": str(Path(str(result.get("file", image_path))).expanduser().resolve()),
            "image_name": name,
            "status": status,
            "predicted_class": result.get("predicted_class"),
            "score": score,
            "low_confidence": result.get("low_confidence"),
            "surya_attempted": surya_attempted,
            "surya_quality_ok": surya_quality_ok if isinstance(surya_quality_ok, bool) else None,
            "table_count": table_count if isinstance(table_count, int) else None,
            "reason": reason if isinstance(reason, str) and reason.strip() else None,
            "error": error if isinstance(error, str) and error.strip() else None,
            "thresholds": thresholds,
            "image_size": image_size,
            "visible_image": visible_image,
            "feature_values": feature_values,
            "score_breakdown": score_breakdown,
            "decision_flags": decision_flags,
            "feature_details": feature_details_payload,
        }
        logger.info(json.dumps(_to_jsonable_copy(debug_payload), ensure_ascii=False, separators=(",", ":")))
