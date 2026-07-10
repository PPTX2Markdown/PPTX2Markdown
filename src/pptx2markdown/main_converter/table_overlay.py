from __future__ import annotations

import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Callable, Dict, List, Optional, Sequence, Tuple


def shape_id_of(elem: ET.Element, ns: Dict[str, str]) -> str:
    c_nv_pr = elem.find(".//p:cNvPr", ns)
    if c_nv_pr is None:
        return ""
    return c_nv_pr.attrib.get("id", "")


def parse_int(v: Optional[str], default: int = 10**18) -> int:
    if v is None:
        return default
    try:
        return int(v)
    except ValueError:
        return default


def first_off(elem: ET.Element, ns: Dict[str, str]) -> Optional[ET.Element]:
    for p in (
        "./p:spPr/a:xfrm/a:off",
        "./p:grpSpPr/a:xfrm/a:off",
        "./p:xfrm/a:off",
        ".//a:off",
    ):
        off = elem.find(p, ns)
        if off is not None:
            return off
    return None


def first_ext(elem: ET.Element, ns: Dict[str, str]) -> Optional[ET.Element]:
    for p in (
        "./p:spPr/a:xfrm/a:ext",
        "./p:grpSpPr/a:xfrm/a:ext",
        "./p:xfrm/a:ext",
        ".//a:ext",
    ):
        ext = elem.find(p, ns)
        if ext is not None:
            return ext
    return None


def extract_bbox_emu(elem: ET.Element, ns: Dict[str, str]) -> Optional[Tuple[int, int, int, int]]:
    off = first_off(elem, ns)
    ext = first_ext(elem, ns)
    if off is None or ext is None:
        return None
    x = parse_int(off.attrib.get("x"))
    y = parse_int(off.attrib.get("y"))
    w = parse_int(ext.attrib.get("cx"))
    h = parse_int(ext.attrib.get("cy"))
    if any(v >= 10**18 for v in (x, y, w, h)) or w <= 0 or h <= 0:
        return None
    return (x, y, x + w, y + h)


def intersection_area(a: Tuple[int, int, int, int], b: Tuple[int, int, int, int]) -> int:
    left = max(a[0], b[0])
    top = max(a[1], b[1])
    right = min(a[2], b[2])
    bottom = min(a[3], b[3])
    if right <= left or bottom <= top:
        return 0
    return (right - left) * (bottom - top)


def bbox_area(bbox: Tuple[int, int, int, int]) -> int:
    return max(0, bbox[2] - bbox[0]) * max(0, bbox[3] - bbox[1])


def bbox_center(bbox: Tuple[int, int, int, int]) -> Tuple[int, int]:
    return ((bbox[0] + bbox[2]) // 2, (bbox[1] + bbox[3]) // 2)


def bbox_contains_point(bbox: Tuple[int, int, int, int], point: Tuple[int, int]) -> bool:
    return bbox[0] <= point[0] <= bbox[2] and bbox[1] <= point[1] <= bbox[3]


def collect_table_overlay_pictures(
    shape_items: Sequence[Dict[str, object]],
    rels_map: Dict[str, str],
    rels_path: Optional[Path],
    *,
    ns: Dict[str, str],
    resolve_image_path_fn: Callable[
        [Dict[str, str], Optional[Path], Optional[str]], Tuple[str, Optional[str]]
    ],
) -> Tuple[Dict[str, List[Dict[str, object]]], set[str], List[str], int, int]:
    table_bboxes: List[Tuple[str, Tuple[int, int, int, int]]] = []
    picture_infos: List[Dict[str, object]] = []

    for item in shape_items:
        child = item.get("elem")
        tag = str(item.get("tag", ""))
        bbox = item.get("bbox")
        sid = str(item.get("shape_id", ""))
        if child is None or not isinstance(bbox, tuple) or len(bbox) != 4:
            continue
        if tag == "graphicFrame":
            tbl = child.find(".//a:tbl", ns)
            if tbl is not None and sid:
                table_bboxes.append((sid, bbox))
        elif tag == "pic":
            if not sid:
                continue
            blip = child.find(".//a:blip", ns)
            embed = blip.attrib.get(f"{{{ns['r']}}}embed") if blip is not None else None
            path, warn = resolve_image_path_fn(rels_map, rels_path, embed)
            picture_infos.append({"shape_id": sid, "bbox": bbox, "path": path, "warn": warn})

    by_table: Dict[str, List[Dict[str, object]]] = {}
    consumed: set[str] = set()
    warnings: List[str] = []
    resolved = 0
    unresolved = 0
    for table_id, table_bbox in table_bboxes:
        table_area = bbox_area(table_bbox)
        if table_area <= 0:
            continue
        for pic_info in picture_infos:
            pic_id = str(pic_info["shape_id"])
            pic_bbox = pic_info["bbox"]
            pic_area = bbox_area(pic_bbox)
            if pic_area <= 0:
                continue
            overlap = intersection_area(table_bbox, pic_bbox)
            overlap_ratio = overlap / pic_area
            center_inside = bbox_contains_point(table_bbox, bbox_center(pic_bbox))
            if center_inside and overlap_ratio >= 0.8:
                consumed.add(pic_id)
                by_table.setdefault(table_id, []).append(pic_info)
                warn = pic_info.get("warn")
                if isinstance(warn, str) and warn:
                    unresolved += 1
                    warnings.append(warn)
                else:
                    resolved += 1
    return by_table, consumed, warnings, resolved, unresolved


def compute_table_cell_bounds(
    graphic_frame: ET.Element,
    ns: Dict[str, str],
) -> Optional[Tuple[List[Tuple[int, int]], List[Tuple[int, int]]]]:
    bbox = extract_bbox_emu(graphic_frame, ns)
    tbl = graphic_frame.find(".//a:tbl", ns)
    if bbox is None or tbl is None:
        return None

    col_elems = tbl.findall("./a:tblGrid/a:gridCol", ns)
    row_elems = tbl.findall("./a:tr", ns)
    col_widths = [parse_int(col.attrib.get("w"), 0) for col in col_elems]
    row_heights = [parse_int(row.attrib.get("h"), 0) for row in row_elems]
    total_col = sum(v for v in col_widths if v > 0)
    total_row = sum(v for v in row_heights if v > 0)
    if total_col <= 0 or total_row <= 0:
        return None

    table_width = bbox[2] - bbox[0]
    table_height = bbox[3] - bbox[1]
    if table_width <= 0 or table_height <= 0:
        return None

    col_bounds: List[Tuple[int, int]] = []
    cursor = bbox[0]
    used_width = 0
    for idx, width in enumerate(col_widths):
        if idx == len(col_widths) - 1:
            right = bbox[2]
        else:
            used_width += width
            right = bbox[0] + int(table_width * used_width / total_col)
        col_bounds.append((cursor, right))
        cursor = right

    row_bounds: List[Tuple[int, int]] = []
    cursor = bbox[1]
    used_height = 0
    for idx, height in enumerate(row_heights):
        if idx == len(row_heights) - 1:
            bottom = bbox[3]
        else:
            used_height += height
            bottom = bbox[1] + int(table_height * used_height / total_row)
        row_bounds.append((cursor, bottom))
        cursor = bottom

    return col_bounds, row_bounds


def find_table_cell_origin(
    parsed_table: Dict[str, object], row_idx: int, col_idx: int
) -> Tuple[int, int]:
    rows = parsed_table.get("rows")
    if not isinstance(rows, list):
        return row_idx, col_idx
    if row_idx >= len(rows):
        return row_idx, col_idx
    row = rows[row_idx]
    if not isinstance(row, list) or col_idx >= len(row):
        return row_idx, col_idx
    cell = row[col_idx]
    if not isinstance(cell, dict):
        return row_idx, col_idx
    origin = cell.get("origin")
    if cell.get("type") in {"hMerge", "vMerge"} and isinstance(origin, list) and len(origin) == 2:
        if isinstance(origin[0], int) and isinstance(origin[1], int):
            return origin[0], origin[1]
    return row_idx, col_idx


def inject_table_overlay_links(
    parsed_table: Dict[str, object],
    graphic_frame: ET.Element,
    overlays: Sequence[Dict[str, object]],
    *,
    ns: Dict[str, str],
    normalize_text_fn: Callable[[str], str],
    overlay_content_text_fn: Callable[..., Tuple[str, Optional[str], bool, bool, bool]],
    output_dir: Optional[Path],
    media_dir: Optional[Path],
    copied_media: Optional[Dict[str, Path]],
    image_vlm_provider: str,
    image_vlm_model: Optional[str],
    image_vlm_prompt: str,
    image_vlm_max_new_tokens: int,
    image_vlm_api_key_env: str,
    ignore_image_vlm_cache: bool,
) -> Tuple[Dict[str, object], List[str], bool]:
    if not overlays:
        return parsed_table, [], False

    return _inject_table_images(
        parsed_table,
        graphic_frame,
        image_entries=overlays,
        ns=ns,
        normalize_text_fn=normalize_text_fn,
        overlay_content_text_fn=overlay_content_text_fn,
        output_dir=output_dir,
        media_dir=media_dir,
        copied_media=copied_media,
        image_vlm_provider=image_vlm_provider,
        image_vlm_model=image_vlm_model,
        image_vlm_prompt=image_vlm_prompt,
        image_vlm_max_new_tokens=image_vlm_max_new_tokens,
        image_vlm_api_key_env=image_vlm_api_key_env,
        ignore_image_vlm_cache=ignore_image_vlm_cache,
    )


def collect_table_cell_fill_images(
    graphic_frame: ET.Element,
    rels_map: Dict[str, str],
    rels_path: Optional[Path],
    *,
    ns: Dict[str, str],
    resolve_image_path_fn: Callable[
        [Dict[str, str], Optional[Path], Optional[str]], Tuple[str, Optional[str]]
    ],
) -> Tuple[List[Dict[str, object]], List[str], int, int]:
    tbl = graphic_frame.find(".//a:tbl", ns)
    if tbl is None:
        return [], [], 0, 0

    items: List[Dict[str, object]] = []
    warnings: List[str] = []
    resolved = 0
    unresolved = 0

    for row_idx, tr in enumerate(tbl.findall("./a:tr", ns)):
        for col_idx, tc in enumerate(tr.findall("./a:tc", ns)):
            tcpr = tc.find("./a:tcPr", ns)
            if tcpr is None:
                continue
            blip = tcpr.find(".//a:blip", ns)
            if blip is None:
                continue
            embed = blip.attrib.get(f"{{{ns['r']}}}embed")
            path, warn = resolve_image_path_fn(rels_map, rels_path, embed)
            item = {
                "row_idx": row_idx,
                "col_idx": col_idx,
                "path": path,
                "warn": warn,
                "kind": "cell_fill",
            }
            items.append(item)
            if isinstance(warn, str) and warn:
                unresolved += 1
                warnings.append(warn)
            else:
                resolved += 1
    return items, warnings, resolved, unresolved


def _table_cell_text(content: str) -> str:
    lines = [line.strip() for line in str(content).splitlines() if line.strip()]
    return "<br>".join(lines)


def _append_cell_content(
    cell: Dict[str, object], content: str, normalize_text_fn: Callable[[str], str]
) -> None:
    if not content:
        return
    existing_raw = str(cell.get("text", ""))
    existing = normalize_text_fn(existing_raw)
    updated = f"{existing} <br> {content}".strip() if existing else content
    cell["text"] = updated


def _inject_table_images(
    parsed_table: Dict[str, object],
    graphic_frame: ET.Element,
    image_entries: Sequence[Dict[str, object]],
    *,
    ns: Dict[str, str],
    normalize_text_fn: Callable[[str], str],
    overlay_content_text_fn: Callable[..., Tuple[str, Optional[str], bool, bool, bool]],
    output_dir: Optional[Path],
    media_dir: Optional[Path],
    copied_media: Optional[Dict[str, Path]],
    image_vlm_provider: str,
    image_vlm_model: Optional[str],
    image_vlm_prompt: str,
    image_vlm_max_new_tokens: int,
    image_vlm_api_key_env: str,
    ignore_image_vlm_cache: bool,
) -> Tuple[Dict[str, object], List[str], bool]:
    bounds = compute_table_cell_bounds(graphic_frame, ns)
    if bounds is None:
        return parsed_table, [], False
    col_bounds, row_bounds = bounds

    rows = parsed_table.get("rows")
    if not isinstance(rows, list):
        return parsed_table, [], False

    warnings: List[str] = []
    unavailable_reported = False

    for image_entry in image_entries:
        path = str(image_entry.get("path", ""))
        if not path:
            continue

        row_idx = image_entry.get("row_idx")
        col_idx = image_entry.get("col_idx")
        if not isinstance(row_idx, int) or not isinstance(col_idx, int):
            bbox = image_entry.get("bbox")
            if not (isinstance(bbox, tuple) and len(bbox) == 4):
                continue
            center = bbox_center(bbox)
            row_idx = next(
                (
                    idx
                    for idx, (top, bottom) in enumerate(row_bounds)
                    if top <= center[1] <= bottom
                ),
                None,
            )
            col_idx = next(
                (
                    idx
                    for idx, (left, right) in enumerate(col_bounds)
                    if left <= center[0] <= right
                ),
                None,
            )
        if row_idx is None or col_idx is None:
            continue
        origin_row, origin_col = find_table_cell_origin(parsed_table, row_idx, col_idx)
        if origin_row >= len(rows):
            continue
        row = rows[origin_row]
        if not isinstance(row, list) or origin_col >= len(row):
            continue
        cell = row[origin_col]
        if not isinstance(cell, dict):
            continue

        rendered, warn, unavailable, _, _ = overlay_content_text_fn(
            path,
            output_dir,
            media_dir=media_dir,
            copied_media=copied_media,
            image_vlm_provider=image_vlm_provider,
            image_vlm_model=image_vlm_model,
            image_vlm_prompt=image_vlm_prompt,
            image_vlm_max_new_tokens=image_vlm_max_new_tokens,
            image_vlm_api_key_env=image_vlm_api_key_env,
            ignore_image_vlm_cache=ignore_image_vlm_cache,
        )
        _append_cell_content(cell, _table_cell_text(rendered), normalize_text_fn)
        if warn:
            if unavailable:
                if not unavailable_reported:
                    warnings.append(warn)
                    unavailable_reported = True
            else:
                warnings.append(warn)
    return parsed_table, warnings, unavailable_reported


def convert_table_to_markdown(
    graphic_frame: ET.Element,
    overlays: Optional[Sequence[Dict[str, object]]] = None,
    output_dir: Optional[Path] = None,
    media_dir: Optional[Path] = None,
    copied_media: Optional[Dict[str, Path]] = None,
    rels_path: Optional[Path] = None,
    rels_map: Optional[Dict[str, str]] = None,
    image_vlm_provider: str = "local",
    image_vlm_model: Optional[str] = None,
    image_vlm_prompt: str = "",
    image_vlm_max_new_tokens: int = 1024,
    image_vlm_api_key_env: str = "GEMINI_API_KEY",
    ignore_image_vlm_cache: bool = False,
    *,
    ns: Dict[str, str],
    normalize_text_fn: Callable[[str], str],
    overlay_content_text_fn: Callable[..., Tuple[str, Optional[str], bool, bool, bool]],
    resolve_image_path_fn: Callable[
        [Dict[str, str], Optional[Path], Optional[str]], Tuple[str, Optional[str]]
    ],
) -> Tuple[Optional[str], Optional[str]]:
    tbl = graphic_frame.find(".//a:tbl", ns)
    if tbl is None:
        return None, "graphicFrame without a:tbl"

    from pptx2markdown.table_pipeline import parse as table_parse  # type: ignore
    from pptx2markdown.table_pipeline import render as table_render  # type: ignore

    parsed = table_parse.parse_table_element(tbl, source="<slide_table>")
    parsed, overlay_warnings, overlay_unavailable = inject_table_overlay_links(
        parsed,
        graphic_frame,
        overlays or [],
        ns=ns,
        normalize_text_fn=normalize_text_fn,
        overlay_content_text_fn=overlay_content_text_fn,
        output_dir=output_dir,
        media_dir=media_dir,
        copied_media=copied_media,
        image_vlm_provider=image_vlm_provider,
        image_vlm_model=image_vlm_model,
        image_vlm_prompt=image_vlm_prompt,
        image_vlm_max_new_tokens=image_vlm_max_new_tokens,
        image_vlm_api_key_env=image_vlm_api_key_env,
        ignore_image_vlm_cache=ignore_image_vlm_cache,
    )
    cell_fill_items, cell_fill_warnings, _, _ = collect_table_cell_fill_images(
        graphic_frame,
        rels_map or {},
        rels_path,
        ns=ns,
        resolve_image_path_fn=resolve_image_path_fn,
    )
    parsed, cell_image_warnings, cell_image_unavailable = _inject_table_images(
        parsed,
        graphic_frame,
        image_entries=cell_fill_items,
        ns=ns,
        normalize_text_fn=normalize_text_fn,
        overlay_content_text_fn=overlay_content_text_fn,
        output_dir=output_dir,
        media_dir=media_dir,
        copied_media=copied_media,
        image_vlm_provider=image_vlm_provider,
        image_vlm_model=image_vlm_model,
        image_vlm_prompt=image_vlm_prompt,
        image_vlm_max_new_tokens=image_vlm_max_new_tokens,
        image_vlm_api_key_env=image_vlm_api_key_env,
        ignore_image_vlm_cache=ignore_image_vlm_cache,
    )
    md = table_render.render_parsed_table_to_markdown(
        parsed_table=parsed,
        header_rows=1,
        fill_merged=table_render.FILL_BOTH,
    )
    warnings = list(cell_fill_warnings) + list(overlay_warnings) + list(cell_image_warnings)
    if warnings:
        if overlay_unavailable or cell_image_unavailable:
            return md, warnings[0]
        return md, "; ".join(warnings)
    return md, None
