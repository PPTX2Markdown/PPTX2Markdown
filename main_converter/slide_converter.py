from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, List, Optional, Sequence, Tuple
import xml.etree.ElementTree as ET

from heading_rules import (
    HeadingPolicy,
    clean_heading_text_for_render as hr_clean_heading_text_for_render,
    infer_heading_depth_fallback as hr_infer_heading_depth_fallback,
    is_body_like_long_sentence as hr_is_body_like_long_sentence,
    looks_like_multi_numbered_items as hr_looks_like_multi_numbered_items,
    strict_heading_depth_from_placeholder as hr_strict_heading_depth_from_placeholder,
)

from converter_models import SlideStats


@dataclass
class SlideRenderState:
    used_headings: set[str] = field(default_factory=set)
    text_block_index: int = 0
    image_pipeline_unavailable_reported: bool = False


@dataclass
class SlideConversionDeps:
    local_name: Callable[[str], str]
    choose_rels_in_package: Callable[[Path, Optional[Path]], Optional[Path]]
    build_rels_map: Callable[[Optional[Path]], Dict[str, str]]
    load_heading_hints: Callable[[Path], Dict[str, Dict[str, object]]]
    collect_table_overlay_pictures: Callable[
        [ET.Element, Path, Dict[str, str], Optional[Path]],
        Tuple[Dict[str, List[Dict[str, object]]], set[str], List[str], int, int],
    ]
    extract_shape_blocks: Callable[[ET.Element], List[Tuple[str, str, Optional[int]]]]
    render_shape_blocks: Callable[[Sequence[Tuple[str, str, Optional[int]]]], str]
    split_triangle_bullets: Callable[[str], List[str]]
    normalize_text: Callable[[str], str]
    shape_id_of: Callable[[ET.Element], str]
    resolve_image_path: Callable[[Dict[str, str], Optional[Path], Optional[str]], Tuple[str, Optional[str]]]
    convert_picture_to_table_markdown: Callable[
        [str],
        Tuple[Optional[str], Optional[str], bool, Optional[Dict[str, object]]],
    ]
    copy_debug_image_asset: Callable[[str, Optional[Path], Optional[Dict[str, Path]]], Optional[str]]
    format_markdown_image: Callable[..., Tuple[str, Optional[str], bool, bool, bool]]
    convert_table_to_markdown: Callable[..., Tuple[Optional[str], Optional[str]]]
    graphic_frame_kind: Callable[[ET.Element], Optional[str]]
    diagram_data_path: Callable[[ET.Element, Optional[Path], Dict[str, str]], Optional[Path]]
    extract_diagram_texts: Callable[[Path], List[str]]
    format_diagram_as_markdown: Callable[[Sequence[str]], Optional[str]]
    normalize_single_heading_to_h1: Callable[[List[str]], List[str]]


def _append_rendered_text_block(lines: List[str], rendered: str, deps: SlideConversionDeps) -> None:
    if rendered.startswith("#"):
        lines.append(rendered)
        lines.append("")
        return
    rendered_lines = deps.split_triangle_bullets(rendered)
    if rendered_lines:
        lines.extend(rendered_lines)
        lines.append("")
        return
    lines.append(rendered)
    lines.append("")


def _handle_text_shape_block(
    child: ET.Element,
    *,
    lines: List[str],
    heading_hints: Dict[str, Dict[str, object]],
    heading_policy: HeadingPolicy,
    strict_headings: bool,
    state: SlideRenderState,
    stats: SlideStats,
    deps: SlideConversionDeps,
    ns: Dict[str, str],
) -> None:
    ph = child.find(".//p:ph", ns)
    ph_type = ph.attrib.get("type") if ph is not None else None
    if ph_type in {"sldNum", "ftr", "dt"}:
        stats.skipped_blocks += 1
        return

    shape_blocks = deps.extract_shape_blocks(child)
    has_list_semantics = any(kind in {"list_ul", "list_ol"} for kind, _, _ in shape_blocks)
    text = deps.render_shape_blocks(shape_blocks)
    if not text:
        stats.skipped_blocks += 1
        return
    if re.fullmatch(r"\d+", text):
        stats.skipped_blocks += 1
        return

    sid = deps.shape_id_of(child)
    hint = heading_hints.get(sid, {})
    depth = hint.get("heading_depth_hint")
    score = float(hint.get("heading_score", 0.0))
    is_candidate = bool(hint.get("is_heading_candidate", False))
    raw_font_pt = hint.get("font_pt")
    try:
        font_pt = float(raw_font_pt) if raw_font_pt is not None else None
    except (TypeError, ValueError):
        font_pt = None

    if strict_headings:
        strict_ph_type = ph_type if ph_type is not None else hint.get("ph_type")
        strict_depth = hr_strict_heading_depth_from_placeholder(strict_ph_type)
        is_candidate = strict_depth is not None
        depth = strict_depth
        score = 1.0 if is_candidate else 0.0
    else:
        non_strict_ph_type = ph_type if ph_type is not None else hint.get("ph_type")
        non_strict_depth = hr_strict_heading_depth_from_placeholder(non_strict_ph_type)
        if non_strict_depth is not None:
            depth = non_strict_depth
            score = max(score, 0.9)
            is_candidate = True
        if not is_candidate:
            fb_depth = hr_infer_heading_depth_fallback(text, state.text_block_index, font_pt=font_pt)
            if fb_depth is not None:
                depth = fb_depth
                score = 0.8
                is_candidate = True

    rendered = text
    if not strict_headings:
        if has_list_semantics:
            is_candidate = False
        if hr_looks_like_multi_numbered_items(rendered):
            is_candidate = False
        if hr_is_body_like_long_sentence(rendered):
            is_candidate = False

    heading_threshold = heading_policy.threshold
    if is_candidate and isinstance(depth, int) and 1 <= depth <= 6 and score >= heading_threshold:
        heading_text = text if strict_headings else hr_clean_heading_text_for_render(text)
        key = deps.normalize_text(heading_text)
        if key not in state.used_headings:
            rendered = f"{'#' * depth} {heading_text}"
            state.used_headings.add(key)
        else:
            stats.skipped_blocks += 1
            return

    _append_rendered_text_block(lines, rendered, deps)
    stats.text_blocks += 1
    state.text_block_index += 1


def _handle_picture_block(
    child: ET.Element,
    *,
    lines: List[str],
    rels_map: Dict[str, str],
    rels_path: Optional[Path],
    consumed_picture_ids: set[str],
    output_dir: Optional[Path],
    media_dir: Optional[Path],
    copied_media: Optional[Dict[str, Path]],
    image_vlm_provider: str,
    image_vlm_model: Optional[str],
    image_vlm_prompt: str,
    image_vlm_max_new_tokens: int,
    image_vlm_api_key_env: str,
    surya_debug_dir: Optional[Path],
    copied_surya_debug_images: Optional[Dict[str, Path]],
    enable_image_table_pipeline: bool,
    state: SlideRenderState,
    stats: SlideStats,
    deps: SlideConversionDeps,
    ns: Dict[str, str],
) -> None:
    sid = deps.shape_id_of(child)
    if sid and sid in consumed_picture_ids:
        stats.skipped_blocks += 1
        return

    blip = child.find(".//a:blip", ns)
    embed = blip.attrib.get(f"{{{ns['r']}}}embed") if blip is not None else None
    img_path, warn = deps.resolve_image_path(rels_map, rels_path, embed)
    if warn:
        stats.unresolved_images += 1
        stats.warnings.append(warn)
    else:
        stats.resolved_images += 1

    if enable_image_table_pipeline and not warn and not img_path.startswith("[unresolved-image"):
        table_md, table_warn, unavailable, table_result = deps.convert_picture_to_table_markdown(img_path)
        if isinstance(table_result, dict) and (
            bool(table_result.get("surya_attempted")) or str(table_result.get("status", "")) == "table_skipped"
        ):
            deps.copy_debug_image_asset(
                img_path,
                debug_dir=surya_debug_dir,
                copied_debug_images=copied_surya_debug_images,
            )
        if table_md is not None:
            lines.append(table_md.strip())
            lines.append("")
            stats.table_blocks += 1
            return
        if isinstance(table_result, dict) and str(table_result.get("status", "")) == "table_skipped":
            stats.table_skipped_blocks += 1
        if table_warn:
            if unavailable:
                if not state.image_pipeline_unavailable_reported:
                    stats.warnings.append(table_warn)
                    state.image_pipeline_unavailable_reported = True
            else:
                stats.warnings.append(table_warn)

    rendered_image, image_warn, unavailable, _, _ = deps.format_markdown_image(
        img_path,
        output_dir=output_dir,
        media_dir=media_dir,
        copied_media=copied_media,
        image_vlm_provider=image_vlm_provider,
        image_vlm_model=image_vlm_model,
        image_vlm_prompt=image_vlm_prompt,
        image_vlm_max_new_tokens=image_vlm_max_new_tokens,
        image_vlm_api_key_env=image_vlm_api_key_env,
    )
    if image_warn:
        if unavailable:
            if not state.image_pipeline_unavailable_reported:
                stats.warnings.append(image_warn)
                state.image_pipeline_unavailable_reported = True
        else:
            stats.warnings.append(image_warn)

    lines.append(rendered_image)
    lines.append("")
    stats.image_blocks += 1


def _append_unsupported_graphic_frame(lines: List[str], stats: SlideStats) -> None:
    lines.append("[unsupported: graphicFrame(non-table)]")
    lines.append("")
    stats.unsupported_blocks += 1


def _handle_graphic_frame_block(
    child: ET.Element,
    *,
    lines: List[str],
    table_overlay_map: Dict[str, List[Dict[str, object]]],
    rels_path: Optional[Path],
    rels_map: Dict[str, str],
    output_dir: Optional[Path],
    media_dir: Optional[Path],
    copied_media: Optional[Dict[str, Path]],
    stats: SlideStats,
    deps: SlideConversionDeps,
) -> None:
    table_md, err = deps.convert_table_to_markdown(
        child,
        overlays=table_overlay_map.get(deps.shape_id_of(child), []),
        output_dir=output_dir,
        media_dir=media_dir,
        copied_media=copied_media,
    )
    if table_md is not None:
        lines.append(table_md.strip())
        lines.append("")
        stats.table_blocks += 1
        return

    gf_kind = deps.graphic_frame_kind(child)
    if gf_kind == "diagram":
        diagram_path = deps.diagram_data_path(child, rels_path, rels_map)
        diagram_text = deps.format_diagram_as_markdown(
            deps.extract_diagram_texts(diagram_path) if diagram_path else []
        )
        if diagram_text:
            lines.append(diagram_text)
            lines.append("")
            stats.text_blocks += 1
        else:
            _append_unsupported_graphic_frame(lines, stats)
            stats.warnings.append("diagram text extraction failed")
    else:
        _append_unsupported_graphic_frame(lines, stats)
    if err:
        stats.warnings.append(err)


def convert_one_slide(
    slide_xml: Path,
    page_no: int,
    *,
    source_slide_xml: Optional[Path] = None,
    output_dir: Optional[Path] = None,
    media_dir: Optional[Path] = None,
    copied_media: Optional[Dict[str, Path]] = None,
    image_vlm_provider: str = "local",
    image_vlm_model: Optional[str] = None,
    image_vlm_prompt: str = "",
    image_vlm_max_new_tokens: int = 1024,
    image_vlm_api_key_env: str = "GEMINI_API_KEY",
    surya_debug_dir: Optional[Path] = None,
    copied_surya_debug_images: Optional[Dict[str, Path]] = None,
    enable_image_table_pipeline: bool = False,
    strict_headings: bool = False,
    deps: SlideConversionDeps,
    ns: Dict[str, str],
) -> Tuple[str, SlideStats]:
    heading_policy = HeadingPolicy(strict=strict_headings)
    root = ET.parse(slide_xml).getroot()
    sp_tree = root.find("p:cSld/p:spTree", ns)
    if sp_tree is None:
        raise ValueError("missing p:cSld/p:spTree")

    rels_path = deps.choose_rels_in_package(slide_xml, source_slide_xml=source_slide_xml)
    rels_map = deps.build_rels_map(rels_path)
    heading_hints = deps.load_heading_hints(slide_xml)
    table_overlay_map, consumed_picture_ids, overlay_warnings, overlay_resolved, overlay_unresolved = (
        deps.collect_table_overlay_pictures(sp_tree, slide_xml, rels_map, rels_path)
    )

    lines: List[str] = [f"[Page_{page_no}]", ""]
    state = SlideRenderState()

    stats = SlideStats(
        warnings=list(overlay_warnings),
        rels_path=(str(rels_path) if rels_path else None),
    )
    stats.resolved_images += overlay_resolved
    stats.unresolved_images += overlay_unresolved

    for child in list(sp_tree):
        tag = deps.local_name(child.tag)
        if tag not in {"sp", "pic", "graphicFrame", "grpSp", "cxnSp"}:
            continue
        stats.blocks_total += 1

        if tag == "cxnSp":
            stats.skipped_blocks += 1
            continue

        if tag in {"sp", "grpSp"}:
            _handle_text_shape_block(
                child,
                lines=lines,
                heading_hints=heading_hints,
                heading_policy=heading_policy,
                strict_headings=strict_headings,
                state=state,
                stats=stats,
                deps=deps,
                ns=ns,
            )
            continue

        if tag == "pic":
            _handle_picture_block(
                child,
                lines=lines,
                rels_map=rels_map,
                rels_path=rels_path,
                consumed_picture_ids=consumed_picture_ids,
                output_dir=output_dir,
                media_dir=media_dir,
                copied_media=copied_media,
                image_vlm_provider=image_vlm_provider,
                image_vlm_model=image_vlm_model,
                image_vlm_prompt=image_vlm_prompt,
                image_vlm_max_new_tokens=image_vlm_max_new_tokens,
                image_vlm_api_key_env=image_vlm_api_key_env,
                surya_debug_dir=surya_debug_dir,
                copied_surya_debug_images=copied_surya_debug_images,
                enable_image_table_pipeline=enable_image_table_pipeline,
                state=state,
                stats=stats,
                deps=deps,
                ns=ns,
            )
            continue

        if tag == "graphicFrame":
            _handle_graphic_frame_block(
                child,
                lines=lines,
                table_overlay_map=table_overlay_map,
                rels_path=rels_path,
                rels_map=rels_map,
                output_dir=output_dir,
                media_dir=media_dir,
                copied_media=copied_media,
                stats=stats,
                deps=deps,
            )
            continue

    lines = deps.normalize_single_heading_to_h1(lines)
    md_text = "\n".join(lines).rstrip() + "\n"
    return md_text, stats
