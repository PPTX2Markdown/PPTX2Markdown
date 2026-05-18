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

from converter_models import ShapeBlock, SlideStats


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
    extract_shape_blocks: Callable[[ET.Element], List[ShapeBlock]]
    render_shape_blocks: Callable[[Sequence[ShapeBlock]], str]
    split_triangle_bullets: Callable[[str], List[str]]
    normalize_text: Callable[[str], str]
    shape_id_of: Callable[[ET.Element], str]
    resolve_image_path: Callable[[Dict[str, str], Optional[Path], Optional[str]], Tuple[str, Optional[str]]]
    format_markdown_image: Callable[..., Tuple[str, Optional[str], bool, bool, bool]]
    convert_table_to_markdown: Callable[..., Tuple[Optional[str], Optional[str]]]
    graphic_frame_kind: Callable[[ET.Element], Optional[str]]
    convert_chart_to_markdown: Callable[[ET.Element, Optional[Path], Dict[str, str], Optional[Path]], Tuple[Optional[str], Optional[str]]]
    convert_smartart_to_markdown: Callable[
        [ET.Element, Optional[Path], Dict[str, str], Optional[Path], Optional[Path], Optional[Path]],
        Tuple[Optional[str], Optional[str]],
    ]
    normalize_single_heading_to_h1: Callable[[List[str]], List[str]]


@dataclass
class SlideRenderAssets:
    output_dir: Optional[Path] = None
    media_dir: Optional[Path] = None
    copied_media: Optional[Dict[str, Path]] = None
    image_vlm_provider: str = "local"
    image_vlm_model: Optional[str] = None
    image_vlm_prompt: str = ""
    image_vlm_max_new_tokens: int = 1024
    image_vlm_api_key_env: str = "GEMINI_API_KEY"
    ignore_image_vlm_cache: bool = False


@dataclass
class SlideConversionContext:
    slide_xml: Path
    page_no: int
    ns: Dict[str, str]
    source_slide_xml: Optional[Path] = None
    source_pptx_path: Optional[Path] = None
    rels_path: Optional[Path] = None
    rels_map: Dict[str, str] = field(default_factory=dict)
    heading_hints: Dict[str, Dict[str, object]] = field(default_factory=dict)
    table_overlay_map: Dict[str, List[Dict[str, object]]] = field(default_factory=dict)
    consumed_picture_ids: set[str] = field(default_factory=set)


def _iter_sp_tree_children(sp_tree: ET.Element, deps: SlideConversionDeps) -> List[ET.Element]:
    expanded: List[ET.Element] = []
    for child in list(sp_tree):
        tag = deps.local_name(child.tag)
        if tag != "AlternateContent":
            expanded.append(child)
            continue

        selected = child.find("./{*}Choice")
        if selected is None:
            selected = child.find("./{*}Fallback")
        if selected is None:
            continue
        expanded.extend(list(selected))
    return expanded


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
    heading_policy: HeadingPolicy,
    strict_headings: bool,
    state: SlideRenderState,
    stats: SlideStats,
    context: SlideConversionContext,
    deps: SlideConversionDeps,
) -> None:
    ph = child.find(".//p:ph", context.ns)
    ph_type = ph.attrib.get("type") if ph is not None else None
    if ph_type in {"sldNum", "ftr", "dt"}:
        stats.skipped_blocks += 1
        return

    shape_blocks = deps.extract_shape_blocks(child)
    has_list_semantics = any(block.kind in {"list_ul", "list_ol"} for block in shape_blocks)
    has_math_shape = any(block.has_math for block in shape_blocks)
    text = deps.render_shape_blocks(shape_blocks)
    plain_text = " ".join(block.plain_text for block in shape_blocks if block.plain_text).strip()
    if not text or not plain_text:
        stats.skipped_blocks += 1
        return
    if re.fullmatch(r"\d+", plain_text):
        stats.skipped_blocks += 1
        return

    if has_math_shape:
        stats.math_blocks += 1
    for block in shape_blocks:
        for segment in block.segments:
            if segment.kind == "math_inline":
                stats.inline_math_segments += 1
            elif segment.kind == "math_block":
                stats.block_math_segments += 1
            elif segment.kind == "math_error":
                stats.math_conversion_failures += 1
                stats.warnings.append("OMML to LaTeX conversion failed; used math fallback text")

    sid = deps.shape_id_of(child)
    hint = context.heading_hints.get(sid, {})
    depth = hint.get("heading_depth_hint")
    score = float(hint.get("heading_score", 0.0))
    is_candidate = bool(hint.get("is_heading_candidate", False))
    raw_font_pt = hint.get("font_pt")
    try:
        font_pt = float(raw_font_pt) if raw_font_pt is not None else None
    except (TypeError, ValueError):
        font_pt = None

    rendered_text = re.sub(r"\s+", " ", (text or "").strip())

    strong_heading_signal = False
    if (
        not strict_headings
        and is_candidate
        and isinstance(depth, int)
        and 1 <= depth <= 6
        and score >= heading_policy.threshold
    ):
        strong_heading_signal = True

    if strict_headings:
        strict_ph_type = ph_type if ph_type is not None else hint.get("ph_type")
        strict_depth = hr_strict_heading_depth_from_placeholder(strict_ph_type)
        is_candidate = strict_depth is not None
        depth = strict_depth
        score = 1.0 if is_candidate else 0.0
        if has_math_shape:
            is_candidate = False
            depth = None
            score = 0.0
    else:
        non_strict_ph_type = ph_type if ph_type is not None else hint.get("ph_type")
        non_strict_depth = hr_strict_heading_depth_from_placeholder(non_strict_ph_type)
        if not has_math_shape and non_strict_depth is not None:
            depth = non_strict_depth
            score = max(score, 0.9)
            is_candidate = True
            strong_heading_signal = True
        if has_math_shape:
            is_candidate = False
            depth = None
            score = 0.0
        if not has_math_shape and not is_candidate:
            if has_list_semantics:
                if re.match(r"^\d+\.\d+(?:\.\d+)*\.?\s+", rendered_text):
                    fb_depth = 3
                elif re.match(r"^\d+\.\s+", rendered_text):
                    fb_depth = 2
                else:
                    fb_depth = None
            else:
                fb_depth = None
            if fb_depth is None:
                fb_depth = hr_infer_heading_depth_fallback(rendered_text, state.text_block_index, font_pt=font_pt)
            if fb_depth is None:
                fb_depth = hr_infer_heading_depth_fallback(plain_text, state.text_block_index, font_pt=font_pt)
            if fb_depth is not None:
                depth = fb_depth
                score = max(score, 0.8)
                is_candidate = True
                strong_heading_signal = True

    rendered = text
    if not strict_headings and not strong_heading_signal:
        if has_list_semantics:
            is_candidate = False
        if hr_looks_like_multi_numbered_items(plain_text):
            is_candidate = False
        if hr_is_body_like_long_sentence(plain_text):
            is_candidate = False

    heading_threshold = heading_policy.threshold
    if is_candidate and isinstance(depth, int) and 1 <= depth <= 6 and score >= heading_threshold:
        if strict_headings:
            heading_text = plain_text
        else:
            heading_source = plain_text
            if re.match(r"^\d+(?:\.\d+)*\.?\s+", rendered_text):
                heading_source = rendered_text
            heading_text = hr_clean_heading_text_for_render(heading_source)
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
    state: SlideRenderState,
    stats: SlideStats,
    context: SlideConversionContext,
    assets: SlideRenderAssets,
    deps: SlideConversionDeps,
) -> None:
    sid = deps.shape_id_of(child)
    if sid and sid in context.consumed_picture_ids:
        stats.skipped_blocks += 1
        return

    blip = child.find(".//a:blip", context.ns)
    embed = blip.attrib.get(f"{{{context.ns['r']}}}embed") if blip is not None else None
    img_path, warn = deps.resolve_image_path(context.rels_map, context.rels_path, embed)
    if warn:
        stats.unresolved_images += 1
        stats.warnings.append(warn)
    else:
        stats.resolved_images += 1

    rendered_image, image_warn, unavailable, _, _ = deps.format_markdown_image(
        img_path,
        output_dir=assets.output_dir,
        media_dir=assets.media_dir,
        copied_media=assets.copied_media,
        image_vlm_provider=assets.image_vlm_provider,
        image_vlm_model=assets.image_vlm_model,
        image_vlm_prompt=assets.image_vlm_prompt,
        image_vlm_max_new_tokens=assets.image_vlm_max_new_tokens,
        image_vlm_api_key_env=assets.image_vlm_api_key_env,
        ignore_image_vlm_cache=assets.ignore_image_vlm_cache,
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
    state: SlideRenderState,
    stats: SlideStats,
    context: SlideConversionContext,
    assets: SlideRenderAssets,
    deps: SlideConversionDeps,
) -> None:
    table_md, err = deps.convert_table_to_markdown(
        child,
        overlays=context.table_overlay_map.get(deps.shape_id_of(child), []),
        output_dir=assets.output_dir,
        media_dir=assets.media_dir,
        copied_media=assets.copied_media,
        rels_path=context.rels_path,
        rels_map=context.rels_map,
        image_vlm_provider=assets.image_vlm_provider,
        image_vlm_model=assets.image_vlm_model,
        image_vlm_prompt=assets.image_vlm_prompt,
        image_vlm_max_new_tokens=assets.image_vlm_max_new_tokens,
        image_vlm_api_key_env=assets.image_vlm_api_key_env,
        ignore_image_vlm_cache=assets.ignore_image_vlm_cache,
    )
    if table_md is not None:
        lines.append(table_md.strip())
        lines.append("")
        stats.table_blocks += 1
        return

    gf_kind = deps.graphic_frame_kind(child)
    if gf_kind == "chart":
        chart_md, chart_err = deps.convert_chart_to_markdown(
            child,
            context.rels_path,
            context.rels_map,
            context.source_pptx_path,
        )
        if chart_md is not None:
            lines.append(chart_md.strip())
            lines.append("")
            stats.chart_blocks += 1
            return
        _append_unsupported_graphic_frame(lines, stats)
        if chart_err:
            stats.warnings.append(chart_err)
        return
    elif gf_kind == "diagram":
        smartart_md, smartart_err = deps.convert_smartart_to_markdown(
            child,
            context.rels_path,
            context.rels_map,
            context.source_pptx_path,
            assets.output_dir,
            assets.media_dir,
        )
        if smartart_md:
            lines.append(smartart_md)
            lines.append("")
            stats.smartart_blocks += 1
        else:
            _append_unsupported_graphic_frame(lines, stats)
            stats.warnings.append(smartart_err or "smartart conversion failed")
    else:
        _append_unsupported_graphic_frame(lines, stats)
    if err:
        normalized_err = str(err).lower()
        if "api key not found" in normalized_err or "modulenotfounderror" in normalized_err or "importerror" in normalized_err:
            if not state.image_pipeline_unavailable_reported:
                stats.warnings.append(err)
                state.image_pipeline_unavailable_reported = True
        else:
            stats.warnings.append(err)


def convert_one_slide(
    *,
    context: SlideConversionContext,
    assets: SlideRenderAssets,
    strict_headings: bool = False,
    deps: SlideConversionDeps,
) -> Tuple[str, SlideStats]:
    heading_policy = HeadingPolicy(strict=strict_headings)

    root = ET.parse(context.slide_xml).getroot()
    sp_tree = root.find("p:cSld/p:spTree", context.ns)
    if sp_tree is None:
        raise ValueError("missing p:cSld/p:spTree")

    context.rels_path = deps.choose_rels_in_package(context.slide_xml, source_slide_xml=context.source_slide_xml)
    context.rels_map = deps.build_rels_map(context.rels_path)
    context.heading_hints = deps.load_heading_hints(context.slide_xml)
    (
        context.table_overlay_map,
        context.consumed_picture_ids,
        overlay_warnings,
        overlay_resolved,
        overlay_unresolved,
    ) = deps.collect_table_overlay_pictures(sp_tree, context.slide_xml, context.rels_map, context.rels_path)

    lines: List[str] = [f"[Page_{context.page_no}]", ""]
    state = SlideRenderState()

    stats = SlideStats(
        warnings=list(overlay_warnings),
        rels_path=(str(context.rels_path) if context.rels_path else None),
    )
    stats.resolved_images += overlay_resolved
    stats.unresolved_images += overlay_unresolved

    for child in _iter_sp_tree_children(sp_tree, deps):
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
                heading_policy=heading_policy,
                strict_headings=strict_headings,
                state=state,
                stats=stats,
                context=context,
                deps=deps,
            )
            continue

        if tag == "pic":
            _handle_picture_block(
                child,
                lines=lines,
                state=state,
                stats=stats,
                context=context,
                assets=assets,
                deps=deps,
            )
            continue

        if tag == "graphicFrame":
            _handle_graphic_frame_block(
                child,
                lines=lines,
                state=state,
                stats=stats,
                context=context,
                assets=assets,
                deps=deps,
            )
            continue

    lines = deps.normalize_single_heading_to_h1(lines)
    md_text = "\n".join(lines).rstrip() + "\n"
    return md_text, stats
