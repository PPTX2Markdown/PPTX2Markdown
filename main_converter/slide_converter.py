from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Sequence, Tuple
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


_DRAWABLE_TAGS = {"sp", "pic", "graphicFrame", "grpSp", "cxnSp"}
_LEAF_DRAWABLE_TAGS = {"sp", "pic", "graphicFrame", "cxnSp"}
_LARGE_INT = 10**18


@dataclass(frozen=True)
class _GroupTransform:
    sx: float = 1.0
    sy: float = 1.0
    tx: float = 0.0
    ty: float = 0.0

    def apply_bbox(self, bbox: Tuple[int, int, int, int]) -> Tuple[int, int, int, int]:
        x1, y1, x2, y2 = bbox
        return (
            int(round(self.sx * x1 + self.tx)),
            int(round(self.sy * y1 + self.ty)),
            int(round(self.sx * x2 + self.tx)),
            int(round(self.sy * y2 + self.ty)),
        )

    def compose(self, inner: "_GroupTransform") -> "_GroupTransform":
        return _GroupTransform(
            sx=self.sx * inner.sx,
            sy=self.sy * inner.sy,
            tx=self.sx * inner.tx + self.tx,
            ty=self.sy * inner.ty + self.ty,
        )


@dataclass(frozen=True)
class FlattenedShape:
    elem: ET.Element
    tag: str
    shape_id: str
    name: str
    group_path: Tuple[str, ...]
    z_path: Tuple[int, ...]
    bbox: Optional[Tuple[int, int, int, int]]


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
        [Sequence[Dict[str, object]], Path, Dict[str, str], Optional[Path]],
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


def _local_name(tag: str) -> str:
    return tag.split("}", 1)[-1]


def _parse_int(value: Optional[str], default: int = _LARGE_INT) -> int:
    if value is None:
        return default
    try:
        return int(value)
    except ValueError:
        return default


def _first(elem: ET.Element, paths: Tuple[str, ...], ns: Dict[str, str]) -> Optional[ET.Element]:
    for path in paths:
        found = elem.find(path, ns)
        if found is not None:
            return found
    return None


def _point_attrs(node: Optional[ET.Element], x_name: str, y_name: str) -> Optional[Tuple[int, int]]:
    if node is None:
        return None
    x = _parse_int(node.attrib.get(x_name))
    y = _parse_int(node.attrib.get(y_name))
    if x >= _LARGE_INT or y >= _LARGE_INT:
        return None
    return x, y


def _shape_id_name(elem: ET.Element, tag: str, ns: Dict[str, str]) -> Tuple[str, str]:
    paths = {
        "sp": "./p:nvSpPr/p:cNvPr",
        "pic": "./p:nvPicPr/p:cNvPr",
        "graphicFrame": "./p:nvGraphicFramePr/p:cNvPr",
        "grpSp": "./p:nvGrpSpPr/p:cNvPr",
        "cxnSp": "./p:nvCxnSpPr/p:cNvPr",
    }
    c_nv_pr = elem.find(paths.get(tag, ".//p:cNvPr"), ns)
    if c_nv_pr is None:
        return "", ""
    return c_nv_pr.attrib.get("id", ""), c_nv_pr.attrib.get("name", "")


def _extract_local_bbox_emu(elem: ET.Element, ns: Dict[str, str]) -> Optional[Tuple[int, int, int, int]]:
    off = _first(
        elem,
        (
            "./p:spPr/a:xfrm/a:off",
            "./p:grpSpPr/a:xfrm/a:off",
            "./p:xfrm/a:off",
        ),
        ns,
    )
    ext = _first(
        elem,
        (
            "./p:spPr/a:xfrm/a:ext",
            "./p:grpSpPr/a:xfrm/a:ext",
            "./p:xfrm/a:ext",
        ),
        ns,
    )
    origin = _point_attrs(off, "x", "y")
    size = _point_attrs(ext, "cx", "cy")
    if origin is None or size is None:
        return None
    x, y = origin
    w, h = size
    if w <= 0 or h <= 0:
        return None
    return (x, y, x + w, y + h)


def _group_child_transform(group: ET.Element, ns: Dict[str, str]) -> _GroupTransform:
    xfrm = group.find("./p:grpSpPr/a:xfrm", ns)
    if xfrm is None:
        return _GroupTransform()

    off = _point_attrs(xfrm.find("./a:off", ns), "x", "y")
    ext = _point_attrs(xfrm.find("./a:ext", ns), "cx", "cy")
    ch_off = _point_attrs(xfrm.find("./a:chOff", ns), "x", "y")
    ch_ext = _point_attrs(xfrm.find("./a:chExt", ns), "cx", "cy")
    if off is None or ext is None or ch_off is None or ch_ext is None:
        return _GroupTransform()

    off_x, off_y = off
    ext_x, ext_y = ext
    ch_off_x, ch_off_y = ch_off
    ch_ext_x, ch_ext_y = ch_ext
    if ch_ext_x <= 0 or ch_ext_y <= 0:
        return _GroupTransform()

    sx = ext_x / ch_ext_x
    sy = ext_y / ch_ext_y
    return _GroupTransform(
        sx=sx,
        sy=sy,
        tx=off_x - sx * ch_off_x,
        ty=off_y - sy * ch_off_y,
    )


def _expanded_children(elem: ET.Element) -> Iterable[ET.Element]:
    for child in list(elem):
        tag = _local_name(child.tag)
        if tag != "AlternateContent":
            yield child
            continue

        selected = child.find("./{*}Choice")
        if selected is None:
            selected = child.find("./{*}Fallback")
        if selected is None:
            continue
        yield from list(selected)


def _iter_flattened_shapes(
    container: ET.Element,
    ns: Dict[str, str],
    *,
    include_groups: bool = False,
) -> Iterable[FlattenedShape]:
    def walk(
        elem: ET.Element,
        transform: _GroupTransform,
        group_path: Tuple[str, ...],
        z_prefix: Tuple[int, ...],
    ) -> Iterable[FlattenedShape]:
        ordinal = 0
        for child in _expanded_children(elem):
            tag = _local_name(child.tag)
            if tag not in _DRAWABLE_TAGS:
                continue

            ordinal += 1
            z_path = z_prefix + (ordinal,)
            shape_id, name = _shape_id_name(child, tag, ns)
            local_bbox = _extract_local_bbox_emu(child, ns)
            bbox = transform.apply_bbox(local_bbox) if local_bbox is not None else None

            if tag == "grpSp":
                if include_groups:
                    yield FlattenedShape(
                        elem=child,
                        tag=tag,
                        shape_id=shape_id,
                        name=name,
                        group_path=group_path,
                        z_path=z_path,
                        bbox=bbox,
                    )
                group_key = shape_id or ".".join(str(part) for part in z_path)
                group_transform = _group_child_transform(child, ns)
                yield from walk(
                    child,
                    transform.compose(group_transform),
                    group_path + (group_key,),
                    z_path,
                )
                continue

            if tag in _LEAF_DRAWABLE_TAGS:
                yield FlattenedShape(
                    elem=child,
                    tag=tag,
                    shape_id=shape_id,
                    name=name,
                    group_path=group_path,
                    z_path=z_path,
                    bbox=bbox,
                )

    yield from walk(container, _GroupTransform(), (), ())


def _flatten_slide_shapes(sp_tree: ET.Element, context: SlideConversionContext) -> List[FlattenedShape]:
    return list(_iter_flattened_shapes(sp_tree, context.ns))


def _ordered_flattened_shapes(items: List[FlattenedShape], context: SlideConversionContext) -> List[FlattenedShape]:
    known = 0
    for item in items:
        hint = context.heading_hints.get(item.shape_id, {})
        if "order_index" in hint:
            known += 1

    if known < max(1, len(items) // 2):
        return items

    def sort_key(item: FlattenedShape) -> Tuple[int, int, Tuple[int, ...]]:
        hint = context.heading_hints.get(item.shape_id, {})
        raw_order = hint.get("order_index")
        try:
            order_index = int(raw_order)
        except (TypeError, ValueError):
            return (1, 0, item.z_path)
        return (0, order_index, item.z_path)

    return sorted(items, key=sort_key)


def _table_overlay_shape_entries(items: Sequence[FlattenedShape]) -> List[Dict[str, object]]:
    return [
        {
            "elem": item.elem,
            "tag": item.tag,
            "shape_id": item.shape_id,
            "bbox": item.bbox,
        }
        for item in items
    ]


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

    if gf_kind == "diagram":
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
        return

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
    flattened_shapes = _ordered_flattened_shapes(_flatten_slide_shapes(sp_tree, context), context)
    (
        context.table_overlay_map,
        context.consumed_picture_ids,
        overlay_warnings,
        overlay_resolved,
        overlay_unresolved,
    ) = deps.collect_table_overlay_pictures(
        _table_overlay_shape_entries(flattened_shapes),
        context.slide_xml,
        context.rels_map,
        context.rels_path,
    )

    lines: List[str] = [f"[Page_{context.page_no}]", ""]
    state = SlideRenderState()

    stats = SlideStats(
        warnings=list(overlay_warnings),
        rels_path=(str(context.rels_path) if context.rels_path else None),
    )
    stats.resolved_images += overlay_resolved
    stats.unresolved_images += overlay_unresolved

    for item in flattened_shapes:
        child = item.elem
        tag = item.tag
        if tag not in {"sp", "pic", "graphicFrame", "cxnSp"}:
            continue
        stats.blocks_total += 1

        if tag == "cxnSp":
            stats.skipped_blocks += 1
            continue

        if tag == "sp":
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
