from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Sequence, Tuple

from pptx2markdown.structure_analyzer.xml_primitives import is_slide_number_only_shape

from .converter_models import BoundingBox, ContentBlock, ShapeBlock, SlideDocument, SlideStats
from .heading_rules import (
    HeadingPolicy,
)
from .heading_rules import (
    clean_heading_text_for_render as hr_clean_heading_text_for_render,
)
from .heading_rules import (
    infer_heading_depth_fallback as hr_infer_heading_depth_fallback,
)
from .heading_rules import (
    is_body_like_long_sentence as hr_is_body_like_long_sentence,
)
from .heading_rules import (
    looks_like_multi_numbered_items as hr_looks_like_multi_numbered_items,
)
from .heading_rules import (
    strict_heading_depth_from_placeholder as hr_strict_heading_depth_from_placeholder,
)

_DRAWABLE_TAGS = {"sp", "pic", "graphicFrame", "grpSp", "cxnSp"}
_LEAF_DRAWABLE_TAGS = {"sp", "pic", "graphicFrame", "cxnSp"}
_LARGE_INT = 10**18


def _is_hidden_slide(root: ET.Element) -> bool:
    return str(root.attrib.get("show", "1")).strip().lower() in {"0", "false", "off"}


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


@dataclass
class SlideConversionDeps:
    local_name: Callable[[str], str]
    choose_rels_in_package: Callable[[Path], Optional[Path]]
    build_rels_map: Callable[[Optional[Path]], Dict[str, str]]
    load_heading_hints: Callable[[Path], Dict[str, Dict[str, object]]]
    load_effective_properties: Callable[[Path], Dict[str, Dict[str, object]]]
    collect_table_overlay_pictures: Callable[
        [Sequence[Dict[str, object]], Path, Dict[str, str], Optional[Path]],
        Tuple[Dict[str, List[Dict[str, object]]], set[str], List[str], int, int],
    ]
    extract_shape_blocks: Callable[[ET.Element, Dict[str, str]], List[ShapeBlock]]
    render_shape_blocks: Callable[[Sequence[ShapeBlock]], str]
    split_triangle_bullets: Callable[[str], List[str]]
    normalize_text: Callable[[str], str]
    shape_id_of: Callable[[ET.Element], str]
    resolve_image_path: Callable[
        [Dict[str, str], Optional[Path], Optional[str]], Tuple[str, Optional[str]]
    ]
    format_markdown_image: Callable[..., str]
    convert_table_to_markdown: Callable[..., Tuple[Optional[str], Optional[str]]]
    graphic_frame_kind: Callable[[ET.Element], Optional[str]]
    convert_chart_to_markdown: Callable[
        [ET.Element, Optional[Path], Dict[str, str], Optional[Path]],
        Tuple[Optional[str], Optional[str]],
    ]
    convert_smartart_to_markdown: Callable[
        [
            ET.Element,
            Optional[Path],
            Dict[str, str],
            Optional[Path],
            Optional[Path],
            Optional[Path],
        ],
        Tuple[Optional[str], Optional[str]],
    ]
    convert_ole_attachment: Callable[..., Tuple[Optional[str], Optional[str]]]
    convert_media_attachment: Callable[..., Tuple[Optional[str], Optional[str]]]
    convert_model3d_attachment: Callable[..., Tuple[Optional[str], Optional[str]]]


@dataclass
class SlideRenderAssets:
    output_dir: Optional[Path] = None
    media_dir: Optional[Path] = None
    copied_media: Optional[Dict[str, Path]] = None
    attachments_dir: Optional[Path] = None
    copied_attachments: Optional[Dict[str, Path]] = None


@dataclass
class SlideConversionContext:
    slide_xml: Path
    page_no: int
    ns: Dict[str, str]
    source_pptx_path: Optional[Path] = None
    heading_mode: str = "auto"
    rels_path: Optional[Path] = None
    rels_map: Dict[str, str] = field(default_factory=dict)
    heading_hints: Dict[str, Dict[str, object]] = field(default_factory=dict)
    effective_properties: Dict[str, Dict[str, object]] = field(default_factory=dict)
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


def _point_attrs(
    node: Optional[ET.Element], x_name: str, y_name: str
) -> Optional[Tuple[int, int]]:
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


def _extract_local_bbox_emu(
    elem: ET.Element, ns: Dict[str, str]
) -> Optional[Tuple[int, int, int, int]]:
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

        choice = child.find("./{*}Choice")
        fallback = child.find("./{*}Fallback")
        if choice is not None and fallback is not None:
            graphic_data = choice.find(".//{*}graphicData")
            uri = graphic_data.attrib.get("uri", "").casefold() if graphic_data is not None else ""
            if uri.endswith("/model3d"):
                # Office stores a static preview beside the editable GLB model.
                # Keep both: the fallback is the visible slide representation,
                # while the Choice branch lets the converter preserve the source.
                yield from list(fallback)
                yield from list(choice)
                continue

        selected = choice if choice is not None else fallback
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


def _flatten_slide_shapes(
    sp_tree: ET.Element, context: SlideConversionContext
) -> List[FlattenedShape]:
    return list(_iter_flattened_shapes(sp_tree, context.ns))


def _ordered_flattened_shapes(
    items: List[FlattenedShape], context: SlideConversionContext
) -> List[FlattenedShape]:
    known = 0
    for item in items:
        hint = context.heading_hints.get(item.shape_id, {})
        if "order_index" in hint:
            known += 1

    # The structure stage has already rewritten top-level XML into XYCut order.
    # Sorting a partially matched list would move every unmatched inherited
    # object behind matched objects and destroy that correct order. Only sort
    # when every flattened leaf can be tied back to its sidecar row; this is
    # still needed for leaves inside groups, whose internal XML is not rewritten.
    if known != len(items):
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


def _with_effective_bboxes(
    items: Sequence[FlattenedShape], context: SlideConversionContext
) -> List[FlattenedShape]:
    enriched: List[FlattenedShape] = []
    for item in items:
        if item.bbox is not None:
            enriched.append(item)
            continue
        raw_bbox = context.effective_properties.get(item.shape_id, {}).get("bbox")
        if not (
            isinstance(raw_bbox, (list, tuple))
            and len(raw_bbox) == 4
            and all(isinstance(value, int) for value in raw_bbox)
        ):
            enriched.append(item)
            continue
        bbox = tuple(raw_bbox)
        if bbox[2] <= bbox[0] or bbox[3] <= bbox[1]:
            enriched.append(item)
            continue
        enriched.append(replace(item, bbox=bbox))
    return enriched


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


def _attach_block_provenance(
    blocks: List[ContentBlock],
    start_index: int,
    *,
    bbox: Optional[Tuple[int, int, int, int]],
    source_part: object,
) -> None:
    normalized_source = str(source_part or "slide")
    if normalized_source not in {"slide", "layout", "master"}:
        normalized_source = "slide"
    block_bbox = BoundingBox.from_corners(bbox) if bbox is not None else None
    for index in range(start_index, len(blocks)):
        blocks[index] = blocks[index].model_copy(
            update={"bbox": block_bbox, "source_part": normalized_source}
        )


def _append_rendered_text_block(
    blocks: List[ContentBlock],
    rendered: str,
    *,
    shape_id: str,
    kind: str,
    heading_level: Optional[int],
    deps: SlideConversionDeps,
) -> None:
    if kind != "heading":
        has_triangle_bullet = "▶" in rendered
        if has_triangle_bullet:
            rendered_lines = deps.split_triangle_bullets(rendered)
            if rendered_lines:
                rendered = "\n".join(rendered_lines)
            kind = "list"
    blocks.append(
        ContentBlock(
            kind=kind,
            content=rendered,
            shape_id=shape_id or None,
            heading_level=heading_level,
        )
    )


def _apply_effective_list_properties(
    blocks: List[ShapeBlock],
    props: Dict[str, object],
) -> List[ShapeBlock]:
    # In an ordinary text box, explicit bullet paragraphs may be mixed with
    # unmarked plain paragraphs. A shape-level summary of the first list item
    # must not turn every sibling paragraph into a bullet. Placeholder text,
    # on the other hand, can legitimately inherit bullet semantics from its
    # layout/master when individual paragraphs omit them.
    if (
        any(block.kind in {"list_ul", "list_ol"} for block in blocks)
        and not str(props.get("ph_type") or "").strip()
    ):
        return blocks
    if not props.get("has_list_semantics"):
        return blocks
    raw_list_kind = str(props.get("list_kind") or "").strip()
    list_kind = {"ul": "list_ul", "ol": "list_ol"}.get(raw_list_kind, raw_list_kind)
    if list_kind not in {"list_ul", "list_ol"}:
        return blocks
    try:
        level = int(props.get("list_level") or 0)
    except (TypeError, ValueError):
        level = 0

    converted: List[ShapeBlock] = []
    changed = False
    for block in blocks:
        if block.kind in {"list_ul", "list_ol"}:
            converted.append(block)
            continue
        if block.kind != "text" or not block.plain_text or block.list_explicit_none:
            converted.append(block)
            continue
        converted.append(
            ShapeBlock(
                kind=list_kind,
                level=level,
                segments=block.segments,
                list_explicit_none=False,
            )
        )
        changed = True
    return converted if changed else blocks


def _font_separated_heading_sections(blocks: Sequence[ShapeBlock]) -> List[str]:
    if len(blocks) != 1 or any(
        segment.kind not in {"text", "break"} for segment in blocks[0].segments
    ):
        return []
    lines: List[Tuple[str, List[float]]] = []
    text_parts: List[str] = []
    font_sizes: List[float] = []

    def flush() -> None:
        text = "".join(text_parts).strip()
        if text:
            lines.append((text, list(font_sizes)))
        text_parts.clear()
        font_sizes.clear()

    for segment in blocks[0].segments:
        if segment.kind == "break":
            flush()
            continue
        text_parts.append(segment.text)
        if segment.font_pt is not None and segment.font_pt > 0:
            font_sizes.append(segment.font_pt)
    flush()
    if len(lines) < 2 or not lines[0][1]:
        return []
    remaining_sizes = [size for _, sizes in lines[1:] for size in sizes]
    if not remaining_sizes or max(lines[0][1]) < max(remaining_sizes) * 1.5:
        return []
    return [lines[0][0], "\n".join(text for text, _ in lines[1:])]


def _handle_text_shape_block(
    child: ET.Element,
    *,
    blocks: List[ContentBlock],
    heading_policy: HeadingPolicy,
    strict_headings: bool,
    state: SlideRenderState,
    stats: SlideStats,
    context: SlideConversionContext,
    deps: SlideConversionDeps,
) -> None:
    ph = child.find(".//p:ph", context.ns)
    ph_type = ph.attrib.get("type") if ph is not None else None
    is_unpositioned_literal_date = (
        ph_type == "dt"
        and child.find(".//a:fld", context.ns) is None
        and child.find("./p:spPr/a:xfrm", context.ns) is None
        and any((node.text or "").strip() for node in child.findall(".//a:t", context.ns))
    )
    if (
        ph_type in {"sldNum", "ftr"}
        or (ph_type == "dt" and not is_unpositioned_literal_date)
        or is_slide_number_only_shape(child)
    ):
        stats.skipped_blocks += 1
        return

    sid = deps.shape_id_of(child)
    props = context.effective_properties.get(sid, {})
    shape_blocks = _apply_effective_list_properties(
        deps.extract_shape_blocks(child, context.rels_map), props
    )
    has_list_semantics = any(block.kind in {"list_ul", "list_ol"} for block in shape_blocks)
    has_math_shape = any(block.has_math for block in shape_blocks)
    text = deps.render_shape_blocks(shape_blocks)
    plain_text = " ".join(block.plain_text for block in shape_blocks if block.plain_text).strip()
    if not text or not plain_text:
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
    text_sections = [
        re.sub(r"\s+", " ", section).strip()
        for section in re.split(r"\n{2,}", text or "")
        if section.strip()
    ]
    font_sections = _font_separated_heading_sections(shape_blocks)
    if len(text_sections) == 1 and font_sections:
        text_sections = font_sections
    heading_ph_type = ph_type if ph_type is not None else hint.get("ph_type")
    can_split_heading_body = (
        len(text_sections) > 1
        and heading_ph_type in {"title", "ctrTitle", "subTitle"}
        and not has_list_semantics
        and not has_math_shape
    )

    strong_heading_signal = (
        not strict_headings
        and is_candidate
        and isinstance(depth, int)
        and 1 <= depth <= 6
        and score >= heading_policy.threshold
    )

    if strict_headings:
        strict_ph_type = ph_type if ph_type is not None else hint.get("ph_type")
        strict_depth = hr_strict_heading_depth_from_placeholder(strict_ph_type)
        if strict_depth is not None:
            is_candidate = True
            depth = strict_depth
            score = 1.0
        else:
            is_candidate = False
            depth = None
            score = 0.0
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
                if re.match(
                    r"^\d+\s*\.\s*\d+(?:\s*\.\s*\d+)*\s*\.?\s+",
                    rendered_text,
                ):
                    fb_depth = 3
                elif re.match(r"^\d+\s*\.\s+", rendered_text):
                    fb_depth = 2
                else:
                    fb_depth = None
            else:
                fb_depth = None
            if fb_depth is None:
                fb_depth = hr_infer_heading_depth_fallback(
                    rendered_text, state.text_block_index, font_pt=font_pt
                )
            if fb_depth is None:
                fb_depth = hr_infer_heading_depth_fallback(
                    plain_text, state.text_block_index, font_pt=font_pt
                )
            if fb_depth is not None:
                depth = fb_depth
                score = max(score, 0.8)
                is_candidate = True
                strong_heading_signal = True

    rendered = text
    trailing_text: Optional[str] = None
    block_kind = "math" if has_math_shape else "text"
    heading_level: Optional[int] = None
    if not strict_headings and heading_ph_type not in {"title", "ctrTitle", "subTitle"}:
        if hr_looks_like_multi_numbered_items(plain_text):
            is_candidate = False
        elif not strong_heading_signal and hr_is_body_like_long_sentence(plain_text):
            is_candidate = False

    heading_threshold = heading_policy.threshold
    if is_candidate and isinstance(depth, int) and 1 <= depth <= 6 and score >= heading_threshold:
        if strict_headings:
            heading_text = text_sections[0] if can_split_heading_body else plain_text
        else:
            heading_source = text_sections[0] if can_split_heading_body else plain_text
            if re.match(
                r"^\d+(?:\s*\.\s*\d+)*\s*\.?\s+",
                rendered_text,
            ):
                heading_source = rendered_text
            heading_text = hr_clean_heading_text_for_render(heading_source)
        key = deps.normalize_text(heading_text)
        if key not in state.used_headings:
            rendered = heading_text
            block_kind = "heading"
            heading_level = depth
            state.used_headings.add(key)
            if can_split_heading_body:
                trailing_text = "\n\n".join(text_sections[1:])
        else:
            stats.skipped_blocks += 1
            return

    if has_list_semantics and block_kind == "text":
        block_kind = "list"
    _append_rendered_text_block(
        blocks,
        rendered,
        shape_id=sid,
        kind=block_kind,
        heading_level=heading_level,
        deps=deps,
    )
    if block_kind == "heading" and trailing_text:
        _append_rendered_text_block(
            blocks,
            trailing_text,
            shape_id=sid,
            kind="text",
            heading_level=None,
            deps=deps,
        )
    stats.text_blocks += 1
    state.text_block_index += 1


def _shape_has_image_fill(
    child: ET.Element,
    ns: Dict[str, str],
    rels_map: Optional[Dict[str, str]] = None,
) -> bool:
    blip = child.find("./p:spPr/a:blipFill/a:blip", ns)
    if blip is None:
        return False
    embed = blip.attrib.get(f"{{{ns['r']}}}embed")
    if not embed:
        return False
    if rels_map is None:
        return True
    target = rels_map.get(embed, "")
    return Path(target).suffix.casefold() in {
        ".bmp",
        ".emf",
        ".gif",
        ".jpeg",
        ".jpg",
        ".png",
        ".svg",
        ".tif",
        ".tiff",
        ".webp",
        ".wmf",
    }


def _shape_has_text_or_math(child: ET.Element, ns: Dict[str, str]) -> bool:
    if any((node.text or "").strip() for node in child.findall(".//a:t", ns)):
        return True
    return any(
        isinstance(node.tag, str) and _local_name(node.tag) in {"oMath", "oMathPara"}
        for node in child.iter()
    )


def _handle_picture_block(
    child: ET.Element,
    *,
    blocks: List[ContentBlock],
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
    emitted = False
    if embed:
        img_path, warn = deps.resolve_image_path(context.rels_map, context.rels_path, embed)
        if warn:
            stats.unresolved_images += 1
            stats.warnings.append(warn)
        else:
            stats.resolved_images += 1

        rendered_image = deps.format_markdown_image(
            img_path,
            output_dir=assets.output_dir,
            media_dir=assets.media_dir,
            copied_media=assets.copied_media,
        )

        blocks.append(ContentBlock(kind="image", content=rendered_image, shape_id=sid or None))
        stats.image_blocks += 1
        emitted = True
    attachment_md, attachment_warning = deps.convert_media_attachment(
        child,
        rels_path=context.rels_path,
        rels_map=context.rels_map,
        output_dir=assets.output_dir,
        attachments_dir=assets.attachments_dir,
        copied_attachments=assets.copied_attachments,
    )
    if attachment_md:
        blocks.append(ContentBlock(kind="attachment", content=attachment_md, shape_id=sid or None))
        stats.attachment_blocks += 1
        emitted = True
    if attachment_warning:
        stats.warnings.append(attachment_warning)
    if not emitted:
        # Empty blips are commonly emitted as inert master decorations or
        # sub-pixel sentinel pictures. They have no visible/static payload.
        stats.skipped_blocks += 1


def _append_unsupported_graphic_frame(
    blocks: List[ContentBlock],
    stats: SlideStats,
    shape_id: str,
    label: str = "graphicFrame(non-table)",
) -> None:
    blocks.append(
        ContentBlock(
            kind="unsupported",
            content=f"[unsupported: {label}]",
            shape_id=shape_id or None,
        )
    )
    stats.unsupported_blocks += 1


def _handle_graphic_frame_block(
    child: ET.Element,
    *,
    blocks: List[ContentBlock],
    state: SlideRenderState,
    stats: SlideStats,
    context: SlideConversionContext,
    assets: SlideRenderAssets,
    deps: SlideConversionDeps,
) -> None:
    shape_id = deps.shape_id_of(child)
    gf_kind = deps.graphic_frame_kind(child)
    if gf_kind == "chart":
        chart_md, chart_err = deps.convert_chart_to_markdown(
            child,
            context.rels_path,
            context.rels_map,
            context.source_pptx_path,
        )
        if chart_md is not None:
            blocks.append(
                ContentBlock(kind="chart", content=chart_md.strip(), shape_id=shape_id or None)
            )
            stats.chart_blocks += 1
            return
        _append_unsupported_graphic_frame(blocks, stats, shape_id)
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
            blocks.append(
                ContentBlock(kind="smartart", content=smartart_md, shape_id=shape_id or None)
            )
            stats.smartart_blocks += 1
        elif smartart_err == "smartart contains no data nodes":
            stats.skipped_blocks += 1
        else:
            _append_unsupported_graphic_frame(blocks, stats, shape_id)
            stats.warnings.append(smartart_err or "smartart conversion failed")
        return

    if gf_kind == "ole":
        attachment_md, attachment_error = deps.convert_ole_attachment(
            child,
            rels_path=context.rels_path,
            rels_map=context.rels_map,
            output_dir=assets.output_dir,
            attachments_dir=assets.attachments_dir,
            copied_attachments=assets.copied_attachments,
        )
        if attachment_md:
            blocks.append(
                ContentBlock(
                    kind="attachment",
                    content=attachment_md,
                    shape_id=shape_id or None,
                )
            )
            stats.attachment_blocks += 1
            if attachment_error:
                stats.warnings.append(attachment_error)
            return
        # Linked OLE objects often have no portable payload but do carry an
        # embedded preview picture in their AlternateContent fallback. Keep
        # that visible representation instead of emitting an unsupported
        # marker for an unavailable machine-local link.
        blip = child.find(".//a:blip", context.ns)
        preview_embed = (
            blip.attrib.get(f"{{{context.ns['r']}}}embed") if blip is not None else None
        )
        if preview_embed:
            preview_path, preview_warning = deps.resolve_image_path(
                context.rels_map,
                context.rels_path,
                preview_embed,
            )
            if preview_warning:
                stats.unresolved_images += 1
                stats.warnings.append(preview_warning)
            else:
                rendered_preview = deps.format_markdown_image(
                    preview_path,
                    output_dir=assets.output_dir,
                    media_dir=assets.media_dir,
                    copied_media=assets.copied_media,
                )
                blocks.append(
                    ContentBlock(
                        kind="image",
                        content=rendered_preview,
                        shape_id=shape_id or None,
                    )
                )
                stats.image_blocks += 1
                stats.resolved_images += 1
                if attachment_error:
                    stats.warnings.append(attachment_error)
                return
        _append_unsupported_graphic_frame(blocks, stats, shape_id, "ole-object")
        stats.warnings.append(attachment_error or "OLE attachment extraction failed")
        return

    if gf_kind == "model3d":
        attachment_md, attachment_error = deps.convert_model3d_attachment(
            child,
            rels_path=context.rels_path,
            rels_map=context.rels_map,
            output_dir=assets.output_dir,
            attachments_dir=assets.attachments_dir,
            copied_attachments=assets.copied_attachments,
        )
        if attachment_md:
            blocks.append(
                ContentBlock(
                    kind="attachment",
                    content=attachment_md,
                    shape_id=shape_id or None,
                )
            )
            stats.attachment_blocks += 1
            if attachment_error:
                stats.warnings.append(attachment_error)
            return
        _append_unsupported_graphic_frame(blocks, stats, shape_id, "model3d")
        stats.warnings.append(attachment_error or "3D model extraction failed")
        return

    table_md, err = deps.convert_table_to_markdown(
        child,
        overlays=context.table_overlay_map.get(shape_id, []),
        output_dir=assets.output_dir,
        media_dir=assets.media_dir,
        copied_media=assets.copied_media,
        rels_path=context.rels_path,
        rels_map=context.rels_map,
    )
    if table_md is not None:
        blocks.append(
            ContentBlock(kind="table", content=table_md.strip(), shape_id=shape_id or None)
        )
        stats.table_blocks += 1
        return

    _append_unsupported_graphic_frame(blocks, stats, shape_id)
    if err:
        stats.warnings.append(err)


def convert_one_slide(
    *,
    context: SlideConversionContext,
    assets: SlideRenderAssets,
    strict_headings: bool = False,
    deps: SlideConversionDeps,
) -> Tuple[SlideDocument, SlideStats]:
    heading_policy = HeadingPolicy(strict=strict_headings)

    root = ET.parse(context.slide_xml).getroot()
    sp_tree = root.find("p:cSld/p:spTree", context.ns)
    if sp_tree is None:
        raise ValueError("missing p:cSld/p:spTree")

    context.rels_path = deps.choose_rels_in_package(context.slide_xml)
    context.rels_map = deps.build_rels_map(context.rels_path)
    context.heading_hints = deps.load_heading_hints(context.slide_xml)
    context.effective_properties = deps.load_effective_properties(context.slide_xml)
    flattened_shapes = _ordered_flattened_shapes(
        _with_effective_bboxes(_flatten_slide_shapes(sp_tree, context), context),
        context,
    )
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

    blocks: List[ContentBlock] = []
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

        sid = deps.shape_id_of(child)
        props = context.effective_properties.get(sid, {})
        if props.get("is_decorative"):
            stats.skipped_blocks += 1
            continue

        block_start = len(blocks)
        if tag == "sp":
            has_image_fill = _shape_has_image_fill(child, context.ns, context.rels_map)
            if has_image_fill:
                _handle_picture_block(
                    child,
                    blocks=blocks,
                    state=state,
                    stats=stats,
                    context=context,
                    assets=assets,
                    deps=deps,
                )
            if _shape_has_text_or_math(child, context.ns):
                _handle_text_shape_block(
                    child,
                    blocks=blocks,
                    heading_policy=heading_policy,
                    strict_headings=strict_headings,
                    state=state,
                    stats=stats,
                    context=context,
                    deps=deps,
                )
            elif not has_image_fill:
                stats.skipped_blocks += 1
        elif tag == "pic":
            _handle_picture_block(
                child,
                blocks=blocks,
                state=state,
                stats=stats,
                context=context,
                assets=assets,
                deps=deps,
            )
        elif tag == "graphicFrame":
            _handle_graphic_frame_block(
                child,
                blocks=blocks,
                state=state,
                stats=stats,
                context=context,
                assets=assets,
                deps=deps,
            )
        _attach_block_provenance(
            blocks,
            block_start,
            bbox=item.bbox,
            source_part=props.get("source_part", "slide"),
        )

    heading_blocks = [block for block in blocks if block.kind == "heading"]
    if len(heading_blocks) == 1 and heading_blocks[0].heading_level != 1:
        only_heading = heading_blocks[0]
        blocks[blocks.index(only_heading)] = only_heading.model_copy(update={"heading_level": 1})

    return SlideDocument(
        page=context.page_no,
        hidden=_is_hidden_slide(root),
        blocks=blocks,
    ), stats
