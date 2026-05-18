from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Tuple

from pptx_inheritance.resolver import (
    EffectiveShape,
    INHERITED_SHAPE_MODES,
    PPTX_INHERITANCE_MODES,
    resolve_effective_slide,
    resolve_slide_layout,
    resolve_slide_master,
)

from .structure import SlideObject


def _to_slide_object(shape: EffectiveShape) -> SlideObject:
    return SlideObject(
        shape_id=shape.shape_id,
        xml_index=shape.xml_index,
        tag=shape.tag,
        name=shape.name,
        ph_type=shape.ph_type,
        ph_idx=shape.ph_idx,
        x=shape.x,
        y=shape.y,
        coord_source=shape.coord_source,
        text=shape.text,
        normalized=shape.normalized,
        is_footer=shape.is_footer,
        is_decorative=shape.is_decorative,
        is_heading=shape.is_heading,
        is_title_placeholder=shape.is_title_placeholder,
        font_pt=shape.font_pt,
        list_kind=shape.list_kind,
        list_level=shape.list_level,
        bbox=shape.bbox,
        source_part=shape.source_part,
        inheritance_kind=shape.inheritance_kind,
    )


def extract_slide_objects_xml(
    slide_xml: Path,
    strict: bool = False,
    pptx_inheritance: str = "style",
    inherited_shapes: str = "visible",
) -> Tuple[List[SlideObject], Dict[str, object]]:
    """Return structure objects from the resolved effective slide model.

    Slide/layout/master inheritance is intentionally resolved in
    pptx_inheritance.resolver before structure analysis sees the objects. This
    keeps downstream ordering and heading code independent of raw PPTX cascade
    details.
    """
    if pptx_inheritance not in PPTX_INHERITANCE_MODES:
        raise ValueError(f"invalid pptx_inheritance: {pptx_inheritance}")
    if inherited_shapes not in INHERITED_SHAPE_MODES:
        raise ValueError(f"invalid inherited_shapes: {inherited_shapes}")

    effective_slide = resolve_effective_slide(
        slide_xml,
        strict=strict,
        pptx_inheritance=pptx_inheritance,
        inherited_shapes=inherited_shapes,
    )
    return [_to_slide_object(shape) for shape in effective_slide.shapes], effective_slide.meta()
