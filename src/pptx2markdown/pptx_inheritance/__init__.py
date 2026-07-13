"""PPTX slide/layout/master inheritance resolution."""

from .resolver import (
    INHERITED_SHAPE_MODES,
    PPTX_INHERITANCE_MODES,
    EffectiveShape,
    EffectiveSlide,
    normalize_inherited_shapes_mode,
    normalize_pptx_inheritance_mode,
    resolve_effective_slide,
    resolve_slide_layout,
    resolve_slide_master,
)

__all__ = [
    "EffectiveShape",
    "EffectiveSlide",
    "INHERITED_SHAPE_MODES",
    "PPTX_INHERITANCE_MODES",
    "normalize_inherited_shapes_mode",
    "normalize_pptx_inheritance_mode",
    "resolve_effective_slide",
    "resolve_slide_layout",
    "resolve_slide_master",
]
