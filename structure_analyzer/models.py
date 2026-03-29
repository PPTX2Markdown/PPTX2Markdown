from __future__ import annotations

from dataclasses import dataclass
from typing import Optional, Tuple


@dataclass
class SlideObject:
    shape_id: str
    xml_index: int
    tag: str
    name: str
    ph_type: Optional[str]
    ph_idx: Optional[str]
    x: int
    y: int
    coord_source: str
    text: str
    normalized: str
    is_footer: bool
    is_decorative: bool
    is_heading: bool
    is_title_placeholder: bool
    font_pt: Optional[float]
    bbox: Optional[Tuple[int, int, int, int]] = None


@dataclass
class OrderContext:
    slide_left: int
    slide_top: int
    slide_right: int
    slide_bottom: int
    slide_width: int
    slide_height: int
    title_band_bottom: int
