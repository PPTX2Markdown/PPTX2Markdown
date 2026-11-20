"""Slide shapes: namespaces, geometry through groups, and shape classification."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from dataclasses import dataclass, replace

from .geometry import Transform, group_transform, integer, xfrm_bbox
from .headings import ShapeSignals
from .inheritance import (
    FOOTER_TYPES,
    PartKind,
    ShapeStyle,
    SlideTemplate,
    is_slide_number_only,
    shape_text,
)
from .model import BoundingBox
from .ooxml import OFFICE_REL_NS
from .tables import cell_bounds

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
CX_NS = "http://schemas.microsoft.com/office/drawing/2014/chartex"
DGM_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
A14_NS = "http://schemas.microsoft.com/office/drawing/2010/main"
# PowerPoint wraps equations in a14:m, whose local name collides with the OMML matrix m:m.
MATH_WRAPPER_TAG = f"{{{A14_NS}}}m"
MATH_TAGS = {f"{{{M_NS}}}oMath", f"{{{M_NS}}}oMathPara"}

NS = {
    "p": P_NS,
    "a": A_NS,
    "c": C_NS,
    "cx": CX_NS,
    "dgm": DGM_NS,
    "m": M_NS,
    "r": OFFICE_REL_NS,
}


@dataclass(frozen=True)
class Shape:
    element: ET.Element
    tag: str
    shape_id: str | None
    bbox: tuple[int, int, int, int] | None
    source_order: int
    source_part: PartKind = "slide"
    # Package part holding the element when it is inherited from a layout or master.
    template_part: str | None = None
    style: ShapeStyle = ShapeStyle()


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def element_shape_id(element: ET.Element) -> str | None:
    non_visual = element.find(".//p:cNvPr", NS)
    if non_visual is None:
        return None
    return non_visual.attrib.get("id") or None


def flatten_shapes(sp_tree: ET.Element) -> list[Shape]:
    shapes: list[Shape] = []
    source_order = 0

    def walk(parent: ET.Element, transform: Transform) -> None:
        nonlocal source_order
        for element in list(parent):
            tag = local_name(element.tag)
            if tag == "AlternateContent":
                choice = next(
                    (
                        child
                        for child in element
                        if local_name(child.tag) == "Choice" and len(child)
                    ),
                    None,
                )
                fallback = next(
                    (
                        child
                        for child in element
                        if local_name(child.tag) == "Fallback" and len(child)
                    ),
                    None,
                )
                selected = choice if choice is not None else fallback
                if selected is not None:
                    walk(selected, transform)
                continue
            if tag == "grpSp":
                walk(element, transform.compose(group_transform(element)))
                continue
            if tag not in {"sp", "pic", "graphicFrame", "cxnSp"}:
                continue
            bbox = xfrm_bbox(element)
            if bbox is not None:
                bbox = transform.apply(bbox)
            shapes.append(
                Shape(
                    element=element,
                    tag=tag,
                    shape_id=element_shape_id(element),
                    bbox=bbox,
                    source_order=source_order,
                )
            )
            source_order += 1

    walk(sp_tree, Transform())
    return shapes


def bbox_model(bbox: tuple[int, int, int, int] | None) -> BoundingBox | None:
    if bbox is None:
        return None
    x1, y1, x2, y2 = bbox
    if x2 <= x1 or y2 <= y1:
        return None
    return BoundingBox(x=x1, y=y1, width=x2 - x1, height=y2 - y1)


def attr(element: ET.Element | None, name: str) -> str | None:
    return element.attrib.get(name) if element is not None else None


_EXT_PATHS = ("./p:spPr/a:xfrm/a:ext", "./p:grpSpPr/a:xfrm/a:ext", "./p:xfrm/a:ext", ".//a:ext")


def is_unpositioned_literal_date(shape: Shape) -> bool:
    """A date placeholder holding typed text but no geometry of its own."""

    element = shape.element
    return (
        shape.style.placeholder_type == "dt"
        and element.find(".//a:fld", NS) is None
        and element.find("./p:spPr/a:xfrm", NS) is None
        and any((node.text or "").strip() for node in element.findall(".//a:t", NS))
    )


def is_footer(shape: Shape) -> bool:
    if shape.style.placeholder_type in FOOTER_TYPES and not is_unpositioned_literal_date(shape):
        return True
    return is_slide_number_only(shape.element)


def is_decorative(shape: Shape, text: str) -> bool:
    element = shape.element
    if shape.tag == "cxnSp":
        return True
    blip = element.find(".//a:blip", NS)
    has_image = blip is not None and bool(blip.attrib.get(f"{{{OFFICE_REL_NS}}}embed"))
    ext = next(
        (found for path in _EXT_PATHS if (found := element.find(path, NS)) is not None), None
    )
    zero_size = ext is not None and (
        integer(attr(ext, "cx"), default=1) <= 0 or integer(attr(ext, "cy"), default=1) <= 0
    )
    if shape.tag == "pic" and zero_size:
        return True
    if shape.tag == "sp" and has_image:
        return zero_size
    if shape.tag == "pic" and shape.style.placeholder_type is not None and not has_image:
        return True
    has_math = any(element.find(f".//m:{name}", NS) is not None for name in ("oMath", "oMathPara"))
    return not text.strip() and not has_math and shape.tag not in {"pic", "graphicFrame"}


def heading_signals(shape: Shape) -> ShapeSignals:
    text = shape_text(shape.element)
    return ShapeSignals(
        text=text,
        placeholder_type=shape.style.placeholder_type,
        font_pt=shape.style.font_pt,
        bbox=shape.bbox,
        position_from_layout=shape.style.position_source == "layout",
        is_footer=is_footer(shape),
        is_decorative=is_decorative(shape, text),
    )


def with_template_style(shape: Shape, template: SlideTemplate) -> Shape:
    style = template.style(shape.element, shape.bbox)
    styled = replace(shape, style=style, bbox=style.bbox)
    if styled.bbox is None and is_unpositioned_literal_date(styled):
        # Keep visible typed dates at the top of the reading order instead of appending them.
        styled = replace(styled, bbox=(0, 0, 1, 1))
    return styled


def table_overlays(shapes: list[Shape]) -> dict[int, list[Shape]]:
    """Pictures lying on a table, keyed by the table's element.

    A picture belongs to a table when its centre is inside the table and at least 80% of
    it overlaps the table; it is then written into the cell under its centre instead of
    after the table. Pictures that carry audio or video keep their own blocks.
    """

    tables = [
        shape
        for shape in shapes
        if shape.tag == "graphicFrame"
        and shape.source_part == "slide"
        and shape.bbox is not None
        and (table := shape.element.find(".//a:tbl", NS)) is not None
        and cell_bounds(table, shape.bbox) is not None
    ]
    overlays: dict[int, list[Shape]] = {}
    for picture in shapes:
        if (
            picture.tag != "pic"
            or picture.source_part != "slide"
            or picture.bbox is None
            or picture.element.find(".//a:audioFile", NS) is not None
            or picture.element.find(".//a:videoFile", NS) is not None
        ):
            continue
        left, top, right, bottom = picture.bbox
        area = (right - left) * (bottom - top)
        if area <= 0:
            continue
        center = ((left + right) // 2, (top + bottom) // 2)
        for table in tables:
            t_left, t_top, t_right, t_bottom = table.bbox  # type: ignore[misc]
            overlap = max(0, min(right, t_right) - max(left, t_left)) * max(
                0, min(bottom, t_bottom) - max(top, t_top)
            )
            inside = t_left <= center[0] <= t_right and t_top <= center[1] <= t_bottom
            if inside and overlap / area >= 0.8:
                overlays.setdefault(id(table.element), []).append(picture)
                break
    return overlays
