"""Shape bounding boxes in slide EMU, through any enclosing group transforms."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from dataclasses import dataclass

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


@dataclass(frozen=True)
class Transform:
    sx: float = 1.0
    sy: float = 1.0
    tx: float = 0.0
    ty: float = 0.0

    def apply(self, bbox: tuple[int, int, int, int]) -> tuple[int, int, int, int]:
        x1, y1, x2, y2 = bbox
        return (
            int(round(x1 * self.sx + self.tx)),
            int(round(y1 * self.sy + self.ty)),
            int(round(x2 * self.sx + self.tx)),
            int(round(y2 * self.sy + self.ty)),
        )

    def compose(self, inner: "Transform") -> "Transform":
        return Transform(
            sx=self.sx * inner.sx,
            sy=self.sy * inner.sy,
            tx=self.sx * inner.tx + self.tx,
            ty=self.sy * inner.ty + self.ty,
        )


def integer(value: str | None, default: int = 0) -> int:
    try:
        return int(value or default)
    except (TypeError, ValueError):
        return default


def xfrm_bbox(element: ET.Element) -> tuple[int, int, int, int] | None:
    xfrm = element.find("./p:spPr/a:xfrm", NS)
    if xfrm is None:
        xfrm = element.find("./p:xfrm", NS)
    if xfrm is None:
        xfrm = element.find("./p:grpSpPr/a:xfrm", NS)
    if xfrm is None:
        return None
    offset = xfrm.find("./a:off", NS)
    extent = xfrm.find("./a:ext", NS)
    if offset is None or extent is None:
        return None
    x = integer(offset.attrib.get("x"))
    y = integer(offset.attrib.get("y"))
    width = integer(extent.attrib.get("cx"))
    height = integer(extent.attrib.get("cy"))
    if width <= 0 or height <= 0:
        return None
    return x, y, x + width, y + height


def group_transform(element: ET.Element) -> Transform:
    xfrm = element.find("./p:grpSpPr/a:xfrm", NS)
    if xfrm is None:
        return Transform()
    offset = xfrm.find("./a:off", NS)
    extent = xfrm.find("./a:ext", NS)
    child_offset = xfrm.find("./a:chOff", NS)
    child_extent = xfrm.find("./a:chExt", NS)
    if None in {offset, extent, child_offset, child_extent}:
        return Transform()
    assert offset is not None
    assert extent is not None
    assert child_offset is not None
    assert child_extent is not None
    child_width = integer(child_extent.attrib.get("cx"))
    child_height = integer(child_extent.attrib.get("cy"))
    if child_width <= 0 or child_height <= 0:
        return Transform()
    sx = integer(extent.attrib.get("cx")) / child_width
    sy = integer(extent.attrib.get("cy")) / child_height
    child_x = integer(child_offset.attrib.get("x"))
    child_y = integer(child_offset.attrib.get("y"))
    return Transform(
        sx=sx,
        sy=sy,
        tx=integer(offset.attrib.get("x")) - child_x * sx,
        ty=integer(offset.attrib.get("y")) - child_y * sy,
    )
