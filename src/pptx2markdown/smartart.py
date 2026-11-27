"""SmartArt data models converted to Markdown through smartart2md."""

from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass

from smartart2md import convert_smartart

from .ooxml import OOXMLPackage

DGM_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
# Points that only structure the diagram; every other point is content.
_STRUCTURAL_POINTS = {"doc", "parTrans", "sibTrans", "pres"}
_IMAGE_PLACEHOLDER = re.compile(r"@@IMG:(\d+)@@")


class _PackageContext:
    """The part of smartart2md's ``ZipContext`` it uses, read through the package.

    smartart2md resolves one relationship per picture, in placeholder order. The
    resolved part names are recorded so pictures become managed assets; the bytes
    smartart2md asks for are never used, so none are read.
    """

    def __init__(self, package: OOXMLPackage, path: str, resolved: list[str | None]) -> None:
        self._package = package
        self.path = path
        self._resolved = resolved

    def resolve(self, rid: str) -> _PackageContext | None:
        target = self._package.related_part(self.path, rid)
        self._resolved.append(target)
        if target is None:
            return None
        return _PackageContext(self._package, target, self._resolved)

    def read_bytes(self) -> bytes:
        return b""


@dataclass(frozen=True)
class SmartArtConversion:
    """Markdown with ``@@IMG:N@@`` placeholders and the picture part behind each N."""

    markdown: str
    picture_parts: tuple[str | None, ...]
    content_points: int

    def render(self, picture: Callable[[str], str]) -> str:
        """Replace each placeholder with ``picture(part)``, or drop it when unresolved."""

        def replace(match: re.Match[str]) -> str:
            index = int(match.group(1))
            part = self.picture_parts[index] if index < len(self.picture_parts) else None
            return picture(part) if part else ""

        return _IMAGE_PLACEHOLDER.sub(replace, self.markdown).strip()


def smartart_markdown(package: OOXMLPackage, data_part: str) -> SmartArtConversion:
    """Convert a SmartArt data model part: hierarchy, process and cycle layouts,
    matrices and pictures, as smartart2md renders them."""

    root = package.xml(data_part)
    resolved: list[str | None] = []
    markdown, _ = convert_smartart(root, _PackageContext(package, data_part, resolved))
    content_points = sum(
        1
        for point in root.iter(f"{{{DGM_NS}}}pt")
        if point.attrib.get("type", "node") not in _STRUCTURAL_POINTS
    )
    return SmartArtConversion(markdown.strip(), tuple(resolved), content_points)
