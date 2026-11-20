"""Placeholder, text-style and shape inheritance from slide layouts and masters."""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import Literal

from .geometry import Transform, group_transform, xfrm_bbox
from .ooxml import OOXMLPackage, resolve_relationship_target

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"p": P_NS, "a": A_NS}

TITLE_TYPES = frozenset({"title", "ctrTitle", "subTitle"})
FOOTER_TYPES = frozenset({"dt", "ftr", "sldNum"})
_DEFAULT_TYPE = "obj"
_DEFAULT_IDX = "0"
_SPECIAL_IDX = str(0xFFFFFFFF)
_DRAWABLE_TAGS = frozenset({"sp", "pic", "graphicFrame", "grpSp", "cxnSp"})
_PROMPT_TOKEN = re.compile(r"[<‹«]\s*(?:#|slide\s+number|date(?:/time)?|footer)\s*[>›»]")

PartKind = Literal["slide", "layout", "master"]
ListKind = Literal["ul", "ol"]
BBox = tuple[int, int, int, int]
# Share of the slide a layout or master picture must cover to count as background art.
BACKDROP_COVERAGE = 0.9


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _int(value: str | None) -> int | None:
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def placeholder_element(shape: ET.Element | None) -> ET.Element | None:
    if shape is None:
        return None
    return shape.find("./*/p:nvPr/p:ph", NS)


def _normalize_idx(value: str | None) -> str:
    raw = (value or "").strip()
    if not raw:
        return _DEFAULT_IDX
    try:
        return str(int(raw, 0))
    except ValueError:
        return raw


def _is_slide_number_field(field: ET.Element) -> bool:
    return (field.attrib.get("type") or "").strip().casefold() == "slidenum"


def shape_text(shape: ET.Element) -> str:
    """Visible text of a shape, excluding slide-number field runs."""

    excluded = {
        id(node)
        for field in shape.findall(".//a:fld", NS)
        if _is_slide_number_field(field)
        for node in field.findall(".//a:t", NS)
    }
    return " ".join(
        node.text.strip()
        for node in shape.findall(".//a:t", NS)
        if id(node) not in excluded and node.text and node.text.strip()
    )


def is_slide_number_only(shape: ET.Element) -> bool:
    has_field = any(_is_slide_number_field(field) for field in shape.findall(".//a:fld", NS))
    return has_field and not shape_text(shape)


def _is_prompt_text(text: str) -> bool:
    folded = re.sub(r"\s+", " ", text).strip().casefold()
    if folded.startswith(("click to edit master", "click to add ")):
        return True
    return bool(_PROMPT_TOKEN.fullmatch(folded))


def _first(element: ET.Element, paths: tuple[str, ...]) -> ET.Element | None:
    for path in paths:
        found = element.find(path, NS)
        if found is not None:
            return found
    return None


def _xfrm_components(shape: ET.Element | None) -> dict[str, int | None]:
    if shape is None:
        return {"x": None, "y": None, "cx": None, "cy": None}
    off = _first(
        shape,
        ("./p:spPr/a:xfrm/a:off", "./p:grpSpPr/a:xfrm/a:off", "./p:xfrm/a:off", ".//a:off"),
    )
    ext = _first(
        shape,
        ("./p:spPr/a:xfrm/a:ext", "./p:grpSpPr/a:xfrm/a:ext", "./p:xfrm/a:ext", ".//a:ext"),
    )
    cx = _int(ext.attrib.get("cx")) if ext is not None else None
    cy = _int(ext.attrib.get("cy")) if ext is not None else None
    autofit = shape.find(".//a:spAutoFit", NS) is not None
    if autofit and cx is not None and cy is not None and cx >= 0 and cy >= 0 and (cx or cy):
        cx, cy = max(1, cx), max(1, cy)
    return {
        "x": _int(off.attrib.get("x")) if off is not None else None,
        "y": _int(off.attrib.get("y")) if off is not None else None,
        "cx": cx if cx is not None and cx > 0 else None,
        "cy": cy if cy is not None and cy > 0 else None,
    }


def _max_font_pt(shape: ET.Element | None) -> float | None:
    if shape is None:
        return None
    sizes = [
        size / 100.0
        for node in [*shape.findall(".//a:rPr", NS), *shape.findall(".//a:endParaRPr", NS)]
        if (size := _int(node.attrib.get("sz"))) is not None and size > 0
    ]
    return max(sizes) if sizes else None


def _paragraph_level(shape: ET.Element) -> int:
    for properties in shape.findall(".//a:pPr", NS):
        level = _int(properties.attrib.get("lvl"))
        if level is not None:
            return max(0, min(level, 8))
    return 0


def _list_kind(properties: ET.Element | None) -> ListKind | None:
    if properties is None or properties.find("a:buNone", NS) is not None:
        return None
    if properties.find("a:buAutoNum", NS) is not None:
        return "ol"
    if properties.find("a:buChar", NS) is not None:
        return "ul"
    return None


def _shape_list_style(shape: ET.Element | None) -> tuple[ListKind, int] | None:
    if shape is None:
        return None
    for properties in shape.findall(".//a:pPr", NS):
        kind = _list_kind(properties)
        if kind is not None:
            level = _int(properties.attrib.get("lvl"))
            return kind, max(0, min(level, 8)) if level is not None else 0
    return None


@dataclass(frozen=True)
class ShapeStyle:
    """Effective properties of one shape after layout and master inheritance."""

    placeholder_type: str | None = None
    bbox: BBox | None = None
    position_source: PartKind = "slide"
    font_pt: float | None = None
    list_kind: ListKind | None = None
    list_level: int = 0


@dataclass(frozen=True)
class InheritedShape:
    """A layout or master shape that is drawn on the slide."""

    source_part: PartKind
    part_name: str
    element: ET.Element
    # Slide EMU, through the groups that hold the shape in its layout or master.
    bbox: BBox | None = None


def _drawn_shapes(
    parent: ET.Element, transform: Transform = Transform()
) -> list[tuple[ET.Element, BBox | None]]:
    """Shapes and pictures of a shape tree, inside groups too, with their bounding boxes."""

    drawn: list[tuple[ET.Element, BBox | None]] = []
    for child in parent:
        name = _local_name(child.tag)
        if name == "grpSp":
            drawn.extend(_drawn_shapes(child, transform.compose(group_transform(child))))
        elif name in {"sp", "pic"}:
            bbox = xfrm_bbox(child)
            drawn.append((child, transform.apply(bbox) if bbox is not None else None))
    return drawn


@dataclass(frozen=True)
class _TemplatePart:
    name: str
    root: ET.Element
    drawables: tuple[ET.Element, ...]
    drawn: tuple[tuple[ET.Element, BBox | None], ...]
    placeholders: dict[str, ET.Element]
    ambiguous: frozenset[str]

    @classmethod
    def from_root(cls, name: str, root: ET.Element) -> "_TemplatePart":
        tree = root.find("./p:cSld/p:spTree", NS)
        drawables = tuple(
            child
            for child in (tree if tree is not None else ())
            if _local_name(child.tag) in _DRAWABLE_TAGS
        )
        placeholders: dict[str, ET.Element] = {}
        ambiguous: set[str] = set()
        for child in drawables:
            placeholder = placeholder_element(child)
            if placeholder is None:
                continue
            idx = _normalize_idx(placeholder.attrib.get("idx"))
            if idx in placeholders:
                ambiguous.add(idx)
            else:
                placeholders[idx] = child
        drawn = tuple(_drawn_shapes(tree)) if tree is not None else ()
        return cls(name, root, drawables, drawn, placeholders, frozenset(ambiguous))

    @property
    def shows_master_shapes(self) -> bool:
        return _shows_master_shapes(self.root)

    def match(self, idx: str, placeholder_type: str) -> ET.Element | None:
        if idx != _SPECIAL_IDX and idx not in self.ambiguous and idx in self.placeholders:
            return self.placeholders[idx]
        matches = [
            child
            for child in self.placeholders.values()
            if _placeholder_type(placeholder_element(child)) == placeholder_type
        ]
        return matches[0] if len(matches) == 1 else None

    def match_type(self, placeholder_type: str) -> ET.Element | None:
        """The master placeholder a layout placeholder inherits from: by type, never idx.

        A master has one placeholder per kind (title, body, date, footer, number), and
        its idx values mean nothing to the layout: a content placeholder with idx 2 on
        a layout is not the master's date placeholder that happens to have idx 2.
        """

        wanted = _MASTER_TYPES.get(placeholder_type, placeholder_type)
        for child in self.placeholders.values():
            found = _placeholder_type(placeholder_element(child)) or _DEFAULT_TYPE
            if _MASTER_TYPES.get(found, found) == wanted:
                return child
        return None


# The master placeholder kind each placeholder type takes its position and style from.
_MASTER_TYPES = {"ctrTitle": "title", "subTitle": "body", "obj": "body"}


def _placeholder_type(placeholder: ET.Element | None) -> str | None:
    if placeholder is None:
        return None
    return placeholder.attrib.get("type", _DEFAULT_TYPE)


def _shows_master_shapes(root: ET.Element) -> bool:
    return str(root.attrib.get("showMasterSp", "")).strip().casefold() not in {"0", "false"}


class SlideTemplate:
    """The layout and master a slide inherits placeholders, styles and shapes from."""

    def __init__(
        self,
        layout: _TemplatePart | None,
        master: _TemplatePart | None,
        shows_master_shapes: bool,
        slide_size: tuple[int, int] | None = None,
    ) -> None:
        self._layout = layout
        self._master = master
        self._shows_master_shapes = shows_master_shapes
        self._slide_size = slide_size

    def style(self, shape: ET.Element, own_bbox: BBox | None) -> ShapeStyle:
        placeholder = placeholder_element(shape)
        layout_shape = master_shape = None
        placeholder_type = None
        if placeholder is not None:
            idx = _normalize_idx(placeholder.attrib.get("idx"))
            slide_type = placeholder.attrib.get("type", _DEFAULT_TYPE)
            if self._layout is not None:
                layout_shape = self._layout.match(idx, slide_type)
            layout_placeholder = placeholder_element(layout_shape)
            master_type = (
                layout_placeholder.attrib.get("type", slide_type)
                if layout_placeholder is not None
                else slide_type
            )
            if self._master is not None:
                master_shape = self._master.match_type(master_type)
            chain = (placeholder, layout_placeholder, placeholder_element(master_shape))
            placeholder_type = next(
                (
                    candidate.attrib["type"]
                    for candidate in chain
                    if candidate is not None and "type" in candidate.attrib
                ),
                _DEFAULT_TYPE,
            )

        bbox, position_source = _inherited_bbox(shape, layout_shape, master_shape, own_bbox)
        font_pt = next(
            (
                size
                for candidate in (shape, layout_shape, master_shape)
                if (size := _max_font_pt(candidate)) is not None
            ),
            None,
        )
        if font_pt is None:
            font_pt = self._text_style_font_pt(placeholder_type, _paragraph_level(shape))
        # Bullets are inherited from the template only; sibling paragraphs never lend theirs.
        list_style = next(
            (
                found
                for candidate in (layout_shape, master_shape)
                if (found := _shape_list_style(candidate)) is not None
            ),
            None,
        )
        if list_style is None:
            style = self._text_style(placeholder_type)
            level = _paragraph_level(shape)
            kind = _list_kind(
                style.find(f"a:lvl{level + 1}pPr", NS) if style is not None else None
            )
            list_style = (kind, 0) if kind is not None else None
        return ShapeStyle(
            placeholder_type=placeholder_type,
            bbox=bbox,
            position_source=position_source,
            font_pt=font_pt,
            list_kind=list_style[0] if list_style else None,
            list_level=list_style[1] if list_style else 0,
        )

    def inherited_text_shapes(self) -> list[InheritedShape]:
        """Non-placeholder text shapes drawn on the slide from its layout and master."""

        parts: list[tuple[PartKind, _TemplatePart | None]] = [("layout", self._layout)]
        if self._shows_master_shapes:
            parts.append(("master", self._master))
        inherited: list[InheritedShape] = []
        for kind, part in parts:
            if part is None:
                continue
            for child, bbox in part.drawn:
                if _local_name(child.tag) != "sp" or placeholder_element(child) is not None:
                    continue
                text = shape_text(child)
                if text and not _is_prompt_text(text):
                    inherited.append(InheritedShape(kind, part.name, child, bbox))
        return inherited

    def inherited_pictures(self) -> list[InheritedShape]:
        """Pictures drawn on the slide from its layout and master (logos, artwork)."""

        parts: list[tuple[PartKind, _TemplatePart | None]] = [("layout", self._layout)]
        if self._shows_master_shapes:
            parts.append(("master", self._master))
        return [
            InheritedShape(kind, part.name, child, bbox)
            for kind, part in parts
            if part is not None
            for child, bbox in part.drawn
            if _local_name(child.tag) == "pic"
            and placeholder_element(child) is None
            and not _is_rule(child)
            and not self._is_backdrop(bbox)
        ]

    def _is_backdrop(self, bbox: BBox | None) -> bool:
        """A template picture covering most of the slide is its background art."""

        if bbox is None or self._slide_size is None:
            return False
        width, height = self._slide_size
        x1, y1, x2, y2 = bbox
        covered = max(0, min(x2, width) - max(x1, 0)) * max(0, min(y2, height) - max(y1, 0))
        return covered >= BACKDROP_COVERAGE * width * height

    def _text_style(self, placeholder_type: str | None) -> ET.Element | None:
        """The master's p:txStyles entry that applies to a placeholder type."""

        if self._master is None:
            return None
        if placeholder_type in TITLE_TYPES:
            name = "titleStyle"
        elif placeholder_type == "body":
            name = "bodyStyle"
        else:
            name = "otherStyle"
        return self._master.root.find(f"./p:txStyles/p:{name}", NS)

    def _text_style_font_pt(self, placeholder_type: str | None, level: int) -> float | None:
        style = self._text_style(placeholder_type)
        if style is None:
            return None
        for path in (f"a:lvl{level + 1}pPr/a:defRPr", "a:defPPr/a:defRPr"):
            properties = style.find(path, NS)
            size = _int(properties.attrib.get("sz")) if properties is not None else None
            if size is not None and size > 0:
                return size / 100.0
        return None


def _is_rule(picture: ET.Element) -> bool:
    """A picture far thinner than it is long is a decorative line or band."""

    ext = picture.find("./p:spPr/a:xfrm/a:ext", NS)
    if ext is None:
        return False
    width, height = _int(ext.attrib.get("cx")) or 0, _int(ext.attrib.get("cy")) or 0
    longer, shorter = max(width, height), min(width, height)
    return longer > 0 and shorter / longer < 0.03


def _inherited_bbox(
    shape: ET.Element,
    layout_shape: ET.Element | None,
    master_shape: ET.Element | None,
    own_bbox: BBox | None,
) -> tuple[BBox | None, PartKind]:
    if layout_shape is None and master_shape is None:
        return own_bbox, "slide"
    sources: list[tuple[PartKind, dict[str, int | None]]] = [
        ("slide", _xfrm_components(shape)),
        ("layout", _xfrm_components(layout_shape)),
        ("master", _xfrm_components(master_shape)),
    ]
    values: dict[str, int] = {}
    used: set[PartKind] = set()
    for key in ("x", "y", "cx", "cy"):
        for source, components in sources:
            value = components[key]
            if value is not None:
                values[key] = value
                used.add(source)
                break
    if len(values) < 4 or used == {"slide"}:
        return own_bbox, "slide"
    bbox = (values["x"], values["y"], values["x"] + values["cx"], values["y"] + values["cy"])
    return bbox, "master" if "master" in used else "layout"


class TemplateResolver:
    """Resolves and caches the layout and master parts used by a package's slides."""

    def __init__(self, package: OOXMLPackage) -> None:
        self._package = package
        self._parts: dict[str, _TemplatePart] = {}
        self._slide_size = package.slide_size()

    def for_slide(self, slide_part: str, slide_root: ET.Element) -> SlideTemplate:
        layout_name = self._related(slide_part, "slideLayout")
        master_name = self._related(layout_name, "slideMaster") if layout_name else None
        layout = self._part(layout_name)
        master = self._part(master_name)
        shows_master = _shows_master_shapes(slide_root) and (
            layout is None or layout.shows_master_shapes
        )
        return SlideTemplate(layout, master, shows_master, self._slide_size)

    def _related(self, part_name: str, relationship_type: str) -> str | None:
        for relationship in self._package.relationships(part_name).values():
            if relationship.external:
                continue
            if relationship.type.rstrip("/").rsplit("/", 1)[-1] != relationship_type:
                continue
            target = resolve_relationship_target(part_name, relationship)
            if target is not None and self._package.exists(target):
                return target
        return None

    def _part(self, part_name: str | None) -> _TemplatePart | None:
        if part_name is None:
            return None
        if part_name not in self._parts:
            self._parts[part_name] = _TemplatePart.from_root(
                part_name, self._package.xml(part_name)
            )
        return self._parts[part_name]
