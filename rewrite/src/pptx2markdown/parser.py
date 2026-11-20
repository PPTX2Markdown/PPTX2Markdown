"""Extraction-free PPTX parser that builds the canonical v2 model."""

from __future__ import annotations

import html
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, replace
from pathlib import Path, PurePosixPath
from typing import Literal

from .assets import AssetLimits, AssetStore
from .charts import chart_markdown
from .errors import (
    AssetLimitError,
    PackageSecurityError,
    PresentationParseError,
    UnsupportedPresentationError,
)
from .headings import (
    HeadingDetector,
    HeadingHint,
    ShapeSignals,
    TextShape,
    heading_hints,
)
from .inheritance import (
    FOOTER_TYPES,
    PartKind,
    ShapeStyle,
    SlideTemplate,
    TemplateResolver,
    is_slide_number_only,
    shape_text,
)
from .model import (
    AttachmentBlock,
    BoundingBox,
    ChartBlock,
    ContentBlock,
    ConversionWarning,
    HeadingBlock,
    ImageBlock,
    ListBlock,
    ListItem,
    MathBlock,
    PresentationDocument,
    SlideDocument,
    SmartArtBlock,
    SourceDocument,
    TableBlock,
    TextBlock,
    UnsupportedBlock,
)
from .omml import omml_to_latex
from .ooxml import (
    OFFICE_REL_NS,
    OOXMLPackage,
    PackageLimits,
    Relationship,
    SlideEntry,
    resolve_relationship_target,
)
from .reading_order import LayoutNode, geometry_reading_order

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
CX_NS = "http://schemas.microsoft.com/office/drawing/2014/chartex"
DGM_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
A14_NS = "http://schemas.microsoft.com/office/drawing/2010/main"
# PowerPoint wraps equations in a14:m, whose local name collides with the OMML matrix m:m.
_MATH_WRAPPER_TAG = f"{{{A14_NS}}}m"
_MATH_TAGS = {f"{{{M_NS}}}oMath", f"{{{M_NS}}}oMathPara"}

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
class ParseOptions:
    """Public parsing behavior with explicit resource limits."""

    include_assets: bool = True
    include_notes: bool = True
    include_hidden: bool = True
    strict: bool = True
    package_limits: PackageLimits = PackageLimits()
    asset_limits: AssetLimits = AssetLimits()


@dataclass(frozen=True)
class ParsedComponents:
    """Internal result consumed by the public ParsedPresentation wrapper."""

    document: PresentationDocument
    assets: AssetStore
    source_path: Path


@dataclass(frozen=True)
class _Transform:
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

    def compose(self, inner: "_Transform") -> "_Transform":
        return _Transform(
            sx=self.sx * inner.sx,
            sy=self.sy * inner.sy,
            tx=self.sx * inner.tx + self.tx,
            ty=self.sy * inner.ty + self.ty,
        )


@dataclass(frozen=True)
class _Shape:
    element: ET.Element
    tag: str
    shape_id: str | None
    bbox: tuple[int, int, int, int] | None
    source_order: int
    source_part: PartKind = "slide"
    # Package part holding the element when it is inherited from a layout or master.
    template_part: str | None = None
    style: ShapeStyle = ShapeStyle()


class _WarningCollector:
    def __init__(self) -> None:
        self._warnings: list[ConversionWarning] = []

    def add(
        self,
        code: str,
        message: str,
        *,
        slide: int | None = None,
        shape_id: str | None = None,
        asset_id: str | None = None,
    ) -> None:
        self._warnings.append(
            ConversionWarning(
                code=code,
                message=message,
                slide=slide,
                shape_id=shape_id,
                asset_id=asset_id,
            )
        )

    def values(self) -> tuple[ConversionWarning, ...]:
        return tuple(
            sorted(
                self._warnings,
                key=lambda warning: (
                    warning.slide or 0,
                    warning.shape_id or "",
                    warning.code,
                    warning.asset_id or "",
                    warning.message,
                ),
            )
        )


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _integer(value: str | None, default: int = 0) -> int:
    try:
        return int(value or default)
    except (TypeError, ValueError):
        return default


def _xfrm_bbox(element: ET.Element) -> tuple[int, int, int, int] | None:
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
    x = _integer(offset.attrib.get("x"))
    y = _integer(offset.attrib.get("y"))
    width = _integer(extent.attrib.get("cx"))
    height = _integer(extent.attrib.get("cy"))
    if width <= 0 or height <= 0:
        return None
    return x, y, x + width, y + height


def _group_transform(element: ET.Element) -> _Transform:
    xfrm = element.find("./p:grpSpPr/a:xfrm", NS)
    if xfrm is None:
        return _Transform()
    offset = xfrm.find("./a:off", NS)
    extent = xfrm.find("./a:ext", NS)
    child_offset = xfrm.find("./a:chOff", NS)
    child_extent = xfrm.find("./a:chExt", NS)
    if None in {offset, extent, child_offset, child_extent}:
        return _Transform()
    assert offset is not None
    assert extent is not None
    assert child_offset is not None
    assert child_extent is not None
    child_width = _integer(child_extent.attrib.get("cx"))
    child_height = _integer(child_extent.attrib.get("cy"))
    if child_width <= 0 or child_height <= 0:
        return _Transform()
    sx = _integer(extent.attrib.get("cx")) / child_width
    sy = _integer(extent.attrib.get("cy")) / child_height
    child_x = _integer(child_offset.attrib.get("x"))
    child_y = _integer(child_offset.attrib.get("y"))
    return _Transform(
        sx=sx,
        sy=sy,
        tx=_integer(offset.attrib.get("x")) - child_x * sx,
        ty=_integer(offset.attrib.get("y")) - child_y * sy,
    )


def _shape_id(element: ET.Element) -> str | None:
    non_visual = element.find(".//p:cNvPr", NS)
    if non_visual is None:
        return None
    return non_visual.attrib.get("id") or None


def _flatten_shapes(sp_tree: ET.Element) -> list[_Shape]:
    shapes: list[_Shape] = []
    source_order = 0

    def walk(parent: ET.Element, transform: _Transform) -> None:
        nonlocal source_order
        for element in list(parent):
            tag = _local_name(element.tag)
            if tag == "AlternateContent":
                choice = next(
                    (
                        child
                        for child in element
                        if _local_name(child.tag) == "Choice" and len(child)
                    ),
                    None,
                )
                fallback = next(
                    (
                        child
                        for child in element
                        if _local_name(child.tag) == "Fallback" and len(child)
                    ),
                    None,
                )
                selected = choice if choice is not None else fallback
                if selected is not None:
                    walk(selected, transform)
                continue
            if tag == "grpSp":
                walk(element, transform.compose(_group_transform(element)))
                continue
            if tag not in {"sp", "pic", "graphicFrame", "cxnSp"}:
                continue
            bbox = _xfrm_bbox(element)
            if bbox is not None:
                bbox = transform.apply(bbox)
            shapes.append(
                _Shape(
                    element=element,
                    tag=tag,
                    shape_id=_shape_id(element),
                    bbox=bbox,
                    source_order=source_order,
                )
            )
            source_order += 1

    walk(sp_tree, _Transform())
    return shapes


def _bbox_model(bbox: tuple[int, int, int, int] | None) -> BoundingBox | None:
    if bbox is None:
        return None
    x1, y1, x2, y2 = bbox
    if x2 <= x1 or y2 <= y1:
        return None
    return BoundingBox(x=x1, y=y1, width=x2 - x1, height=y2 - y1)


def _relationship_target(
    relationships: dict[str, Relationship],
    relationship_id: str,
) -> str | None:
    relationship = relationships.get(relationship_id)
    if relationship is None:
        return None
    return relationship.target if relationship.external else None


def _run_markdown(run: ET.Element, relationships: dict[str, Relationship]) -> str:
    text = "".join(node.text or "" for node in run.findall(".//a:t", NS))
    if not text:
        return ""
    properties = run.find("./a:rPr", NS)
    hyperlink = properties.find("./a:hlinkClick", NS) if properties is not None else None
    rid = hyperlink.attrib.get(f"{{{OFFICE_REL_NS}}}id", "") if hyperlink is not None else ""
    target = _relationship_target(relationships, rid)
    if target and re.match(r"^(?:https?://|mailto:)", target, flags=re.IGNORECASE):
        label = text.replace("\\", "\\\\").replace("[", r"\[").replace("]", r"\]")
        return f"[{label}]({target})"
    return text


@dataclass(frozen=True)
class _Paragraph:
    text: str
    plain: str
    math: tuple[str, ...]
    is_list: bool
    ordered: bool
    level: int
    bullets_disabled: bool
    # Break-separated lines with their run font sizes; None when links or math are present.
    font_lines: tuple[tuple[str, tuple[float, ...]], ...] | None
    # The typed text starts with whitespace, i.e. it is indented by hand.
    indented: bool = False

    @property
    def is_math_only(self) -> bool:
        remainder = self.text
        for latex in self.math:
            remainder = remainder.replace(f"${latex}$", "", 1)
        return bool(self.math) and not remainder.strip()


def _normalize_breaks(text: str) -> str:
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r" *\n *", "\n", text).strip()


_INDENT = re.compile(r"[ \t\u00a0\u3000]")
# A bullet glyph typed as text rather than set as the paragraph's bullet.
_TYPED_BULLET = re.compile(r"^▶+\s*")


def _parse_paragraph(
    paragraph: ET.Element,
    relationships: dict[str, Relationship],
) -> tuple[_Paragraph, int]:
    """Return the paragraph and how many of its equations could not be converted."""

    parts: list[str] = []
    plain_parts: list[str] = []
    math: list[str] = []
    lines: list[tuple[str, list[float]]] = [("", [])]
    has_rich_content = False
    unconverted_math = 0

    def visit(container: ET.Element) -> None:
        nonlocal has_rich_content, unconverted_math
        for child in container:
            if child.tag == _MATH_WRAPPER_TAG:
                visit(child)
                continue
            if child.tag in _MATH_TAGS:
                has_rich_content = True
                latex = omml_to_latex(child)
                if latex:
                    math.append(latex)
                    parts.append(f"${latex}$")
                    plain_parts.append(f"${latex}$")
                else:
                    unconverted_math += 1
                continue
            name = _local_name(child.tag)
            if name == "fld" and (child.attrib.get("type") or "").casefold() == "slidenum":
                continue
            if name in {"r", "fld"}:
                markdown = _run_markdown(child, relationships)
                run_text = "".join(node.text or "" for node in child.findall(".//a:t", NS))
                parts.append(markdown)
                plain_parts.append(run_text)
                if markdown != run_text:
                    has_rich_content = True
                text, sizes = lines[-1]
                size = _integer(_attr(child.find("./a:rPr", NS), "sz"), default=-1)
                lines[-1] = (text + run_text, [*sizes, size / 100.0] if size > 0 else sizes)
            elif name == "br":
                parts.append("\n")
                plain_parts.append(" ")
                lines.append(("", []))

    visit(paragraph)
    properties = paragraph.find("./a:pPr", NS)
    bullets_disabled = properties is not None and properties.find("./a:buNone", NS) is not None
    ordered = properties is not None and properties.find("./a:buAutoNum", NS) is not None
    explicit = ordered or (
        properties is not None and properties.find("./a:buChar", NS) is not None
    )
    level = max(0, min(_integer(_attr(properties, "lvl")), 8))
    return (
        _Paragraph(
            text=_normalize_breaks("".join(parts)),
            plain=re.sub(r"\s+", " ", "".join(plain_parts)).strip(),
            math=tuple(math),
            is_list=not bullets_disabled and (explicit or level > 0),
            ordered=ordered and not bullets_disabled,
            level=level if not bullets_disabled else 0,
            bullets_disabled=bullets_disabled,
            font_lines=(
                None
                if has_rich_content
                else tuple((text.strip(), tuple(sizes)) for text, sizes in lines if text.strip())
            ),
            indented=bool(_INDENT.match("".join(parts))),
        ),
        unconverted_math,
    )


def _attr(element: ET.Element | None, name: str) -> str | None:
    return element.attrib.get(name) if element is not None else None


def _apply_inherited_lists(paragraphs: list[_Paragraph], style: ShapeStyle) -> list[_Paragraph]:
    """Placeholders draw bullets from their layout/master unless a paragraph disables them."""

    if style.list_kind is None:
        return paragraphs
    if style.placeholder_type is None and any(p.is_list for p in paragraphs):
        return paragraphs
    return [
        paragraph
        if paragraph.is_list or paragraph.bullets_disabled
        else replace(
            paragraph, is_list=True, ordered=style.list_kind == "ol", level=style.list_level
        )
        for paragraph in paragraphs
    ]


def _shape_markdown(paragraphs: list[_Paragraph]) -> str:
    """Single-string rendering of a text shape, as heading rules inspect it."""

    lines: list[str] = []
    counters: dict[int, int] = {}
    for paragraph in paragraphs:
        if not paragraph.is_list:
            lines.append(paragraph.text)
            continue
        indent = "  " * paragraph.level
        if paragraph.ordered:
            counters[paragraph.level] = counters.get(paragraph.level, 0) + 1
            for deeper in [level for level in counters if level > paragraph.level]:
                del counters[deeper]
            item = re.sub(r"^\d+\s*\.\s*", "", paragraph.text).strip() or paragraph.text
            lines.append(f"{indent}{counters[paragraph.level]}. {item}")
        else:
            lines.append(f"{indent}- {paragraph.text}")
    return "\n".join(lines).strip()


def _font_sections(paragraphs: list[_Paragraph]) -> tuple[str, ...]:
    """Split a one-paragraph shape whose first line is set much larger than the rest."""

    if len(paragraphs) != 1 or paragraphs[0].font_lines is None:
        return ()
    lines = paragraphs[0].font_lines
    if len(lines) < 2 or not lines[0][1]:
        return ()
    rest = [size for _, sizes in lines[1:] for size in sizes]
    if not rest or max(lines[0][1]) < max(rest) * 1.5:
        return ()
    return (lines[0][0], "\n".join(text for text, _ in lines[1:]))


def _paragraph_blocks(
    paragraphs: list[_Paragraph],
    *,
    shape: _Shape,
    bbox: BoundingBox | None,
) -> list[ContentBlock]:
    """One block per run of alike paragraphs, so a plain text box stays one text block.

    Consecutive plain paragraphs share a ``TextBlock`` (one line each) and consecutive
    list paragraphs share a ``ListBlock``; equations and changes between text and list
    start a new block. A paragraph that starts with a typed ``▶`` is a list item, and a
    paragraph indented with typed spaces right after an item continues that item: it is
    how slides wrap a bullet by hand.
    """

    provenance = {"shape_id": shape.shape_id, "bbox": bbox, "source_part": shape.source_part}
    blocks: list[ContentBlock] = []
    pending_lines: list[str] = []
    pending_items: list[ListItem] = []
    pending_ordered = False

    def flush_text() -> None:
        if pending_lines:
            blocks.append(TextBlock(text="\n".join(pending_lines), **provenance))
        pending_lines.clear()

    def flush_list() -> None:
        nonlocal pending_ordered
        if pending_items:
            blocks.append(
                ListBlock(items=tuple(pending_items), ordered=pending_ordered, **provenance)
            )
        pending_items.clear()
        pending_ordered = False

    def add_item(text: str, *, ordered: bool, level: int) -> None:
        nonlocal pending_ordered
        flush_text()
        if pending_items and ordered != pending_ordered:
            flush_list()
        pending_ordered = ordered
        pending_items.append(ListItem(text=text, level=level))

    for paragraph in paragraphs:
        if paragraph.is_list:
            add_item(paragraph.text, ordered=paragraph.ordered, level=paragraph.level)
            continue
        if paragraph.is_math_only:
            flush_text()
            flush_list()
            blocks.append(MathBlock(latex=" ".join(paragraph.math), **provenance))
            continue
        typed_item = _TYPED_BULLET.sub("", paragraph.text, count=1)
        if typed_item != paragraph.text and typed_item:
            add_item(typed_item, ordered=False, level=paragraph.level)
            continue
        if pending_items and paragraph.indented:
            last = pending_items[-1]
            pending_items[-1] = last.model_copy(update={"text": f"{last.text}\n{paragraph.text}"})
            continue
        flush_list()
        pending_lines.append(paragraph.text)
    flush_text()
    flush_list()
    return blocks


_EXT_PATHS = ("./p:spPr/a:xfrm/a:ext", "./p:grpSpPr/a:xfrm/a:ext", "./p:xfrm/a:ext", ".//a:ext")


def _is_unpositioned_literal_date(shape: _Shape) -> bool:
    """A date placeholder holding typed text but no geometry of its own."""

    element = shape.element
    return (
        shape.style.placeholder_type == "dt"
        and element.find(".//a:fld", NS) is None
        and element.find("./p:spPr/a:xfrm", NS) is None
        and any((node.text or "").strip() for node in element.findall(".//a:t", NS))
    )


def _is_footer(shape: _Shape) -> bool:
    if shape.style.placeholder_type in FOOTER_TYPES and not _is_unpositioned_literal_date(shape):
        return True
    return is_slide_number_only(shape.element)


def _is_decorative(shape: _Shape, text: str) -> bool:
    element = shape.element
    if shape.tag == "cxnSp":
        return True
    blip = element.find(".//a:blip", NS)
    has_image = blip is not None and bool(blip.attrib.get(f"{{{OFFICE_REL_NS}}}embed"))
    ext = next(
        (found for path in _EXT_PATHS if (found := element.find(path, NS)) is not None), None
    )
    zero_size = ext is not None and (
        _integer(_attr(ext, "cx"), default=1) <= 0 or _integer(_attr(ext, "cy"), default=1) <= 0
    )
    if shape.tag == "pic" and zero_size:
        return True
    if shape.tag == "sp" and has_image:
        return zero_size
    if shape.tag == "pic" and shape.style.placeholder_type is not None and not has_image:
        return True
    has_math = any(element.find(f".//m:{name}", NS) is not None for name in ("oMath", "oMathPara"))
    return not text.strip() and not has_math and shape.tag not in {"pic", "graphicFrame"}


def _heading_signals(shape: _Shape) -> ShapeSignals:
    text = shape_text(shape.element)
    return ShapeSignals(
        text=text,
        placeholder_type=shape.style.placeholder_type,
        font_pt=shape.style.font_pt,
        bbox=shape.bbox,
        position_from_layout=shape.style.position_source == "layout",
        is_footer=_is_footer(shape),
        is_decorative=_is_decorative(shape, text),
    )


def _with_template_style(shape: _Shape, template: SlideTemplate) -> _Shape:
    style = template.style(shape.element, shape.bbox)
    styled = replace(shape, style=style, bbox=style.bbox)
    if styled.bbox is None and _is_unpositioned_literal_date(styled):
        # Keep visible typed dates at the top of the reading order instead of appending them.
        styled = replace(styled, bbox=(0, 0, 1, 1))
    return styled


def _markdown_cell(text: str) -> str:
    normalized = re.sub(r"\s+", " ", text).strip()
    return normalized.replace("\\", "\\\\").replace("|", r"\|")


def _table_markdown(table: ET.Element) -> str:
    rows: list[list[str]] = []
    for row in table.findall("./a:tr", NS):
        cells: list[str] = []
        for cell in row.findall("./a:tc", NS):
            paragraphs: list[str] = []
            for paragraph in cell.findall("./a:txBody/a:p", NS):
                text = "".join(node.text or "" for node in paragraph.findall(".//a:t", NS))
                if text.strip():
                    paragraphs.append(text.strip())
            cells.append(_markdown_cell("<br>".join(paragraphs)))
        rows.append(cells)
    if not rows:
        return ""
    width = max(len(row) for row in rows)
    normalized_rows = [row + [""] * (width - len(row)) for row in rows]
    lines = [
        "| " + " | ".join(normalized_rows[0]) + " |",
        "| " + " | ".join("---" for _ in range(width)) + " |",
    ]
    lines.extend("| " + " | ".join(row) + " |" for row in normalized_rows[1:])
    return "\n".join(lines)


class _PresentationParser:
    def __init__(self, source: Path, options: ParseOptions) -> None:
        self.source = source
        self.options = options
        self.assets = AssetStore(options.asset_limits)
        self.warnings = _WarningCollector()

    def parse(self) -> ParsedComponents:
        try:
            with OOXMLPackage(self.source, self.options.package_limits) as package:
                templates = TemplateResolver(package)
                slides: list[SlideDocument] = []
                for entry in package.slide_entries():
                    try:
                        slide = self._parse_slide(package, entry, templates)
                    except Exception as exc:
                        if self.options.strict:
                            raise PresentationParseError(
                                self.source,
                                str(exc),
                                slide_number=entry.page,
                            ) from exc
                        self.warnings.add(
                            "slide-parse-failed",
                            str(exc),
                            slide=entry.page,
                        )
                        slide = SlideDocument(page=entry.page)
                    if self.options.include_hidden or not slide.hidden:
                        slides.append(slide)

                document = PresentationDocument(
                    source=SourceDocument(
                        name=self.source.name,
                        sha256=package.source_sha256,
                    ),
                    assets=self.assets.descriptors(),
                    slides=tuple(slides),
                    warnings=self.warnings.values(),
                )
                return ParsedComponents(document, self.assets, self.source)
        except Exception:
            self.assets.close()
            raise

    def _parse_slide(
        self, package: OOXMLPackage, entry: SlideEntry, templates: TemplateResolver
    ) -> SlideDocument:
        root = package.xml(entry.part_name)
        hidden = str(root.attrib.get("show", "1")).strip().casefold() in {
            "0",
            "false",
            "off",
        }
        tree = root.find("./p:cSld/p:spTree", NS)
        if tree is None:
            raise ValueError("missing p:cSld/p:spTree")
        relationships = package.relationships(entry.part_name)
        template = templates.for_slide(entry.part_name, root)
        shapes = _flatten_shapes(tree)
        shapes.extend(
            _Shape(
                element=inherited.element,
                tag="sp",
                shape_id=_shape_id(inherited.element),
                bbox=_xfrm_bbox(inherited.element),
                source_order=len(shapes) + offset,
                source_part=inherited.source_part,
                template_part=inherited.part_name,
            )
            for offset, inherited in enumerate(template.inherited_text_shapes())
        )
        styled = [_with_template_style(shape, template) for shape in shapes]
        ordered_shapes = geometry_reading_order(
            LayoutNode(shape, shape.source_order, shape.bbox) for shape in styled
        )
        hints = heading_hints([_heading_signals(shape) for shape in ordered_shapes])
        detector = HeadingDetector()
        blocks: list[ContentBlock] = []
        for shape, hint in zip(ordered_shapes, hints):
            if shape.tag in {"sp", "cxnSp"}:
                shape_relationships = (
                    package.relationships(shape.template_part)
                    if shape.template_part is not None
                    else relationships
                )
                blocks.extend(
                    self._text_shape_blocks(entry, shape, shape_relationships, hint, detector)
                )
            else:
                blocks.extend(self._shape_blocks(package, entry, shape, relationships))
        heading_indexes = [i for i, block in enumerate(blocks) if isinstance(block, HeadingBlock)]
        if len(heading_indexes) == 1:
            # A slide's only heading is its title.
            index = heading_indexes[0]
            blocks[index] = blocks[index].model_copy(update={"level": 1})
        notes = self._notes(package, entry.part_name) if self.options.include_notes else None
        return SlideDocument(
            page=entry.page,
            hidden=hidden,
            blocks=tuple(blocks),
            notes=notes,
        )

    def _shape_blocks(
        self,
        package: OOXMLPackage,
        entry: SlideEntry,
        shape: _Shape,
        relationships: dict[str, Relationship],
    ) -> list[ContentBlock]:
        if shape.tag == "pic":
            return self._picture_blocks(package, entry, shape, relationships)
        if shape.tag == "graphicFrame":
            return self._graphic_frame_blocks(package, entry, shape, relationships)
        return []

    def _text_shape_blocks(
        self,
        entry: SlideEntry,
        shape: _Shape,
        relationships: dict[str, Relationship],
        hint: HeadingHint,
        detector: HeadingDetector,
    ) -> list[ContentBlock]:
        body = shape.element.find("./p:txBody", NS)
        if body is None or _is_footer(shape):
            return []
        paragraphs: list[_Paragraph] = []
        unconverted_math = 0
        for element in body.findall("./a:p", NS):
            paragraph, unconverted = _parse_paragraph(element, relationships)
            unconverted_math += unconverted
            if paragraph.text:
                paragraphs.append(paragraph)
        if unconverted_math:
            self.warnings.add(
                "math-unconverted",
                f"{unconverted_math} equation(s) could not be converted to LaTeX",
                slide=entry.page,
                shape_id=shape.shape_id,
            )
        if not paragraphs:
            return []

        paragraphs = _apply_inherited_lists(paragraphs, shape.style)
        heading = detector.classify(
            TextShape(
                markdown=_shape_markdown(paragraphs),
                plain=" ".join(p.plain for p in paragraphs if p.plain),
                placeholder_type=shape.style.placeholder_type,
                font_pt=shape.style.font_pt,
                has_list=any(p.is_list for p in paragraphs),
                has_math=any(p.math for p in paragraphs),
                font_sections=_font_sections(paragraphs),
            ),
            hint,
        )
        bbox = _bbox_model(shape.bbox)
        if heading == "duplicate":
            return []
        if heading is None:
            return _paragraph_blocks(paragraphs, shape=shape, bbox=bbox)
        provenance = {"shape_id": shape.shape_id, "bbox": bbox, "source_part": shape.source_part}
        blocks: list[ContentBlock] = [
            HeadingBlock(text=heading.text, level=heading.level, **provenance)
        ]
        if heading.trailing:
            blocks.append(TextBlock(text=heading.trailing, **provenance))
        return blocks

    def _register_related_asset(
        self,
        package: OOXMLPackage,
        *,
        source_part: str,
        relationship: Relationship,
        kind: Literal["image", "attachment"],
        slide: int,
        shape_id: str | None,
    ):
        if relationship.external:
            self.warnings.add(
                "external-asset-ignored",
                f"external asset was not downloaded: {relationship.target}",
                slide=slide,
                shape_id=shape_id,
            )
            return None
        target = resolve_relationship_target(source_part, relationship)
        if target is None or not package.exists(target):
            self.warnings.add(
                "asset-target-missing",
                f"asset relationship target is unavailable: {relationship.target}",
                slide=slide,
                shape_id=shape_id,
            )
            return None
        try:
            with package.open_part(target) as stream:
                return self.assets.register_stream(
                    stream,
                    logical_key=target,
                    kind=kind,
                    original_name=PurePosixPath(target).name,
                    source_part=target,
                    preserve_payload=self.options.include_assets,
                )
        except AssetLimitError as exc:
            if self.options.strict:
                raise
            self.warnings.add(
                "asset-limit-exceeded",
                str(exc),
                slide=slide,
                shape_id=shape_id,
            )
            return None

    def _picture_blocks(
        self,
        package: OOXMLPackage,
        entry: SlideEntry,
        shape: _Shape,
        relationships: dict[str, Relationship],
    ) -> list[ContentBlock]:
        blocks: list[ContentBlock] = []
        bbox = _bbox_model(shape.bbox)
        blip = shape.element.find(".//a:blip", NS)
        rid = blip.attrib.get(f"{{{OFFICE_REL_NS}}}embed", "") if blip is not None else ""
        relationship = relationships.get(rid)
        if relationship is not None:
            descriptor = self._register_related_asset(
                package,
                source_part=entry.part_name,
                relationship=relationship,
                kind="image",
                slide=entry.page,
                shape_id=shape.shape_id,
            )
            if descriptor is not None:
                properties = shape.element.find(".//p:cNvPr", NS)
                alt_text = (
                    properties.attrib.get("descr") or properties.attrib.get("name") or "image"
                    if properties is not None
                    else "image"
                )
                blocks.append(
                    ImageBlock(
                        asset_id=descriptor.id,
                        alt_text=alt_text,
                        shape_id=shape.shape_id,
                        bbox=bbox,
                    )
                )
        elif rid:
            self.warnings.add(
                "image-relationship-missing",
                f"image relationship not found: {rid}",
                slide=entry.page,
                shape_id=shape.shape_id,
            )

        for tag in ("audioFile", "videoFile"):
            media = shape.element.find(f".//a:{tag}", NS)
            media_rid = (
                media.attrib.get(f"{{{OFFICE_REL_NS}}}link", "") if media is not None else ""
            )
            media_relationship = relationships.get(media_rid)
            if media_relationship is None:
                continue
            descriptor = self._register_related_asset(
                package,
                source_part=entry.part_name,
                relationship=media_relationship,
                kind="attachment",
                slide=entry.page,
                shape_id=shape.shape_id,
            )
            if descriptor is not None:
                blocks.append(
                    AttachmentBlock(
                        asset_id=descriptor.id,
                        label=descriptor.original_name,
                        shape_id=shape.shape_id,
                        bbox=bbox,
                    )
                )
        return blocks

    def _chart_block(
        self,
        package: OOXMLPackage,
        entry: SlideEntry,
        shape: _Shape,
        part_name: str,
        bbox: BoundingBox | None,
    ) -> ChartBlock | UnsupportedBlock | None:
        try:
            conversion = chart_markdown(package, part_name)
        except (PresentationParseError, PackageSecurityError):
            raise
        except Exception as exc:  # third-party converter: report and keep the slide
            self.warnings.add(
                "chart-conversion-failed",
                f"chart conversion failed: {type(exc).__name__}: {exc}",
                slide=entry.page,
                shape_id=shape.shape_id,
            )
            return UnsupportedBlock(
                message="chart conversion failed", shape_id=shape.shape_id, bbox=bbox
            )
        for message in conversion.rejected_workbooks:
            self.warnings.add(
                "chart-workbook-rejected",
                f"embedded chart workbook rejected; using cached chart values: {message}",
                slide=entry.page,
                shape_id=shape.shape_id,
            )
        if not conversion.markdown:
            return None
        return ChartBlock(markdown=conversion.markdown, shape_id=shape.shape_id, bbox=bbox)

    def _graphic_frame_blocks(
        self,
        package: OOXMLPackage,
        entry: SlideEntry,
        shape: _Shape,
        relationships: dict[str, Relationship],
    ) -> list[ContentBlock]:
        bbox = _bbox_model(shape.bbox)
        table = shape.element.find(".//a:tbl", NS)
        if table is not None:
            markdown = _table_markdown(table)
            return (
                [
                    TableBlock(
                        markdown=markdown,
                        shape_id=shape.shape_id,
                        bbox=bbox,
                    )
                ]
                if markdown
                else []
            )

        chart = shape.element.find(".//c:chart", NS)
        if chart is None:
            chart = shape.element.find(".//cx:chart", NS)
        if chart is not None:
            rid = chart.attrib.get(f"{{{OFFICE_REL_NS}}}id", "")
            relationship = relationships.get(rid)
            if relationship is not None and not relationship.external:
                target = resolve_relationship_target(entry.part_name, relationship)
                if target is not None and package.exists(target):
                    block = self._chart_block(package, entry, shape, target, bbox)
                    if block is not None:
                        return [block]
            self.warnings.add(
                "chart-data-unavailable",
                "chart relationship or cached data is unavailable",
                slide=entry.page,
                shape_id=shape.shape_id,
            )
            return [
                UnsupportedBlock(
                    message="chart data unavailable",
                    shape_id=shape.shape_id,
                    bbox=bbox,
                )
            ]

        diagram = shape.element.find(".//dgm:relIds", NS)
        if diagram is not None:
            rid = diagram.attrib.get(f"{{{OFFICE_REL_NS}}}dm", "")
            relationship = relationships.get(rid)
            if relationship is not None and not relationship.external:
                target = resolve_relationship_target(entry.part_name, relationship)
                if target is not None and package.exists(target):
                    root = package.xml(target)
                    texts = [
                        re.sub(r"\s+", " ", node.text or "").strip()
                        for node in root.findall(".//a:t", NS)
                    ]
                    texts = [text for text in texts if text]
                    if texts:
                        markdown = "\n".join(f"- {text}" for text in texts)
                        return [
                            SmartArtBlock(
                                markdown=markdown,
                                shape_id=shape.shape_id,
                                bbox=bbox,
                            )
                        ]
            self.warnings.add(
                "smartart-data-unavailable",
                "SmartArt data relationship is unavailable",
                slide=entry.page,
                shape_id=shape.shape_id,
            )
            return [
                UnsupportedBlock(
                    message="SmartArt data unavailable",
                    shape_id=shape.shape_id,
                    bbox=bbox,
                )
            ]

        ole = shape.element.find(".//p:oleObj", NS)
        if ole is not None:
            rid = ole.attrib.get(f"{{{OFFICE_REL_NS}}}id", "")
            relationship = relationships.get(rid)
            if relationship is not None:
                descriptor = self._register_related_asset(
                    package,
                    source_part=entry.part_name,
                    relationship=relationship,
                    kind="attachment",
                    slide=entry.page,
                    shape_id=shape.shape_id,
                )
                if descriptor is not None:
                    return [
                        AttachmentBlock(
                            asset_id=descriptor.id,
                            label=ole.attrib.get("name") or descriptor.original_name,
                            shape_id=shape.shape_id,
                            bbox=bbox,
                        )
                    ]

        model = next(
            (element for element in shape.element.iter() if _local_name(element.tag) == "model3d"),
            None,
        )
        if model is not None:
            model_blocks: list[ContentBlock] = []
            preview = next(
                (element for element in model.iter() if _local_name(element.tag) == "blip"),
                None,
            )
            preview_rid = (
                preview.attrib.get(f"{{{OFFICE_REL_NS}}}embed", "") if preview is not None else ""
            )
            preview_relationship = relationships.get(preview_rid)
            if preview_relationship is not None:
                preview_descriptor = self._register_related_asset(
                    package,
                    source_part=entry.part_name,
                    relationship=preview_relationship,
                    kind="image",
                    slide=entry.page,
                    shape_id=shape.shape_id,
                )
                if preview_descriptor is not None:
                    properties = shape.element.find(".//p:cNvPr", NS)
                    alt_text = (
                        properties.attrib.get("descr")
                        or properties.attrib.get("name")
                        or "3D model preview"
                        if properties is not None
                        else "3D model preview"
                    )
                    model_blocks.append(
                        ImageBlock(
                            asset_id=preview_descriptor.id,
                            alt_text=alt_text,
                            shape_id=shape.shape_id,
                            bbox=bbox,
                        )
                    )
            rid = model.attrib.get(f"{{{OFFICE_REL_NS}}}embed", "")
            relationship = relationships.get(rid)
            if relationship is not None:
                descriptor = self._register_related_asset(
                    package,
                    source_part=entry.part_name,
                    relationship=relationship,
                    kind="attachment",
                    slide=entry.page,
                    shape_id=shape.shape_id,
                )
                if descriptor is not None:
                    model_blocks.append(
                        AttachmentBlock(
                            asset_id=descriptor.id,
                            label=descriptor.original_name,
                            shape_id=shape.shape_id,
                            bbox=bbox,
                        )
                    )
            if model_blocks:
                return model_blocks

        self.warnings.add(
            "unsupported-graphic-frame",
            "graphic frame type is not supported",
            slide=entry.page,
            shape_id=shape.shape_id,
        )
        return [
            UnsupportedBlock(
                message="unsupported graphic frame",
                shape_id=shape.shape_id,
                bbox=bbox,
            )
        ]

    def _notes(self, package: OOXMLPackage, slide_part: str) -> str | None:
        relationships = package.relationships(slide_part)
        notes_relationship = next(
            (
                relationship
                for relationship in relationships.values()
                if relationship.type.endswith("/notesSlide") and not relationship.external
            ),
            None,
        )
        if notes_relationship is None:
            return None
        target = resolve_relationship_target(slide_part, notes_relationship)
        if target is None or not package.exists(target):
            return None
        root = package.xml(target)
        paragraphs: list[str] = []
        for shape in root.findall(".//p:sp", NS):
            placeholder = shape.find("./p:nvSpPr/p:nvPr/p:ph", NS)
            placeholder_type = (
                placeholder.attrib.get("type", "body") if placeholder is not None else ""
            )
            if placeholder_type in {"hdr", "ftr", "dt", "sldNum"}:
                continue
            for paragraph in shape.findall("./p:txBody/a:p", NS):
                text = "".join(node.text or "" for node in paragraph.findall(".//a:t", NS))
                text = re.sub(r"\s+", " ", html.unescape(text)).strip()
                if text:
                    paragraphs.append(text)
        return "\n".join(paragraphs) or None


def parse_components(
    path: str | Path,
    *,
    options: ParseOptions | None = None,
) -> ParsedComponents:
    """Parse one local PPTX into canonical components."""

    source = Path(path).expanduser()
    if source.suffix.casefold() != ".pptx":
        raise UnsupportedPresentationError("only local .pptx files are supported")
    if not source.is_file():
        raise FileNotFoundError(source)
    source = source.resolve()
    try:
        return _PresentationParser(source, options or ParseOptions()).parse()
    except (PresentationParseError, PackageSecurityError, AssetLimitError):
        raise
    except Exception as exc:
        raise PresentationParseError(source, str(exc)) from exc
