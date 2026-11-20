"""Text paragraphs: runs to inline Markdown, list structure, and text/list blocks."""

from __future__ import annotations

import re
import unicodedata
import xml.etree.ElementTree as ET
from dataclasses import dataclass, replace

from .inheritance import ShapeStyle
from .inline import CLOSE, OPEN, emphasis, finalize, markup, remove_markup
from .model import BoundingBox, ContentBlock, ListBlock, ListItem, MathBlock, TextBlock
from .omml import omml_to_latex
from .ooxml import OFFICE_REL_NS, Relationship
from .shapes import MATH_TAGS, MATH_WRAPPER_TAG, NS, Shape, attr, integer, local_name


def relationship_target(
    relationships: dict[str, Relationship],
    relationship_id: str,
) -> str | None:
    relationship = relationships.get(relationship_id)
    if relationship is None:
        return None
    return relationship.target if relationship.external else None


def _run_link(run: ET.Element, relationships: dict[str, Relationship]) -> str | None:
    properties = run.find("./a:rPr", NS)
    hyperlink = properties.find("./a:hlinkClick", NS) if properties is not None else None
    rid = hyperlink.attrib.get(f"{{{OFFICE_REL_NS}}}id", "") if hyperlink is not None else ""
    target = relationship_target(relationships, rid)
    if target and re.match(r"^(?:https?://|mailto:)", target, flags=re.IGNORECASE):
        return target
    return None


def _run_emphasis(run: ET.Element) -> tuple[bool, bool]:
    properties = run.find("./a:rPr", NS)
    if properties is None:
        return False, False
    return (
        properties.attrib.get("b") in {"1", "true"},
        properties.attrib.get("i") in {"1", "true"},
    )


def _link_markdown(label: str, target: str) -> str:
    destination = f"<{target}>" if any(c.isspace() or c in "()<>" for c in target) else target
    # Brackets in a label would end the link early unless escaped; the rest is source text.
    label = label.replace("[", markup(r"\[")).replace("]", markup(r"\]"))
    return f"{markup('[')}{label}{markup(f']({destination})')}"


@dataclass(frozen=True)
class Paragraph:
    # Inline Markdown with generated markup wrapped in sentinels (see ``inline``):
    # ``text`` has links and math, ``rich`` also has bold/italic emphasis.
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
    rich: str = ""
    # The number an auto-numbered paragraph starts from (a:buAutoNum startAt).
    start: int = 1
    # Every run carries the same bold/italic: formatting of the whole paragraph.
    uniform_emphasis: bool = False

    @property
    def is_math_only(self) -> bool:
        return bool(self.math) and not remove_markup(self.text).strip()


def _normalize_breaks(text: str) -> str:
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r" *\n *", "\n", text).strip()


_INDENT = re.compile(r"[ \t\u00a0\u3000]")
# A bullet glyph typed as text rather than set as the paragraph's bullet.
# A typed bullet may follow leading generated markup such as an opening emphasis marker.
_TYPED_BULLET_RICH = re.compile(f"^((?:{OPEN}[^{CLOSE}]*{CLOSE})*)▶+\\s*")


def parse_paragraph(
    paragraph: ET.Element,
    relationships: dict[str, Relationship],
) -> tuple[Paragraph, int]:
    """Return the paragraph and how many of its equations could not be converted."""

    parts: list[str] = []
    rich_parts: list[str] = []
    plain_parts: list[str] = []
    math: list[str] = []
    lines: list[tuple[str, list[float]]] = [("", [])]
    has_rich_content = False
    unconverted_math = 0
    # Consecutive runs with the same link and emphasis, merged before markup is added.
    group: list[str] = []
    group_style: tuple[str | None, bool, bool] | None = None
    emphasis_styles: set[tuple[bool, bool]] = set()

    def flush_group() -> None:
        nonlocal group_style
        if group_style is None:
            return
        text = "".join(group)
        link, bold, italic = group_style
        if text.strip():
            emphasis_styles.add((bold, italic))
        plain_markdown = _link_markdown(text, link) if link and text else text
        parts.append(plain_markdown)
        rich_parts.append(emphasis(plain_markdown, bold=bold, italic=italic) if text else "")
        group.clear()
        group_style = None

    def visit(container: ET.Element) -> None:
        nonlocal has_rich_content, unconverted_math, group_style
        for child in container:
            if child.tag == MATH_WRAPPER_TAG:
                visit(child)
                continue
            if child.tag in MATH_TAGS:
                flush_group()
                has_rich_content = True
                latex = omml_to_latex(child)
                if latex:
                    emphasis_styles.add((False, False))
                    math.append(latex)
                    parts.append(markup(f"${latex}$"))
                    rich_parts.append(markup(f"${latex}$"))
                    plain_parts.append(f"${latex}$")
                else:
                    unconverted_math += 1
                continue
            name = local_name(child.tag)
            if name == "fld" and (child.attrib.get("type") or "").casefold() == "slidenum":
                continue
            if name in {"r", "fld"}:
                run_text = "".join(node.text or "" for node in child.findall(".//a:t", NS))
                if not run_text:
                    continue
                link = _run_link(child, relationships)
                if link:
                    has_rich_content = True
                style = (link, *_run_emphasis(child))
                if style != group_style:
                    flush_group()
                    group_style = style
                group.append(run_text)
                plain_parts.append(run_text)
                text, sizes = lines[-1]
                size = integer(attr(child.find("./a:rPr", NS), "sz"), default=-1)
                lines[-1] = (text + run_text, [*sizes, size / 100.0] if size > 0 else sizes)
            elif name == "br":
                flush_group()
                parts.append("\n")
                rich_parts.append("\n")
                plain_parts.append(" ")
                lines.append(("", []))

    visit(paragraph)
    flush_group()
    properties = paragraph.find("./a:pPr", NS)
    bullets_disabled = properties is not None and properties.find("./a:buNone", NS) is not None
    auto_number = properties.find("./a:buAutoNum", NS) if properties is not None else None
    ordered = auto_number is not None
    explicit = ordered or (
        properties is not None and properties.find("./a:buChar", NS) is not None
    )
    level = max(0, min(integer(attr(properties, "lvl")), 8))
    return (
        Paragraph(
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
            rich=_normalize_breaks("".join(rich_parts)),
            start=max(1, integer(attr(auto_number, "startAt"), default=1)),
            uniform_emphasis=len(emphasis_styles) == 1 and emphasis_styles != {(False, False)},
        ),
        unconverted_math,
    )


def apply_inherited_lists(paragraphs: list[Paragraph], style: ShapeStyle) -> list[Paragraph]:
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


def shape_markdown(paragraphs: list[Paragraph]) -> str:
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


def font_sections(paragraphs: list[Paragraph]) -> tuple[str, ...]:
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


def paragraph_blocks(
    paragraphs: list[Paragraph],
    *,
    shape: Shape,
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
    pending_items: list[tuple[str, int]] = []
    pending_ordered = False

    def flush_text() -> None:
        if pending_lines:
            blocks.append(TextBlock(text=finalize("\n".join(pending_lines)), **provenance))
        pending_lines.clear()

    pending_start = 1

    def flush_list() -> None:
        nonlocal pending_ordered
        if pending_items:
            items = tuple(
                ListItem(text=finalize(text), level=level) for text, level in pending_items
            )
            blocks.append(
                ListBlock(
                    items=items,
                    ordered=pending_ordered,
                    start=pending_start if pending_ordered else 1,
                    **provenance,
                )
            )
        pending_items.clear()
        pending_ordered = False

    def add_item(text: str, *, ordered: bool, level: int, start: int = 1) -> None:
        nonlocal pending_ordered, pending_start
        flush_text()
        if pending_items and ordered != pending_ordered:
            flush_list()
        if not pending_items:
            pending_start = start
        pending_ordered = ordered
        pending_items.append((text, level))

    for paragraph in paragraphs:
        if paragraph.is_list:
            add_item(
                paragraph.rich,
                ordered=paragraph.ordered,
                level=paragraph.level,
                start=paragraph.start,
            )
            continue
        if paragraph.is_math_only:
            flush_text()
            flush_list()
            blocks.append(MathBlock(latex=" ".join(paragraph.math), **provenance))
            continue
        typed_item = _TYPED_BULLET_RICH.sub(r"\1", paragraph.rich, count=1)
        if typed_item != paragraph.rich and remove_markup(typed_item).strip():
            add_item(typed_item, ordered=False, level=paragraph.level)
            continue
        if pending_items and paragraph.indented:
            text, level = pending_items[-1]
            pending_items[-1] = (f"{text}\n{paragraph.rich}", level)
            continue
        flush_list()
        pending_lines.append(paragraph.rich)
    flush_text()
    flush_list()
    return blocks


def _cell_list_marker(element: ET.Element, body: ET.Element, level: int) -> str | None:
    """``ul``/``ol`` from the paragraph or the cell's own list style; tables have no template."""

    direct = element.find("./a:pPr", NS)
    if direct is not None and direct.find("./a:buNone", NS) is not None:
        return None
    for properties in (direct, body.find(f"./a:lstStyle/a:lvl{level + 1}pPr", NS)):
        if properties is None:
            continue
        if properties.find("./a:buAutoNum", NS) is not None:
            return "ol"
        if properties.find("./a:buChar", NS) is not None:
            return "ul"
    return None


def _is_stacked(pieces: list[str], *, paragraphs: bool) -> bool:
    """One character per line, set upright in a narrow cell ("구" / "분").

    Full-width characters stack by line breaks or by paragraphs. Narrow letters only
    count when at least three are broken within one paragraph: single letters in
    separate paragraphs are as often one value per row ("M" / "M").
    """

    pieces = [piece.strip() for piece in pieces if piece.strip()]
    if len(pieces) < 2 or any(len(piece) != 1 for piece in pieces):
        return False
    if not any(piece.isalpha() for piece in pieces):
        return False
    if all(unicodedata.east_asian_width(piece) in {"W", "F"} for piece in pieces):
        return True
    return not paragraphs and len(pieces) >= 3


def table_cell_lines(
    cell: ET.Element, relationships: dict[str, Relationship]
) -> tuple[list[str], int]:
    """One display line per paragraph, with list markers; and the unconverted equations."""

    body = cell.find("./a:txBody", NS)
    if body is None:
        return [], 0
    lines: list[str] = []
    # Plain text of lines without list markers, to recognise stacked characters.
    plain_lines: list[str | None] = []
    counters: dict[int, int] = {}
    unconverted = 0
    for element in body.findall("./a:p", NS):
        paragraph, missing = parse_paragraph(element, relationships)
        unconverted += missing
        # A line break inside a cell paragraph wraps a label to fit a narrow column
        # ("전자제품 / 사업부"); only paragraphs become separate lines.
        # Bold over a whole cell is the table's styling, not emphasis.
        source = paragraph.text if paragraph.uniform_emphasis else paragraph.rich
        plain = remove_markup(paragraph.text)
        if _is_stacked(plain.split("\n"), paragraphs=False):
            source = re.sub(r"\s*\n\s*", "", source)
            plain = re.sub(r"\s*\n\s*", "", plain)
        text = re.sub(r"\s+", " ", source).strip()
        if not text:
            continue
        text = finalize(text, line_start=False)
        level = max(0, min(integer(attr(element.find("./a:pPr", NS), "lvl")), 8))
        marker = _cell_list_marker(element, body, level)
        if marker == "ul":
            text = f"{'  ' * level}- {text}"
        elif marker == "ol":
            counters = {key: value for key, value in counters.items() if key <= level}
            counters[level] = counters.get(level, 0) + 1
            text = f"{'  ' * level}{counters[level]}. {text}"
        lines.append(text)
        plain_lines.append(plain.strip() if marker is None else None)
    if None not in plain_lines and _is_stacked(plain_lines, paragraphs=True):  # type: ignore[arg-type]
        return ["".join(lines)], unconverted
    return lines, unconverted
