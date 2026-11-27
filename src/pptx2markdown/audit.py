"""Self-check: text on the slides that the converted document does not contain.

The parser is deterministic, so silent loss is its worst failure: a shape it does
not understand simply disappears. ``audit`` reads the slide XML again on its own,
without the parser's rules, and checks that every piece of typed text appears in
the document's version of that slide. It needs no model and is cheap enough to run
on every conversion; an agent can use its findings as the signal to look closer.

Text is compared as letters and digits only, so escaping, emphasis, list markers,
spacing and table line breaks never count as loss. Footer, date and slide-number
placeholders are left out on purpose and not audited, and neither are charts,
SmartArt or equations, which are converted from other parts or into LaTeX.
"""

from __future__ import annotations

import html
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .inheritance import FOOTER_TYPES, placeholder_element
from .model import PresentationDocument, SlideDocument
from .ooxml import OOXMLPackage, PackageLimits
from .shapes import NS, flatten_shapes, local_name

# Document fields that carry slide content; ids, kinds and geometry do not.
_TEXT_FIELDS = frozenset({"text", "latex", "alt_text", "description", "markdown", "label"})
# A Markdown link or image destination: "[label](url)" keeps only its label.
_DESTINATION = re.compile(r"\]\((?:<[^>]*>|[^)\s]*)\)")
# Chunks shorter than this (after normalising) are too short to tell anything.
MIN_CHUNK = 2


@dataclass(frozen=True)
class TextLoss:
    """Text typed in a slide shape that the document's slide does not contain."""

    slide: int
    shape_id: str | None
    text: str


@dataclass(frozen=True)
class AuditReport:
    """What the document covers of the text typed on the package's slides."""

    slides_in_package: int
    slides_audited: int
    chunks: int
    covered_chunks: int
    losses: tuple[TextLoss, ...]
    # Pages of slides the package has but the document does not (hidden ones excepted).
    missing_slides: tuple[int, ...] = ()

    @property
    def coverage(self) -> float:
        return self.covered_chunks / self.chunks if self.chunks else 1.0

    @property
    def ok(self) -> bool:
        return not self.losses and not self.missing_slides

    def to_data(self) -> dict[str, Any]:
        return {
            "ok": self.ok,
            "coverage": round(self.coverage, 4),
            "slides_in_package": self.slides_in_package,
            "slides_audited": self.slides_audited,
            "chunks": self.chunks,
            "covered_chunks": self.covered_chunks,
            "missing_slides": list(self.missing_slides),
            "losses": [
                {"slide": loss.slide, "shape_id": loss.shape_id, "text": loss.text}
                for loss in self.losses
            ],
        }


def normalize(text: str) -> str:
    """Letters and digits only, case-folded: what must survive any rendering."""

    return "".join(
        character for character in html.unescape(text).casefold() if character.isalnum()
    )


def _paragraph_chunks(paragraph: ET.Element) -> list[str]:
    """Runs of typed text, split wherever something other than text interrupts them."""

    chunks: list[str] = []
    current: list[str] = []
    for child in paragraph:
        name = local_name(child.tag)
        if name in {"r", "fld"}:
            if name == "fld" and (child.attrib.get("type") or "").casefold() == "slidenum":
                continue
            current.extend(node.text or "" for node in child.iter(f"{{{NS['a']}}}t"))
        elif name == "br":
            current.append(" ")
        elif name not in {"pPr", "endParaRPr"}:
            chunks.append("".join(current))
            current = []
    chunks.append("".join(current))
    return [chunk.strip() for chunk in chunks if chunk.strip()]


def _shape_chunks(element: ET.Element) -> list[str]:
    placeholder = placeholder_element(element)
    if placeholder is not None and placeholder.attrib.get("type") in FOOTER_TYPES:
        return []
    # A cell covered by a merge is not shown, whatever text it still holds.
    covered = {
        id(paragraph)
        for cell in element.iter(f"{{{NS['a']}}}tc")
        if cell.attrib.get("hMerge") in {"1", "true"} or cell.attrib.get("vMerge") in {"1", "true"}
        for paragraph in cell.iter(f"{{{NS['a']}}}p")
    }
    return [
        chunk
        for paragraph in element.iter(f"{{{NS['a']}}}p")
        if id(paragraph) not in covered
        for chunk in _paragraph_chunks(paragraph)
    ]


def _strings(value: Any, key: str | None = None) -> list[str]:
    if isinstance(value, dict):
        return [text for name, item in value.items() for text in _strings(item, name)]
    if isinstance(value, (list, tuple)):
        return [text for item in value for text in _strings(item, key)]
    if isinstance(value, str) and key in _TEXT_FIELDS:
        return [value]
    return []


def slide_text(slide: SlideDocument) -> str:
    """Every piece of content text of a document slide, normalised and concatenated."""

    data = slide.model_dump()
    text = " ".join(_strings(data.get("blocks", [])) + [slide.notes or ""])
    return normalize(_DESTINATION.sub("]", text).replace("<br>", " "))


def audit(source: str | Path | bytes, document: PresentationDocument) -> AuditReport:
    """Check ``document`` against the text typed on the slides of ``source``."""

    if isinstance(source, (str, Path)):
        source = Path(source).expanduser()
    pages = {slide.page: slide for slide in document.slides}
    losses: list[TextLoss] = []
    missing_slides: list[int] = []
    chunks = covered = audited = 0
    with OOXMLPackage(source, PackageLimits(), name=document.source.name) as package:
        entries = package.slide_entries()
        for entry in entries:
            root = package.xml(entry.part_name)
            slide = pages.get(entry.page)
            if slide is None:
                hidden = str(root.attrib.get("show", "1")).strip().casefold() in {"0", "false"}
                if not hidden:
                    missing_slides.append(entry.page)
                continue
            tree = root.find("./p:cSld/p:spTree", NS)
            if tree is None:
                continue
            audited += 1
            output = slide_text(slide)
            for shape in flatten_shapes(tree):
                for chunk in _shape_chunks(shape.element):
                    normalized = normalize(chunk)
                    if len(normalized) < MIN_CHUNK:
                        continue
                    chunks += 1
                    if normalized in output:
                        covered += 1
                    else:
                        losses.append(TextLoss(entry.page, shape.shape_id, chunk))
    return AuditReport(
        slides_in_package=len(entries),
        slides_audited=audited,
        chunks=chunks,
        covered_chunks=covered,
        losses=tuple(losses),
        missing_slides=tuple(missing_slides),
    )
