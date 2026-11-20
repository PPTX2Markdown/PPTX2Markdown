"""Heading detection rules ported from the legacy converter.

Detection runs in two passes over a slide's shapes in reading order:

1. ``heading_hints`` proposes a depth for title-like and numbered shapes using
   placeholder types, position in the title band and font-size bands.
2. ``HeadingDetector.classify`` makes the final per-shape decision, adding
   placeholder overrides, font/numbering fallbacks and duplicate suppression.
"""

from __future__ import annotations

import re
import statistics
import unicodedata
from dataclasses import dataclass
from typing import Literal, Sequence

TITLE_TYPES = frozenset({"title", "ctrTitle", "subTitle"})
_PLACEHOLDER_DEPTHS = {"title": 1, "ctrTitle": 1, "subTitle": 2}
# Why a block became a heading: a title placeholder, a placeholder in the layout's
# title band, section numbering, heading-like text, or font size alone.
HeadingSource = Literal["placeholder", "position", "numbering", "text", "font"]
_THRESHOLD = 0.7
_MIN_FALLBACK_FONT_PT = 24.0
_MIN_TITLE_BAND_EMU = 600_000
_BULLET_PREFIXES = ("▶", "-", "*", "√")
_MULTI_LEVEL_NUMBER = re.compile(r"^\d+\s*\.\s*\d+(?:\s*\.\s*\d+)*\s*\.?\s+")
_SINGLE_NUMBER = re.compile(r"^\d+\s*\.\s+")
_ANY_NUMBER_PREFIX = re.compile(r"^\d+(?:\s*\.\s*\d+)*\s*\.?\s+")
_NUMBERED_HEADING = re.compile(r"^(?:\d+\s*\.\s*\d+(?:\s*\.\s*\d+)*\s*\.?|\d+\s*\.)\s+")
_NUMBER_MARKER = re.compile(r"(?:^|\s)\d+\s*\.\s+")
_LEADING_MARKERS = re.compile(r"^(?:[-*•▶√]+\s*)+")

BBox = tuple[int, int, int, int]


def _collapse(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _fold(text: str) -> str:
    """Case- and width-folded text keeping only letters, numbers, dots and spaces."""

    value = unicodedata.normalize("NFKC", text).casefold().replace("\n", " ")
    value = "".join(
        char
        if char == "." or char.isspace() or unicodedata.category(char).startswith(("L", "N"))
        else " "
        for char in value
    )
    return _collapse(value)


def _strip_markers(text: str) -> str:
    raw = _collapse(text)
    return _LEADING_MARKERS.sub("", raw).strip() or raw


def is_numbered_heading(text: str) -> bool:
    return bool(_NUMBERED_HEADING.match(_fold(text)))


def looks_like_heading(text: str, placeholder_type: str | None) -> bool:
    if text.strip().startswith(_BULLET_PREFIXES):
        return False
    folded = _fold(text)
    if not folded:
        return False
    return is_numbered_heading(text) or (placeholder_type in TITLE_TYPES and len(folded) <= 80)


@dataclass(frozen=True)
class ShapeSignals:
    """What heading detection needs to know about any shape on a slide."""

    text: str
    placeholder_type: str | None
    font_pt: float | None
    bbox: BBox | None
    position_from_layout: bool = False
    is_footer: bool = False
    is_decorative: bool = False


@dataclass(frozen=True)
class HeadingHint:
    depth: int | None = None
    score: float = 0.0
    source: HeadingSource | None = None

    @property
    def candidate(self) -> bool:
        return self.depth is not None and self.score >= _THRESHOLD


def _title_band_bottom(shapes: Sequence[ShapeSignals]) -> int | None:
    positioned = [s for s in shapes if s.bbox is not None]
    content = [s for s in positioned if not s.is_footer and not s.is_decorative] or positioned
    if not content:
        return None
    top = min(s.bbox[1] for s in content if s.bbox is not None)
    bottom = max(s.bbox[3] for s in content if s.bbox is not None)
    return top + max(_MIN_TITLE_BAND_EMU, int(max(1, bottom - top) * 0.22))


def _is_top_title(shape: ShapeSignals, title_band_bottom: int | None) -> bool:
    if shape.is_footer or shape.is_decorative:
        return False
    if shape.placeholder_type in TITLE_TYPES:
        return True
    if (
        shape.position_from_layout
        and shape.placeholder_type is not None
        and shape.bbox is not None
        and title_band_bottom is not None
        and shape.bbox[1] <= title_band_bottom
    ):
        return len(_fold(shape.text)) <= 100
    return False


def _numbered_depth(text: str, last_section_depth: int | None, seen_title: bool) -> int | None:
    raw = _strip_markers(text) if text.strip() else ""
    if _MULTI_LEVEL_NUMBER.match(raw):
        if last_section_depth is not None:
            return min(last_section_depth + 1, 3)
        return 2 if seen_title else 1
    if _SINGLE_NUMBER.match(raw):
        return 2
    return None


def heading_hints(shapes: Sequence[ShapeSignals]) -> list[HeadingHint]:
    """Propose heading depths for shapes given in reading order."""

    is_heading = [looks_like_heading(s.text, s.placeholder_type) for s in shapes]
    title_band_bottom = _title_band_bottom(shapes)

    heading_fonts = [s.font_pt for s, heading in zip(shapes, is_heading) if heading and s.font_pt]
    tolerance = max(1.0, statistics.median(heading_fonts) * 0.06) if heading_fonts else 1.0
    bands: list[float] = []
    for size in sorted(heading_fonts, reverse=True):
        if not bands or abs(size - bands[-1]) > tolerance:
            bands.append(size)

    def depth_from_font(size: float | None) -> int | None:
        if not size or not bands:
            return None
        closest = min(range(len(bands)), key=lambda i: abs(size - bands[i]))
        return min(closest + 1, 6) if abs(size - bands[closest]) <= tolerance else None

    hints: list[HeadingHint] = []
    seen_title = False
    h1_assigned = False
    seen_numbered = False
    last_section_depth: int | None = None
    previous_depth: int | None = None
    previous_font: float | None = None

    for shape, heading in zip(shapes, is_heading):
        depth: int | None = None
        source: HeadingSource | None = None
        if _is_top_title(shape, title_band_bottom):
            source = "placeholder" if shape.placeholder_type in TITLE_TYPES else "position"
            depth = 2 if h1_assigned else 1
            seen_title = True
            last_section_depth = 1
        elif heading:
            source = "numbering" if is_numbered_heading(shape.text) else "text"
            font_depth = depth_from_font(shape.font_pt)
            if (
                font_depth is not None
                and previous_depth is not None
                and shape.font_pt is not None
                and previous_font is not None
                and abs(shape.font_pt - previous_font) <= tolerance
            ):
                depth = previous_depth
            elif font_depth is not None:
                depth = font_depth
            else:
                depth = 2 if seen_title else 1
                last_section_depth = depth

            suggested = _numbered_depth(shape.text, last_section_depth, seen_title)
            if suggested is not None:
                if not seen_numbered:
                    suggested = min(suggested, 2)
                depth = max(depth, suggested)
                seen_numbered = True

            if h1_assigned and depth == 1:
                depth = 2
            if previous_depth is not None and depth > previous_depth + 1:
                depth = previous_depth + 1
            if depth == 1:
                h1_assigned = True
            previous_depth = depth
            previous_font = shape.font_pt

        if heading:
            score = 0.95 if is_numbered_heading(shape.text) else 0.75
        else:
            score = 0.0
        hints.append(HeadingHint(depth=depth, score=score, source=source))
    return hints


@dataclass(frozen=True)
class TextShape:
    """A text shape's rendered content, as the final heading decision sees it."""

    markdown: str
    plain: str
    placeholder_type: str | None
    font_pt: float | None
    has_list: bool
    has_math: bool
    font_sections: tuple[str, ...] = ()


@dataclass(frozen=True)
class Heading:
    level: int
    text: str
    trailing: str | None = None
    source: HeadingSource = "text"


def _fallback_depth(text: str, text_shape_index: int, font_pt: float | None) -> int | None:
    raw = _collapse(text)
    if not raw or raw.startswith(_BULLET_PREFIXES):
        return None
    if len(_NUMBER_MARKER.findall(raw)) >= 2:
        return None
    if _MULTI_LEVEL_NUMBER.match(raw):
        return 3
    large = font_pt is not None and font_pt >= _MIN_FALLBACK_FONT_PT
    if _SINGLE_NUMBER.match(raw):
        return 2 if large else None
    if text_shape_index == 0 and len(raw) <= 80:
        return 1 if large else None
    return None


class HeadingDetector:
    """Final heading decisions for the text shapes of one slide, in reading order."""

    def __init__(self) -> None:
        self._used: set[str] = set()
        self._text_shapes = 0

    def classify(
        self, shape: TextShape, hint: HeadingHint
    ) -> Heading | Literal["duplicate"] | None:
        """Return the heading, ``"duplicate"`` when the shape repeats one, or None."""

        rendered = _collapse(shape.markdown)
        sections = [_collapse(s) for s in re.split(r"\n{2,}", shape.markdown) if s.strip()]
        if len(sections) == 1 and shape.font_sections:
            sections = list(shape.font_sections)
        title_like = shape.placeholder_type in TITLE_TYPES
        can_split = len(sections) > 1 and title_like and not shape.has_list and not shape.has_math

        depth, score, candidate = hint.depth, hint.score, hint.candidate
        heading_source: HeadingSource = hint.source or "text"
        strong = candidate and depth is not None and 1 <= depth <= 6
        if shape.has_math:
            depth, score, candidate = None, 0.0, False
        else:
            placeholder_depth = _PLACEHOLDER_DEPTHS.get(shape.placeholder_type or "")
            if placeholder_depth is not None:
                depth, score, candidate, strong = placeholder_depth, max(score, 0.9), True, True
                heading_source = "placeholder"
            if not candidate:
                fallback = None
                if shape.has_list:
                    if _MULTI_LEVEL_NUMBER.match(rendered):
                        fallback = 3
                    elif _SINGLE_NUMBER.match(rendered):
                        fallback = 2
                if fallback is None:
                    fallback = _fallback_depth(rendered, self._text_shapes, shape.font_pt)
                if fallback is None:
                    fallback = _fallback_depth(shape.plain, self._text_shapes, shape.font_pt)
                if fallback is not None:
                    depth, score, candidate, strong = fallback, max(score, 0.8), True, True
                    numbered = _ANY_NUMBER_PREFIX.match(rendered) or _ANY_NUMBER_PREFIX.match(
                        _collapse(shape.plain)
                    )
                    heading_source = "numbering" if numbered else "font"

        if not title_like:
            plain = _collapse(shape.plain)
            if len(_NUMBER_MARKER.findall(plain)) >= 2:
                candidate = False
            elif not strong and len(plain) >= 30:
                candidate = False

        if not (candidate and depth is not None and 1 <= depth <= 6 and score >= _THRESHOLD):
            self._text_shapes += 1
            return None

        source = sections[0] if can_split else shape.plain
        if _ANY_NUMBER_PREFIX.match(rendered):
            source = rendered
        text = _strip_markers(source)
        key = _collapse(text)
        if key in self._used:
            return "duplicate"
        self._used.add(key)
        self._text_shapes += 1
        trailing = "\n\n".join(sections[1:]) if can_split else None
        return Heading(level=depth, text=text, trailing=trailing, source=heading_source)
