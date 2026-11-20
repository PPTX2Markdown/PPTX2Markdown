"""Inline Markdown: generated markup versus source text that must be escaped.

The parser writes links, inline math and emphasis as Markdown while it still runs
text-based rules (headings, typed bullets) over the paragraph. Generated markup is
wrapped in private-use sentinels so those rules see the same text as before, and
``finalize`` later escapes only the source characters and drops the sentinels.
"""

from __future__ import annotations

import re

OPEN = ""
CLOSE = ""
_SPANS = re.compile(f"{OPEN}([^{CLOSE}]*){CLOSE}")


def markup(text: str) -> str:
    """Mark generated Markdown so ``finalize`` keeps it verbatim."""

    return f"{OPEN}{text}{CLOSE}" if text else ""


def strip_markup(text: str) -> str:
    """Drop sentinels but keep the generated Markdown, unescaped."""

    return text.replace(OPEN, "").replace(CLOSE, "")


def remove_markup(text: str) -> str:
    """Drop generated Markdown entirely, leaving only source text."""

    return _SPANS.sub("", text)


_HTML_START = re.compile(r"<(?=[A-Za-z/!?])")
_ENTITY = re.compile(r"&(?=#?[0-9A-Za-z]+;)")
_LINK_LIKE = re.compile(r"\[(?=[^\]\n]*\][(\[:])")
_MARKER_LIKE = re.compile(r"\[(?=(?:Page_\d+|Speaker_Notes|unsupported:|attachment:)[^\]\n]*\])")

_BLOCK_STARTS = (
    re.compile(r"^(#{1,6})(?=\s|$)"),
    re.compile(r"^(>)"),
    re.compile(r"^([-+*])(?=\s|$)"),
    re.compile(r"^(?:[*_-][ \t]*){3,}$"),
    re.compile(r"^[-=]+[ \t]*$"),
    re.compile(r"^(?:`{3,}|~{3,})"),
)
_ORDERED_START = re.compile(r"^(\d{1,9})([.)])(?=\s|$)")


def _is_word(char: str) -> bool:
    return char.isalnum()


def _escape_inline(text: str, *, dollars: bool) -> str:
    out: list[str] = []
    for index, char in enumerate(text):
        before = text[index - 1] if index else " "
        after = text[index + 1] if index + 1 < len(text) else " "
        if char in "\\`":
            out.append("\\" + char)
        elif char == "*":
            out.append("*" if before.isspace() and after.isspace() else r"\*")
        elif char == "_":
            # Intraword underscores never open emphasis (snake_case stays readable).
            out.append("_" if _is_word(before) and _is_word(after) else r"\_")
        elif char == "~" and (before == "~" or after == "~"):
            out.append(r"\~")
        elif char == "$" and dollars:
            out.append(r"\$")
        else:
            out.append(char)
    escaped = "".join(out)
    escaped = _HTML_START.sub(r"\\<", escaped)
    escaped = _ENTITY.sub(r"\\&", escaped)
    escaped = _LINK_LIKE.sub(r"\\[", escaped)
    return _MARKER_LIKE.sub(r"\\[", escaped)


def _escape_line_start(line: str) -> str:
    match = _ORDERED_START.match(line)
    if match:
        return f"{match.group(1)}\\{line[len(match.group(1)) :]}"
    for pattern in _BLOCK_STARTS:
        if pattern.match(line):
            return "\\" + line
    return line


def _finalize_line(line: str, *, line_start: bool) -> str:
    pieces = re.split(f"({OPEN}[^{CLOSE}]*{CLOSE})", line)
    source = "".join(piece for piece in pieces if not piece.startswith(OPEN))
    # Two or more dollar signs could pair up as inline math in renderers that support it.
    dollars = source.count("$") >= 2
    rendered: list[str] = []
    for position, piece in enumerate(pieces):
        if piece.startswith(OPEN):
            rendered.append(piece[1:-1])
            continue
        if not piece:
            continue
        escaped = _escape_inline(piece, dollars=dollars)
        if line_start and position == 0:
            escaped = _escape_line_start(escaped)
        rendered.append(escaped)
    return "".join(rendered)


def finalize(text: str, *, line_start: bool = True) -> str:
    """Escape source text so it renders literally, keeping generated markup.

    ``line_start`` also escapes characters that would start a Markdown block
    (headings, quotes, lists, rules, fences) at the beginning of each line; table
    cells and headings, whose content is inline, do not need it.
    """

    lines = text.split("\n")
    return "\n".join(_finalize_line(line, line_start=line_start) for line in lines)


def finalize_heading(text: str) -> str:
    """Escape heading text, including a trailing ``#`` run CommonMark would drop."""

    escaped = finalize(" ".join(text.split("\n")), line_start=False)
    return re.sub(r"(\s)(#+)\s*$", lambda m: f"{m.group(1)}\\{m.group(2)}", escaped)


def escape_text(text: str) -> str:
    """Escape plain text (no generated markup) for a Markdown block."""

    return finalize(text.replace(OPEN, "").replace(CLOSE, ""))


def emphasis(text: str, *, bold: bool, italic: bool) -> str:
    """Wrap source text in emphasis markers, keeping outer whitespace outside them."""

    if not (bold or italic) or not text.strip():
        return text
    marker = "***" if bold and italic else "**" if bold else "*"
    stripped = text.strip()
    leading = text[: len(text) - len(text.lstrip())]
    trailing = text[len(text.rstrip()) :]
    return f"{leading}{markup(marker)}{stripped}{markup(marker)}{trailing}"
