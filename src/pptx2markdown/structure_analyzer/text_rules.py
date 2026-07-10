from __future__ import annotations

import re
from typing import Optional


def normalize_text(text: str) -> str:
    value = text.lower().replace("\n", " ")
    value = re.sub(r"[^0-9a-z\uac00-\ud7a3\.\s]", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def is_numbered_heading_text(text: str) -> bool:
    value = normalize_text(text)
    return bool(
        re.match(
            r"^(?:\d+\s*\.\s*\d+(?:\s*\.\s*\d+)*\s*\.?|\d+\s*\.)\s+",
            value,
        )
    )


def strip_leading_heading_markers(text: str) -> str:
    raw = re.sub(r"\s+", " ", (text or "").strip())
    if not raw:
        return ""
    return re.sub(r"^(?:[-*•▶√]+\s*)+", "", raw).strip()


def numbered_suggested_depth(
    text: str,
    last_section_depth: Optional[int],
    seen_title: bool,
) -> Optional[int]:
    raw = strip_leading_heading_markers(text)
    if not raw:
        return None
    if re.match(r"^\d+\s*\.\s*\d+(?:\s*\.\s*\d+)*\s*\.?\s+", raw):
        if last_section_depth is not None:
            return min(last_section_depth + 1, 3)
        return 2 if seen_title else 1
    if re.match(r"^\d+\s*\.\s+", raw):
        return 2
    return None


def numbered_heading_kind(text: str) -> Optional[str]:
    value = normalize_text(text)
    if re.match(r"^\d+\s*\.\s*\d+(?:\s*\.\s*\d+)*\s*\.?\s+", value):
        return "dotted-multi"
    if re.match(r"^\d+\s*\.\s+", value):
        return "dotted-single"
    return None
