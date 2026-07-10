from __future__ import annotations

import re
from dataclasses import dataclass
from typing import List, Optional


@dataclass(frozen=True)
class HeadingPolicy:
    strict: bool = False
    strict_threshold: float = 0.88
    non_strict_threshold: float = 0.7
    numeric_heading_min_font_pt: float = 24.0
    first_block_heading_min_font_pt: float = 24.0

    @property
    def threshold(self) -> float:
        return self.strict_threshold if self.strict else self.non_strict_threshold


def infer_heading_depth_fallback(
    text: str,
    text_block_index: int,
    font_pt: Optional[float] = None,
    policy: Optional[HeadingPolicy] = None,
) -> Optional[int]:
    policy = policy or HeadingPolicy(strict=False)
    raw = re.sub(r"\s+", " ", (text or "").strip())
    if not raw:
        return None
    if raw.startswith(("▶", "-", "*", "√")):
        return None
    numeric_markers = re.findall(r"(?:^|\s)\d+\s*\.\s+", raw)
    if len(numeric_markers) >= 2:
        return None
    if re.match(r"^\d+\s*\.\s*\d+(?:\s*\.\s*\d+)*\s*\.?\s+", raw):
        return 3
    if re.match(r"^\d+\s*\.\s+", raw):
        if font_pt is None or font_pt < policy.numeric_heading_min_font_pt:
            return None
        return 2
    if text_block_index == 0 and len(raw) <= 80:
        if font_pt is None or font_pt < policy.first_block_heading_min_font_pt:
            return None
        return 1
    return None


def clean_heading_text_for_render(text: str) -> str:
    raw = re.sub(r"\s+", " ", (text or "").strip())
    if not raw:
        return ""
    cleaned = re.sub(r"^(?:[-*•▶√]+\s*)+", "", raw).strip()
    return cleaned or raw


def looks_like_multi_numbered_items(text: str) -> bool:
    raw = re.sub(r"\s+", " ", (text or "").strip())
    if not raw:
        return False
    markers = re.findall(r"(?:^|\s)\d+\s*\.\s+", raw)
    return len(markers) >= 2


def strict_heading_depth_from_placeholder(ph_type: Optional[str]) -> Optional[int]:
    if ph_type is None:
        return None
    key = str(ph_type).strip().lower()
    if key in {"ctrtitle", "title"}:
        return 1
    if key == "subtitle":
        return 2
    return None


def is_body_like_long_sentence(text: str) -> bool:
    raw = re.sub(r"\s+", " ", (text or "").strip())
    if not raw:
        return False
    return len(raw) >= 30


def normalize_single_heading_to_h1(lines: List[str]) -> List[str]:
    heading_lines = [i for i, line in enumerate(lines) if re.match(r"^#{1,6}\s+", line)]
    if len(heading_lines) != 1:
        return lines
    idx = heading_lines[0]
    m = re.match(r"^(#{2,6})\s+(.*)$", lines[idx])
    if not m:
        return lines
    out = list(lines)
    out[idx] = f"# {m.group(2)}"
    return out
