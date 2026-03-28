"""Markdown cleanup helpers shared across providers."""

from __future__ import annotations

import re


_PROMPT_ECHO_PHRASES = {
    "return markdown only.",
    "extract only document-worthy information from this image as concise markdown for retrieval.",
    "rules:",
    "1. keep only information that is useful for search, retrieval, or understanding the document.",
    "2. omit unreadable, uncertain, or purely decorative content.",
    "3. do not invent missing text or details.",
    "4. do not wrap the answer in triple backticks.",
    "output policy by image type:",
    "if the image is mainly a table, recreate it as a markdown table.",
    "preserve readable headers, row labels, units, and cell values.",
    "omit cells that are unreadable instead of guessing.",
    "if the image contains formulas or equations:",
    "extract them faithfully.",
    "use latex-style math notation when possible.",
    "keep variable names, subscripts, superscripts, and operators.",
    "if the image is mainly plaintext:",
    "transcribe the readable text in clean markdown.",
    "preserve headings, bullet points, and short paragraph structure when visible.",
    "if the image is an informative figure, chart, diagram, screenshot, or illustration:",
    "first write exactly one concise sentence summarizing what information the image conveys.",
    "then list any clearly readable labels, legends, axis names, key values, or embedded text in markdown bullets.",
    "if the image is decorative, redundant, or not useful for document retrieval:",
    "answer exactly: 불필요한 정보",
    "additional constraints:",
    "prefer faithful extraction over fluent rewriting.",
    "keep the output concise, but do not drop important entities, numbers, labels, or relationships.",
    "preserve technical terms as written in the image.",
    "if only part of the image is readable, extract only that readable part.",
    "do not wrap the answer in triple backticks.",
}


def _unwrap_latex_text_line(line: str) -> str:
    match = re.fullmatch(r"\$\\text\s*\{\s*(.*?)\s*\}\$", line.strip())
    if match:
        return match.group(1).strip()
    return line


def _is_prompt_echo_line(line: str) -> bool:
    normalized = _unwrap_latex_text_line(line).strip()
    if not normalized:
        return False
    normalized = re.sub(r"^[\-\*\u2022]\s*", "", normalized)
    normalized = re.sub(r"\s+", " ", normalized).strip().lower()
    return normalized in _PROMPT_ECHO_PHRASES


def normalize_markdown(text: str) -> str:
    normalized = text.strip()
    fence_match = re.fullmatch(r"```(?:markdown|md)?\s*(.*?)```", normalized, flags=re.DOTALL | re.IGNORECASE)
    if fence_match:
        normalized = fence_match.group(1).strip()
    kept_lines = [line.rstrip() for line in normalized.splitlines() if not _is_prompt_echo_line(line)]
    normalized = "\n".join(line for line in kept_lines).strip()
    return normalized.rstrip() + "\n" if normalized else ""
