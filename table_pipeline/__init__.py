"""Unified table pipeline: extract -> parse -> render."""

from .render import render_parsed_table_to_markdown

__all__ = [
    "render_parsed_table_to_markdown",
]
