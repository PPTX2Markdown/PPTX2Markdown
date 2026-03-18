"""Image pipeline package for bbox-based table classification + Surya extraction."""

from .service import classify_image, extract_table_markdown_from_image

__all__ = ["classify_image", "extract_table_markdown_from_image"]
