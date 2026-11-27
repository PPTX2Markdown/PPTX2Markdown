"""Single slides rendered by LibreOffice, for judges that need to see the layout.

Review items whose ``look`` is ``"slide"`` cannot be settled from text. These
helpers render just the slides asked for, locally: ``slide_pdf`` exports one page
through LibreOffice's PDF export with a page range, and ``slide_png`` draws that
page as a PNG (LibreOffice's PNG export ignores page ranges, so it goes through
the one-page PDF). Both return None when LibreOffice is missing or times out.
"""

from __future__ import annotations

import subprocess
import tempfile
from pathlib import Path

from .libreoffice import convert_files, find_soffice

RENDER_TIMEOUT_SECONDS = 120


def slide_pdf(
    source: str | Path, page: int, *, timeout: float = RENDER_TIMEOUT_SECONDS
) -> bytes | None:
    """Slide ``page`` of ``source`` as a one-page PDF, or None without LibreOffice."""

    if find_soffice() is None:
        return None
    source = Path(source)
    target = 'pdf:impress_pdf_Export:{"PageRange":{"type":"string","value":"%d"}}' % page
    with tempfile.TemporaryDirectory(prefix="pptx2markdown-slide-") as work:
        try:
            convert_files(
                [source], target, Path(work), purpose="rendering a slide", timeout=timeout
            )
        except subprocess.TimeoutExpired:
            return None
        pdf = Path(work) / f"{source.stem}.pdf"
        return pdf.read_bytes() if pdf.is_file() else None


def slide_png(
    source: str | Path, page: int, *, timeout: float = RENDER_TIMEOUT_SECONDS
) -> bytes | None:
    """Slide ``page`` of ``source`` as a PNG (720 pixels high), or None."""

    pdf = slide_pdf(source, page, timeout=timeout)
    if pdf is None:
        return None
    with tempfile.TemporaryDirectory(prefix="pptx2markdown-slide-") as work:
        page_pdf = Path(work) / "slide.pdf"
        page_pdf.write_bytes(pdf)
        try:
            convert_files(
                [page_pdf],
                "png:draw_png_Export",
                Path(work) / "png",
                purpose="rendering a slide",
                timeout=timeout,
            )
        except subprocess.TimeoutExpired:
            return None
        png = Path(work) / "png" / "slide.png"
        return png.read_bytes() if png.is_file() else None
