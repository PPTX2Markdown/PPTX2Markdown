"""Legacy binary .ppt input, converted to .pptx by LibreOffice when asked to.

Nothing here runs unless the caller opts in with ``ppt_converter="libreoffice"``.
The converted package only passes through a temporary directory that is removed
afterwards; the parser then reads its bytes with the usual limits.
"""

from __future__ import annotations

import subprocess
import tempfile
from pathlib import Path
from typing import Literal

from .errors import PresentationParseError
from .libreoffice import convert_files, find_soffice

PptConverter = Literal["none", "libreoffice"]
CONVERSION_TIMEOUT_SECONDS = 300

__all__ = ["CONVERSION_TIMEOUT_SECONDS", "PptConverter", "find_soffice", "ppt_to_pptx"]


def ppt_to_pptx(path: Path, *, timeout: float = CONVERSION_TIMEOUT_SECONDS) -> bytes:
    """The bytes of ``path`` (a .ppt file) saved as .pptx by LibreOffice."""

    with tempfile.TemporaryDirectory(prefix="pptx2markdown-ppt-") as work:
        output = Path(work)
        try:
            completed = convert_files(
                [path], "pptx", output, purpose="converting .ppt", timeout=timeout
            )
        except subprocess.TimeoutExpired as exc:
            raise PresentationParseError(
                path, f"LibreOffice did not finish within {timeout:g} seconds"
            ) from exc
        converted = output / f"{path.stem}.pptx"
        if completed.returncode != 0 or not converted.is_file():
            details = (completed.stderr or completed.stdout or "").strip()
            raise PresentationParseError(
                path, f"LibreOffice could not convert the file{': ' + details if details else ''}"
            )
        return converted.read_bytes()
