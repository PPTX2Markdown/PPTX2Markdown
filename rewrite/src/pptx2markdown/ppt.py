"""Legacy binary .ppt input, converted to .pptx by LibreOffice when asked to.

Nothing here runs unless the caller opts in with ``ppt_converter="libreoffice"``:
it starts an external program on the input. LibreOffice runs headless with a
throwaway profile, and the converted package only passes through temporary
directories that are removed afterwards; the parser then reads its bytes with
the usual limits.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Literal

from .errors import PresentationParseError, UnsupportedPresentationError

PptConverter = Literal["none", "libreoffice"]
CONVERSION_TIMEOUT_SECONDS = 300


def find_soffice() -> str | None:
    """LibreOffice's ``soffice``: $SOFFICE_PATH, the usual install places, then PATH."""

    candidates = [os.environ.get("SOFFICE_PATH", "").strip()]
    if sys.platform == "darwin":
        bundle = Path("LibreOffice.app/Contents/MacOS/soffice")
        candidates += [
            str(Path("/Applications") / bundle),
            str(Path.home() / "Applications" / bundle),
        ]
    elif sys.platform.startswith("win"):
        candidates += [
            "C:/Program Files/LibreOffice/program/soffice.exe",
            "C:/Program Files (x86)/LibreOffice/program/soffice.exe",
        ]
    candidates += [shutil.which(name) or "" for name in ("soffice", "libreoffice")]
    return next((c for c in candidates if c and Path(c).is_file()), None)


def ppt_to_pptx(path: Path, *, timeout: float = CONVERSION_TIMEOUT_SECONDS) -> bytes:
    """The bytes of ``path`` (a .ppt file) saved as .pptx by LibreOffice."""

    soffice = find_soffice()
    if soffice is None:
        raise UnsupportedPresentationError(
            "converting .ppt needs LibreOffice: install it or set SOFFICE_PATH"
        )
    with tempfile.TemporaryDirectory(prefix="pptx2markdown-ppt-") as work:
        profile = Path(work) / "profile"
        output = Path(work) / "out"
        command = [
            soffice,
            f"-env:UserInstallation={profile.as_uri()}",
            "--headless",
            "--invisible",
            "--nodefault",
            "--nolockcheck",
            "--nologo",
            "--norestore",
            "--convert-to",
            "pptx",
            "--outdir",
            str(output),
            str(path),
        ]
        try:
            completed = subprocess.run(
                command, capture_output=True, text=True, timeout=timeout, check=False
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
