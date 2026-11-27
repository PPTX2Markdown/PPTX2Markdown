"""Running LibreOffice headless, for the inputs it can convert and the parser cannot.

Only opt-in features use it (``.ppt`` input, EMF/WMF images): it starts an external
program on the input. Each run gets a throwaway profile and cache inside a
temporary directory that is removed afterwards, and a timeout.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from .errors import UnsupportedPresentationError


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


def require_soffice(purpose: str) -> str:
    soffice = find_soffice()
    if soffice is None:
        raise UnsupportedPresentationError(
            f"{purpose} needs LibreOffice: install it or set SOFFICE_PATH"
        )
    return soffice


def convert_files(
    inputs: list[Path], target: str, output: Path, *, purpose: str, timeout: float
) -> subprocess.CompletedProcess[str]:
    """``soffice --convert-to target`` for every input in one process, into ``output``.

    Raises ``subprocess.TimeoutExpired`` when LibreOffice does not finish in time.
    """

    soffice = require_soffice(purpose)
    with tempfile.TemporaryDirectory(prefix="pptx2markdown-soffice-") as work:
        profile = Path(work) / "profile"
        environment = {**os.environ, "XDG_CACHE_HOME": str(Path(work) / "cache")}
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
            target,
            "--outdir",
            str(output),
            *(str(path) for path in inputs),
        ]
        return subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
            env=environment,
        )
