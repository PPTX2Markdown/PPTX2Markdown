from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Literal, Optional

PptConverterMode = Literal["auto", "powerpoint", "libreoffice"]
_WINDOWS_RESERVED_FILENAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


@dataclass(frozen=True)
class PptConversionResult:
    pptx_path: Path
    converter: str
    reused_cache: bool = False


class PptConversionError(RuntimeError):
    pass


def _safe_stem(path: Path) -> str:
    stem = path.stem.strip() or "presentation"
    stem = "".join("_" if ch in '<>:"/\\|?*' or ord(ch) < 32 else ch for ch in stem)
    stem = re.sub(r"\s+", " ", stem).strip(" .")
    stem = re.sub(r"_+", "_", stem)
    if not stem:
        return "presentation"
    if stem.upper() in _WINDOWS_RESERVED_FILENAMES:
        return f"{stem}_"
    return stem


def _cache_key(ppt_path: Path) -> str:
    return hashlib.sha256(str(ppt_path.resolve()).encode("utf-8")).hexdigest()[:12]


def _marker_matches(marker_path: Path, ppt_path: Path, mode: PptConverterMode) -> bool:
    try:
        marker = json.loads(marker_path.read_text(encoding="utf-8"))
    except Exception:
        return False
    if not isinstance(marker, dict):
        return False
    try:
        stat = ppt_path.stat()
    except OSError:
        return False
    return (
        marker.get("source_path") == str(ppt_path.resolve())
        and marker.get("size") == stat.st_size
        and marker.get("mtime_ns") == stat.st_mtime_ns
        and marker.get("mode") == mode
    )


def _write_marker(
    marker_path: Path, ppt_path: Path, mode: PptConverterMode, converter: str
) -> None:
    stat = ppt_path.stat()
    marker_path.write_text(
        json.dumps(
            {
                "source_path": str(ppt_path.resolve()),
                "size": stat.st_size,
                "mtime_ns": stat.st_mtime_ns,
                "mode": mode,
                "converter": converter,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


def _resolve_soffice_cmd() -> Optional[str]:
    candidates: list[Path] = []
    env_path = os.getenv("SOFFICE_PATH", "").strip()
    if env_path:
        candidates.append(Path(env_path))
    if sys.platform == "darwin":
        candidates.extend(
            [
                Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
                Path.home()
                / "Applications"
                / "LibreOffice.app"
                / "Contents"
                / "MacOS"
                / "soffice",
            ]
        )
    elif sys.platform.startswith("win"):
        candidates.extend(
            [
                Path("C:/Program Files/LibreOffice/program/soffice.exe"),
                Path("C:/Program Files (x86)/LibreOffice/program/soffice.exe"),
            ]
        )
    for cmd_name in ("soffice", "libreoffice", "soffice.exe"):
        found = shutil.which(cmd_name)
        if found:
            candidates.append(Path(found))
    for candidate in candidates:
        if candidate.exists() and candidate.is_file():
            return str(candidate)
    return None


def _convert_with_libreoffice(ppt_path: Path, output_path: Path) -> None:
    soffice_cmd = _resolve_soffice_cmd()
    if not soffice_cmd:
        raise PptConversionError(
            "LibreOffice converter not found. Install LibreOffice or set SOFFICE_PATH."
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    profile_dir = Path(tempfile.mkdtemp(prefix="libreoffice-profile-"))
    temp_out_dir = Path(tempfile.mkdtemp(prefix="ppt-to-pptx-"))
    cmd = [
        soffice_cmd,
        f"-env:UserInstallation={profile_dir.resolve().as_uri()}",
        "--headless",
        "--invisible",
        "--nodefault",
        "--nologo",
        "--nolockcheck",
        "--norestore",
        "--convert-to",
        "pptx",
        str(ppt_path),
        "--outdir",
        str(temp_out_dir),
    ]
    env = os.environ.copy()
    if sys.platform == "darwin":
        env["HOME"] = "/tmp"
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, env=env)
        converted = temp_out_dir / f"{ppt_path.stem}.pptx"
        if proc.returncode != 0 or not converted.exists():
            details = "\n".join(x.strip() for x in (proc.stdout, proc.stderr) if x.strip())
            if details:
                raise PptConversionError(f"LibreOffice conversion failed: {details}")
            raise PptConversionError("LibreOffice conversion failed without diagnostic output.")
        shutil.move(str(converted), output_path)
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)
        shutil.rmtree(temp_out_dir, ignore_errors=True)


def _convert_with_powerpoint(ppt_path: Path, output_path: Path) -> None:
    if not sys.platform.startswith("win"):
        raise PptConversionError("PowerPoint conversion is only available on Windows.")
    try:
        import pythoncom  # type: ignore[import-not-found]
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise PptConversionError("PowerPoint conversion requires pywin32 on Windows.") from exc

    output_path.parent.mkdir(parents=True, exist_ok=True)
    app = None
    presentation = None
    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        app.DisplayAlerts = 0
        presentation = app.Presentations.Open(str(ppt_path.resolve()), True, False, False)
        presentation.SaveAs(str(output_path.resolve()), 24)
    except Exception as exc:
        raise PptConversionError(f"PowerPoint conversion failed: {exc}") from exc
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def convert_ppt_to_pptx(
    ppt_path: Path, cache_root: Path, mode: PptConverterMode = "auto"
) -> PptConversionResult:
    if mode not in {"auto", "powerpoint", "libreoffice"}:
        raise ValueError(f"unsupported ppt converter mode: {mode}")
    ppt_path = ppt_path.resolve()
    if ppt_path.suffix.lower() != ".ppt":
        raise ValueError(f"expected .ppt input: {ppt_path}")
    if not ppt_path.exists() or not ppt_path.is_file():
        raise FileNotFoundError(f"ppt not found: {ppt_path}")

    cache_dir = cache_root / f"{_safe_stem(ppt_path)}-{_cache_key(ppt_path)}"
    output_path = cache_dir / f"{_safe_stem(ppt_path)}.pptx"
    marker_path = cache_dir / ".ppt_to_pptx_source.json"
    if output_path.exists() and _marker_matches(marker_path, ppt_path, mode):
        try:
            if output_path.stat().st_size > 0:
                converter = json.loads(marker_path.read_text(encoding="utf-8")).get(
                    "converter", "cache"
                )
                return PptConversionResult(
                    output_path.resolve(), str(converter), reused_cache=True
                )
        except OSError:
            pass

    cache_dir.mkdir(parents=True, exist_ok=True)
    if output_path.exists():
        output_path.unlink()

    errors: list[str] = []
    if mode == "auto":
        converters = (
            ("powerpoint", "libreoffice") if sys.platform.startswith("win") else ("libreoffice",)
        )
    else:
        converters = (mode,)
    for converter in converters:
        try:
            if converter == "powerpoint":
                _convert_with_powerpoint(ppt_path, output_path)
            else:
                _convert_with_libreoffice(ppt_path, output_path)
            if not output_path.exists() or output_path.stat().st_size == 0:
                raise PptConversionError(f"{converter} produced no pptx output.")
            _write_marker(marker_path, ppt_path, mode, converter)
            return PptConversionResult(output_path.resolve(), converter, reused_cache=False)
        except Exception as exc:
            errors.append(f"{converter}: {exc}")
            if output_path.exists():
                output_path.unlink()

    raise PptConversionError("PPT to PPTX conversion failed. " + " | ".join(errors))
