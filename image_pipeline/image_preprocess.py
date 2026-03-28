"""Image preparation helpers shared across providers and clients."""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator, Optional, Tuple

from .constants import DEFAULT_TRANSPARENT_BG_GRAY


SUPPORTED_IMAGE_SUFFIXES = {
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".bmp",
    ".webp",
    ".tif",
    ".tiff",
}
VECTOR_IMAGE_SUFFIXES = {
    ".emf",
    ".wmf",
}
GEMINI_MIME_BY_SUFFIX = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".webp": "image/webp",
}


def is_supported_image_suffix(suffix: str) -> bool:
    return suffix in SUPPORTED_IMAGE_SUFFIXES or suffix in VECTOR_IMAGE_SUFFIXES


def _resolve_vector_converter() -> Optional[str]:
    env_path = os.getenv("SOFFICE_PATH", "").strip()
    if env_path and Path(env_path).exists():
        return env_path

    if os.name == "nt":
        for candidate in (
            "C:/Program Files/LibreOffice/program/soffice.exe",
            "C:/Program Files (x86)/LibreOffice/program/soffice.exe",
        ):
            if Path(candidate).exists():
                return candidate

    for candidate in ("soffice", "libreoffice"):
        found = shutil.which(candidate)
        if found:
            return found
    soffice_exe = shutil.which("soffice.exe")
    if soffice_exe:
        return soffice_exe
    return None


def _rasterize_vector_image(image_path: Path) -> Path:
    converter = _resolve_vector_converter()
    if not converter:
        raise RuntimeError("LibreOffice is required to rasterize vector images such as EMF/WMF")

    with tempfile.TemporaryDirectory(prefix="vector_raster_") as tmpdir:
        tmp_root = Path(tmpdir)
        staged_input = tmp_root / image_path.name
        shutil.copy2(image_path, staged_input)
        proc = subprocess.run(
            [
                converter,
                "--headless",
                "--convert-to",
                "png",
                "--outdir",
                str(tmp_root),
                str(staged_input),
            ],
            text=True,
            capture_output=True,
            check=False,
        )
        if proc.returncode != 0:
            raise RuntimeError(
                "vector image rasterization failed\n"
                f"stdout={proc.stdout.strip()}\n"
                f"stderr={proc.stderr.strip()}"
            )

        output_path = tmp_root / f"{image_path.stem}.png"
        if not output_path.exists():
            pngs = sorted(tmp_root.glob("*.png"))
            if not pngs:
                raise RuntimeError(f"vector image rasterization produced no PNG output: {image_path.name}")
            output_path = pngs[0]
        persisted_output = Path(tempfile.mkdtemp(prefix="vector_raster_png_")) / output_path.name
        shutil.copy2(output_path, persisted_output)
        return persisted_output


def _flatten_transparent_image(image_path: Path, bg_gray: int = DEFAULT_TRANSPARENT_BG_GRAY) -> Optional[Path]:
    from PIL import Image, ImageOps

    with Image.open(image_path) as loaded:
        image = ImageOps.exif_transpose(loaded)
        try:
            image.seek(0)
        except Exception:
            pass

        has_alpha = image.mode in {"RGBA", "LA"} or "transparency" in image.info
        if not has_alpha:
            return None

        rgba = image.convert("RGBA")
        bg_gray = max(0, min(255, int(bg_gray)))
        background = Image.new("RGBA", rgba.size, (bg_gray, bg_gray, bg_gray, 255))
        composited = Image.alpha_composite(background, rgba).convert("RGB")

        persisted_output = Path(tempfile.mkdtemp(prefix="prepared_image_png_")) / f"{image_path.stem}.png"
        composited.save(persisted_output, format="PNG")
        return persisted_output


@contextmanager
def prepared_image_path(image_path: Path) -> Iterator[Path]:
    suffix = image_path.suffix.lower()
    raster_path: Optional[Path] = None
    flattened_path: Optional[Path] = None

    if suffix in VECTOR_IMAGE_SUFFIXES:
        raster_path = _rasterize_vector_image(image_path)

    source_path = raster_path or image_path
    flattened_path = _flatten_transparent_image(source_path)
    prepared_path_value = flattened_path or source_path

    try:
        yield prepared_path_value
    finally:
        if flattened_path is not None:
            shutil.rmtree(flattened_path.parent, ignore_errors=True)
        if raster_path is not None:
            shutil.rmtree(raster_path.parent, ignore_errors=True)


@contextmanager
def gemini_ready_image_path(image_path: Path) -> Iterator[Tuple[Path, str]]:
    suffix = image_path.suffix.lower()
    mime_type = GEMINI_MIME_BY_SUFFIX.get(suffix)
    if mime_type:
        yield image_path, mime_type
        return

    from PIL import Image, ImageOps

    tmp_root = Path(tempfile.mkdtemp(prefix="gemini_image_png_"))
    output_path = tmp_root / f"{image_path.stem}.png"
    try:
        with Image.open(image_path) as loaded:
            image = ImageOps.exif_transpose(loaded)
            try:
                image.seek(0)
            except Exception:
                pass

            if image.mode in {"RGBA", "LA"} or "transparency" in image.info:
                image.convert("RGBA").save(output_path, format="PNG")
            else:
                image.convert("RGB").save(output_path, format="PNG")
        yield output_path, "image/png"
    finally:
        shutil.rmtree(tmp_root, ignore_errors=True)
