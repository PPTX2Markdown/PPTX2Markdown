"""EMF/WMF pictures drawn as PNG by LibreOffice, when asked to.

Browsers, Markdown viewers and vision models cannot read Windows metafiles, which
older decks are full of. With ``vector_images="png"`` every EMF/WMF image of a
deck is converted in one LibreOffice run and the asset gets the PNG in its place
(same id). LibreOffice draws a metafile on a page; when Pillow is installed the
page margin is trimmed away.
"""

from __future__ import annotations

import io
import subprocess
import tempfile
from pathlib import Path, PurePosixPath
from typing import Literal

from .assets import AssetStore
from .libreoffice import convert_files

VectorImages = Literal["keep", "png"]
VECTOR_MEDIA_TYPES = {"image/emf": ".emf", "image/wmf": ".wmf"}
CONVERSION_TIMEOUT_SECONDS = 300


def _trim(png: bytes) -> bytes:
    """The PNG cropped to its drawing, when Pillow is available to do it."""

    try:
        from PIL import Image, ImageChops
    except ImportError:
        return png
    try:
        with Image.open(io.BytesIO(png)) as image:
            rgb = image.convert("RGB")
            box = ImageChops.difference(rgb, Image.new("RGB", rgb.size, "white")).getbbox()
            if box is None or box == (0, 0, *rgb.size):
                return png
            output = io.BytesIO()
            image.crop(box).save(output, format="PNG", optimize=True)
            return output.getvalue()
    except (OSError, ValueError):
        return png


def render_vector_images(
    assets: AssetStore, *, timeout: float = CONVERSION_TIMEOUT_SECONDS
) -> tuple[list[str], list[str]]:
    """Convert every EMF/WMF image asset to PNG; return converted and failed asset ids."""

    vectors = [
        descriptor
        for descriptor in assets.descriptors()
        if descriptor.kind == "image"
        and descriptor.media_type in VECTOR_MEDIA_TYPES
        and descriptor.payload_available
    ]
    if not vectors:
        return [], []
    with tempfile.TemporaryDirectory(prefix="pptx2markdown-vector-") as work:
        source, output = Path(work) / "in", Path(work) / "out"
        source.mkdir()
        inputs = []
        for descriptor in vectors:
            path = source / f"{descriptor.id}{VECTOR_MEDIA_TYPES[descriptor.media_type]}"
            path.write_bytes(assets.read_bytes(descriptor.id))
            inputs.append(path)
        try:
            convert_files(
                inputs, "png", output, purpose="converting EMF/WMF images", timeout=timeout
            )
        except subprocess.TimeoutExpired:
            return [], [descriptor.id for descriptor in vectors]
        converted, failed = [], []
        for descriptor in vectors:
            png = output / f"{descriptor.id}.png"
            if not png.is_file():
                failed.append(descriptor.id)
                continue
            name = PurePosixPath(descriptor.original_name).stem + ".png"
            assets.replace(
                descriptor.id, _trim(png.read_bytes()), original_name=name, media_type="image/png"
            )
            converted.append(descriptor.id)
    return converted, failed
