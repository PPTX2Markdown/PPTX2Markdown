"""Public exception hierarchy for the rewritten converter."""

from __future__ import annotations

from pathlib import Path


class PPTX2MarkdownError(Exception):
    """Base class for all library-owned failures."""


class UnsupportedPresentationError(PPTX2MarkdownError, ValueError):
    """Raised when a source is not a supported local PPTX file."""


class PackageSecurityError(PPTX2MarkdownError):
    """Raised when an OOXML package violates a configured safety limit."""


class PresentationParseError(PPTX2MarkdownError):
    """Raised when a presentation cannot be parsed."""

    def __init__(
        self,
        path: str | Path,
        message: str,
        *,
        slide_number: int | None = None,
    ) -> None:
        self.path = Path(path)
        self.slide_number = slide_number
        self.detail = message
        location = f"slide {slide_number}" if slide_number is not None else "presentation"
        super().__init__(f"{self.path}: {location} parse failed: {message}")

    def __reduce__(self) -> tuple[object, ...]:
        # Rebuilt from its fields, so it crosses process boundaries (convert_many jobs).
        return (_rebuild_parse_error, (type(self), self.path, self.detail, self.slide_number))


def _rebuild_parse_error(
    cls: type[PresentationParseError], path: Path, detail: str, slide_number: int | None
) -> PresentationParseError:
    return cls(path, detail, slide_number=slide_number)


class AssetError(PPTX2MarkdownError):
    """Base class for asset registration and access failures."""


class AssetLimitError(AssetError):
    """Raised when an asset store limit would be exceeded."""


class AssetUnavailableError(AssetError):
    """Raised when an asset payload was intentionally omitted or already closed."""


class AssetResolutionError(PPTX2MarkdownError):
    """Raised when rendered output cannot resolve an asset reference."""


class SaveConflictError(PPTX2MarkdownError):
    """Raised when a bundle destination already exists."""


class BundleWriteError(PPTX2MarkdownError):
    """Raised when a portable output bundle cannot be completed."""
