"""Public Python API for the standalone rewrite."""

from __future__ import annotations

import hashlib
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Literal

from .assets import safe_basename
from .errors import PPTX2MarkdownError
from .parser import ParseOptions, parse_components
from .presentation import ParsedPresentation
from .tables import MergedCells
from .writer import SaveResult


@dataclass(frozen=True)
class ConversionFailure:
    source: Path
    error: PPTX2MarkdownError | OSError | ValueError


@dataclass(frozen=True)
class BatchConversionResult:
    succeeded: tuple[SaveResult, ...]
    failed: tuple[ConversionFailure, ...]

    @property
    def exit_code(self) -> int:
        return 1 if self.failed else 0


def _source_hash_prefix(source: Path) -> str:
    digest = hashlib.sha256()
    with source.open("rb") as stream:
        while chunk := stream.read(1024 * 1024):
            digest.update(chunk)
    return digest.hexdigest()[:8]


def parse(
    path: str | Path,
    *,
    include_assets: bool = True,
    include_notes: bool = True,
    include_hidden: bool = True,
    strict: bool = True,
    merged_cells: MergedCells = "fill",
) -> ParsedPresentation:
    """Parse one local PPTX without creating persistent output.

    ``merged_cells`` controls table Markdown: ``fill`` repeats a merged value in
    every cell it covers, ``header`` only in the header row, ``blank`` nowhere.
    """

    components = parse_components(
        path,
        options=ParseOptions(
            include_assets=include_assets,
            include_notes=include_notes,
            include_hidden=include_hidden,
            strict=strict,
            merged_cells=merged_cells,
        ),
    )
    return ParsedPresentation(
        document=components.document,
        assets=components.assets,
        source_path=components.source_path,
    )


def convert(
    path: str | Path,
    *,
    output_dir: str | Path = "output",
    bundle_name: str | None = None,
    output_format: Literal["markdown", "json"] = "markdown",
    assets: Literal["copy", "omit"] = "copy",
    overwrite: Literal["error", "replace"] = "error",
    strict: bool = True,
    merged_cells: MergedCells = "fill",
) -> SaveResult:
    """Parse and save one presentation as a portable bundle."""

    source = Path(path).expanduser()
    resolved_bundle_name = safe_basename(
        bundle_name or source.stem,
        fallback="presentation",
    )
    bundle_dir = Path(output_dir).expanduser() / resolved_bundle_name
    with parse(
        source, include_assets=assets == "copy", strict=strict, merged_cells=merged_cells
    ) as presentation:
        return presentation.save(
            bundle_dir,
            output_format=output_format,
            assets=assets,
            overwrite=overwrite,
        )


def convert_many(
    paths: Iterable[str | Path],
    *,
    output_dir: str | Path = "output",
    output_format: Literal["markdown", "json"] = "markdown",
    assets: Literal["copy", "omit"] = "copy",
    overwrite: Literal["error", "replace"] = "error",
    strict: bool = True,
    merged_cells: MergedCells = "fill",
) -> BatchConversionResult:
    """Convert inputs independently and preserve successful bundles."""

    sources = [Path(raw_path).expanduser() for raw_path in paths]
    stem_counts: dict[str, int] = {}
    for source in sources:
        stem = safe_basename(source.stem, fallback="presentation")
        key = stem.casefold()
        stem_counts[key] = stem_counts.get(key, 0) + 1

    succeeded: list[SaveResult] = []
    failed: list[ConversionFailure] = []
    used_bundle_names: dict[str, int] = {}
    for source in sources:
        try:
            stem = safe_basename(source.stem, fallback="presentation")
            if stem_counts[stem.casefold()] > 1:
                stem = f"{stem}-{_source_hash_prefix(source)}"
            occurrence = used_bundle_names.get(stem.casefold(), 0) + 1
            used_bundle_names[stem.casefold()] = occurrence
            bundle_name = stem if occurrence == 1 else f"{stem}-{occurrence}"
            succeeded.append(
                convert(
                    source,
                    output_dir=output_dir,
                    bundle_name=bundle_name,
                    output_format=output_format,
                    assets=assets,
                    overwrite=overwrite,
                    strict=strict,
                    merged_cells=merged_cells,
                )
            )
        except (PPTX2MarkdownError, OSError, ValueError) as exc:
            failed.append(ConversionFailure(source=source, error=exc))
    return BatchConversionResult(tuple(succeeded), tuple(failed))
