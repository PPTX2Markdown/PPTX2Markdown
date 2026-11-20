"""Public Python API for the standalone rewrite."""

from __future__ import annotations

import hashlib
from concurrent.futures import ProcessPoolExecutor
from concurrent.futures.process import BrokenProcessPool
from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO, Iterable, Literal

from .assets import safe_basename
from .errors import PPTX2MarkdownError, UnsupportedPresentationError
from .parser import ParseOptions, parse_components
from .ppt import PptConverter, ppt_to_pptx
from .presentation import ParsedPresentation
from .recipe import Recipe
from .tables import MergedCells
from .vector import VectorImages
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
    path: str | Path | bytes | BinaryIO,
    *,
    name: str | None = None,
    include_assets: bool = True,
    include_notes: bool = True,
    include_hidden: bool = True,
    strict: bool = True,
    merged_cells: MergedCells = "fill",
    recipe: Recipe | None = None,
    ppt_converter: PptConverter = "none",
    vector_images: VectorImages = "keep",
) -> ParsedPresentation:
    """Parse one PPTX without creating persistent output.

    ``path`` may be a local path, the file's bytes or a binary stream (for an
    upload, say); ``name`` is the file name recorded for in-memory input.

    ``merged_cells`` controls table Markdown: ``fill`` repeats a merged value in
    every cell it covers, ``header`` only in the header row, ``blank`` nowhere.
    ``recipe`` is applied to the parsed document before it is returned.
    ``ppt_converter="libreoffice"`` accepts a legacy ``.ppt`` path by having
    LibreOffice save it as ``.pptx`` first (see ``pptx2markdown.ppt``), and
    ``vector_images="png"`` has it draw EMF/WMF images as PNG (``pptx2markdown.vector``).
    """

    if isinstance(path, (str, Path)) and Path(path).suffix.casefold() == ".ppt":
        source = Path(path).expanduser()
        if ppt_converter != "libreoffice":
            raise UnsupportedPresentationError(
                f"{source.name} is a legacy .ppt file: save it as .pptx, or pass "
                "ppt_converter='libreoffice' (--ppt-converter libreoffice) to convert it"
            )
        if not source.is_file():
            raise FileNotFoundError(source)
        path, name = ppt_to_pptx(source), name or f"{source.stem}.pptx"
    components = parse_components(
        path,
        name=name,
        options=ParseOptions(
            include_assets=include_assets,
            include_notes=include_notes,
            include_hidden=include_hidden,
            strict=strict,
            merged_cells=merged_cells,
            vector_images=vector_images,
        ),
    )
    presentation = ParsedPresentation(
        document=components.document,
        assets=components.assets,
        source_path=components.source_path,
    )
    if recipe is not None:
        presentation.apply(recipe)
    return presentation


def convert(
    path: str | Path,
    *,
    output_dir: str | Path = "output",
    bundle_name: str | None = None,
    output_format: Literal["markdown", "json", "chunks"] = "markdown",
    assets: Literal["copy", "omit"] = "copy",
    overwrite: Literal["error", "replace"] = "error",
    strict: bool = True,
    merged_cells: MergedCells = "fill",
    recipe: Recipe | None = None,
    max_chunk_chars: int | None = None,
    ppt_converter: PptConverter = "none",
    vector_images: VectorImages = "keep",
) -> SaveResult:
    """Parse and save one presentation as a portable bundle."""

    source = Path(path).expanduser()
    resolved_bundle_name = safe_basename(
        bundle_name or source.stem,
        fallback="presentation",
    )
    bundle_dir = Path(output_dir).expanduser() / resolved_bundle_name
    with parse(
        source,
        include_assets=assets == "copy",
        strict=strict,
        merged_cells=merged_cells,
        recipe=recipe,
        ppt_converter=ppt_converter,
        vector_images=vector_images,
    ) as presentation:
        return presentation.save(
            bundle_dir,
            output_format=output_format,
            assets=assets,
            overwrite=overwrite,
            max_chunk_chars=max_chunk_chars,
        )


def convert_many(
    paths: Iterable[str | Path],
    *,
    output_dir: str | Path = "output",
    output_format: Literal["markdown", "json", "chunks"] = "markdown",
    assets: Literal["copy", "omit"] = "copy",
    overwrite: Literal["error", "replace"] = "error",
    strict: bool = True,
    merged_cells: MergedCells = "fill",
    recipe: Recipe | None = None,
    max_chunk_chars: int | None = None,
    ppt_converter: PptConverter = "none",
    vector_images: VectorImages = "keep",
    jobs: int = 1,
) -> BatchConversionResult:
    """Convert inputs independently and preserve successful bundles.

    ``jobs`` greater than 1 converts that many decks at once in worker processes;
    bundle names and the order of results are the same as with one job.
    """

    if jobs < 1:
        raise ValueError("jobs must be at least 1")

    sources = [Path(raw_path).expanduser() for raw_path in paths]
    stem_counts: dict[str, int] = {}
    for source in sources:
        stem = safe_basename(source.stem, fallback="presentation")
        key = stem.casefold()
        stem_counts[key] = stem_counts.get(key, 0) + 1

    options = {
        "output_dir": output_dir,
        "output_format": output_format,
        "assets": assets,
        "overwrite": overwrite,
        "strict": strict,
        "merged_cells": merged_cells,
        "recipe": recipe,
        "max_chunk_chars": max_chunk_chars,
        "ppt_converter": ppt_converter,
        "vector_images": vector_images,
    }
    # Bundle names are settled up front, so they never depend on which job ends first.
    outcomes: list[SaveResult | ConversionFailure | tuple[Path, str]] = []
    used_bundle_names: dict[str, int] = {}
    for source in sources:
        try:
            stem = safe_basename(source.stem, fallback="presentation")
            if stem_counts[stem.casefold()] > 1:
                stem = f"{stem}-{_source_hash_prefix(source)}"
        except OSError as exc:
            outcomes.append(ConversionFailure(source=source, error=exc))
            continue
        occurrence = used_bundle_names.get(stem.casefold(), 0) + 1
        used_bundle_names[stem.casefold()] = occurrence
        outcomes.append((source, stem if occurrence == 1 else f"{stem}-{occurrence}"))

    def run(source: Path, bundle_name: str) -> SaveResult:
        return convert(source, bundle_name=bundle_name, **options)  # type: ignore[arg-type]

    pending = [(index, item) for index, item in enumerate(outcomes) if isinstance(item, tuple)]
    if jobs == 1 or len(pending) <= 1:
        for index, (source, bundle_name) in pending:
            try:
                outcomes[index] = run(source, bundle_name)
            except (PPTX2MarkdownError, OSError, ValueError) as exc:
                outcomes[index] = ConversionFailure(source=source, error=exc)
    else:
        with ProcessPoolExecutor(max_workers=min(jobs, len(pending))) as pool:
            futures = {
                index: pool.submit(convert, source, bundle_name=bundle_name, **options)
                for index, (source, bundle_name) in pending
            }
            for index, (source, _) in pending:
                try:
                    outcomes[index] = futures[index].result()
                except (PPTX2MarkdownError, OSError, ValueError, BrokenProcessPool) as exc:
                    outcomes[index] = ConversionFailure(source=source, error=exc)
    succeeded = tuple(item for item in outcomes if isinstance(item, SaveResult))
    failed = tuple(item for item in outcomes if isinstance(item, ConversionFailure))
    return BatchConversionResult(succeeded, failed)
