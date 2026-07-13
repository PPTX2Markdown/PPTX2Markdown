"""Canonical filesystem layout for pptx2markdown.

Final documents live under ``output_dir``. Every generated intermediate file
lives under ``work_dir`` so individual pipelines do not create their own
``output`` or ``target_slides`` directories relative to their launch location.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Union

PathLike = Union[str, Path]

DEFAULT_WORK_DIR_NAME = ".pptx2markdown"
DEFAULT_OUTPUT_DIR_NAME = "output"


def _resolved(path: PathLike) -> Path:
    return Path(path).expanduser().resolve()


@dataclass(frozen=True)
class WorkspacePaths:
    """All directories used by a conversion workspace.

    Constructing this object never creates directories. Callers create only the
    directory they are about to write, which avoids scattering empty folders
    during input lookup.
    """

    project_dir: Path
    work_dir: Path
    output_dir: Path

    @classmethod
    def from_base(
        cls,
        base_dir: Optional[PathLike] = None,
        *,
        work_dir: Optional[PathLike] = None,
        output_dir: Optional[PathLike] = None,
    ) -> "WorkspacePaths":
        project_dir = _resolved(base_dir or Path.cwd())
        resolved_work_dir = _resolved(work_dir or project_dir / DEFAULT_WORK_DIR_NAME)
        resolved_output_dir = _resolved(output_dir or project_dir / DEFAULT_OUTPUT_DIR_NAME)
        return cls(
            project_dir=project_dir,
            work_dir=resolved_work_dir,
            output_dir=resolved_output_dir,
        )

    @property
    def target_slides(self) -> Path:
        return self.work_dir / "target_slides"

    @property
    def target_pptx(self) -> Path:
        return self.work_dir / "target_pptx"

    @property
    def ppt_conversion_cache(self) -> Path:
        return self.work_dir / ".cache" / "ppt_to_pptx"

    @property
    def structure_analysis(self) -> Path:
        return self.work_dir / "structure_analysis"

    @property
    def table_pipeline(self) -> Path:
        return self.work_dir / "table_pipeline"

    @property
    def table_extract_results(self) -> Path:
        return self.table_pipeline / "extract_results"

    @property
    def table_parsing_results(self) -> Path:
        return self.table_pipeline / "parsing_results"

    @property
    def table_markdown_results(self) -> Path:
        return self.table_pipeline / "tables"


def ensure_directory(path: Path, *, label: str = "directory") -> Path:
    """Create and return ``path``, rejecting an existing non-directory."""
    if path.exists() and not path.is_dir():
        raise NotADirectoryError(f"{label} path exists but is not a directory: {path}")
    path.mkdir(parents=True, exist_ok=True)
    return path
