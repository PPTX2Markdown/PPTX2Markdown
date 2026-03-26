from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional

from pydantic import BaseModel, ConfigDict, Field


def utc_now_z() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


class SlideStats(BaseModel):
    model_config = ConfigDict(extra="forbid")

    blocks_total: int = 0
    text_blocks: int = 0
    image_blocks: int = 0
    table_blocks: int = 0
    table_skipped_blocks: int = 0
    unsupported_blocks: int = 0
    skipped_blocks: int = 0
    resolved_images: int = 0
    unresolved_images: int = 0
    warnings: List[str] = Field(default_factory=list)
    rels_path: Optional[str] = None

    def to_slide_row_fields(self) -> Dict[str, object]:
        return {
            "blocks_total": self.blocks_total,
            "text_blocks": self.text_blocks,
            "image_blocks": self.image_blocks,
            "table_blocks": self.table_blocks,
            "table_skipped_blocks": self.table_skipped_blocks,
            "unsupported_blocks": self.unsupported_blocks,
            "skipped_blocks": self.skipped_blocks,
            "rels_path": self.rels_path,
            "warnings": list(self.warnings),
        }


class ManifestSummary(BaseModel):
    model_config = ConfigDict(extra="forbid")

    processed_packages: int = 0
    processed_slides: int = 0
    failed: int = 0
    resolved_images: int = 0
    unresolved_images: int = 0
    table_blocks: int = 0
    table_skipped_blocks: int = 0

    def add_slide(self, stats: SlideStats) -> None:
        self.processed_slides += 1
        self.resolved_images += stats.resolved_images
        self.unresolved_images += stats.unresolved_images
        self.table_blocks += stats.table_blocks
        self.table_skipped_blocks += stats.table_skipped_blocks


class ConversionManifest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    started_at: str = Field(default_factory=utc_now_z)
    packages: List[Dict[str, object]] = Field(default_factory=list)
    summary: ManifestSummary = Field(default_factory=ManifestSummary)
    finished_at: Optional[str] = None

    def mark_finished(self) -> None:
        self.finished_at = utc_now_z()


class ConverterConfig(BaseModel):
    model_config = ConfigDict(arbitrary_types_allowed=True)

    cwd: Path
    repo_root: Path
    output_dir: Path
    debug_output_dir: Path
    inputs: List[str] = Field(default_factory=list)
    reading_order: str = "xml"
    strict: bool = False
    reuse_surya_cache: bool = False
    image_table_pipeline: bool = False
