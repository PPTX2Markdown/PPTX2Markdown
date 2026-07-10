from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional

from pydantic import BaseModel, ConfigDict, Field


def utc_now_z() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


@dataclass(frozen=True)
class ParagraphSegment:
    kind: str
    text: str


@dataclass(frozen=True)
class ShapeBlock:
    kind: str
    segments: List[ParagraphSegment] = field(default_factory=list)
    level: Optional[int] = None

    @property
    def plain_text(self) -> str:
        parts = [
            segment.text
            for segment in self.segments
            if segment.kind != "break" and segment.text.strip()
        ]
        return " ".join(parts).strip()

    @property
    def markdown_text(self) -> str:
        rendered: List[str] = []
        for segment in self.segments:
            if segment.kind == "break":
                if rendered and not rendered[-1].endswith("\n"):
                    rendered.append("\n")
                continue
            text = segment.text.strip()
            if not text:
                continue
            if rendered and not rendered[-1].endswith("\n"):
                rendered.append(" ")
            rendered.append(text)
        return "".join(rendered).strip()

    @property
    def has_math(self) -> bool:
        return any(segment.kind in {"math_inline", "math_block"} for segment in self.segments)

    @property
    def is_math_only(self) -> bool:
        non_empty = [segment for segment in self.segments if segment.text.strip()]
        return bool(non_empty) and all(
            segment.kind in {"math_inline", "math_block"} for segment in non_empty
        )


class SlideStats(BaseModel):
    model_config = ConfigDict(extra="forbid")

    blocks_total: int = 0
    text_blocks: int = 0
    math_blocks: int = 0
    inline_math_segments: int = 0
    block_math_segments: int = 0
    math_conversion_failures: int = 0
    image_blocks: int = 0
    chart_blocks: int = 0
    smartart_blocks: int = 0
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
            "math_blocks": self.math_blocks,
            "inline_math_segments": self.inline_math_segments,
            "block_math_segments": self.block_math_segments,
            "math_conversion_failures": self.math_conversion_failures,
            "image_blocks": self.image_blocks,
            "chart_blocks": self.chart_blocks,
            "smartart_blocks": self.smartart_blocks,
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
    math_blocks: int = 0
    inline_math_segments: int = 0
    block_math_segments: int = 0
    math_conversion_failures: int = 0
    resolved_images: int = 0
    unresolved_images: int = 0
    chart_blocks: int = 0
    smartart_blocks: int = 0
    table_blocks: int = 0
    table_skipped_blocks: int = 0

    def add_slide(self, stats: SlideStats) -> None:
        self.processed_slides += 1
        self.math_blocks += stats.math_blocks
        self.inline_math_segments += stats.inline_math_segments
        self.block_math_segments += stats.block_math_segments
        self.math_conversion_failures += stats.math_conversion_failures
        self.resolved_images += stats.resolved_images
        self.unresolved_images += stats.unresolved_images
        self.chart_blocks += stats.chart_blocks
        self.smartart_blocks += stats.smartart_blocks
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


class PreparedPackage(BaseModel):
    model_config = ConfigDict(arbitrary_types_allowed=True, extra="forbid")

    package_dir: Path
    source_pptx_path: Path

    @property
    def source_stem(self) -> str:
        stem = self.source_pptx_path.stem.strip()
        return stem or self.package_dir.name

    @property
    def package_dir_name(self) -> str:
        return self.source_stem

    @property
    def name(self) -> str:
        return self.package_dir.name

    @property
    def output_markdown_name(self) -> str:
        return f"{self.source_stem}.md"

    def output_markdown_path(self, output_dir: Path) -> Path:
        return output_dir / self.name / self.output_markdown_name

    def with_package_dir(self, package_dir: Path) -> "PreparedPackage":
        return self.model_copy(update={"package_dir": package_dir})


# 전체 변환 파이프라인의 정규화된 실행 설정이다.
class ConverterConfig(BaseModel):
    model_config = ConfigDict(arbitrary_types_allowed=True, extra="forbid")

    cwd: Path
    output_dir: Path
    inputs: List[str] = Field(default_factory=list)
    reading_order: str = "xml"
    heading_mode: str = "auto"
    strict: bool = False
    pptx_inheritance: str = "style"
    inherited_shapes: str = "visible"
    reuse_surya_cache: bool = False
    image_vlm_provider: str = "local"
    image_vlm_model: Optional[str] = None
    image_vlm_prompt: str = ""
    image_vlm_max_new_tokens: int = 1024
    image_vlm_api_key_env: str = "GEMINI_API_KEY"
    ignore_image_vlm_cache: bool = False
    ppt_converter: str = "auto"
