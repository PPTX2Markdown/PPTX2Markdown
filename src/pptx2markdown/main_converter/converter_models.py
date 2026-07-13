from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Literal, Optional

from pydantic import BaseModel, ConfigDict, Field, model_validator


def utc_now_z() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


@dataclass(frozen=True)
class ParagraphSegment:
    kind: str
    text: str
    target: Optional[str] = None
    font_pt: Optional[float] = None


@dataclass(frozen=True)
class ShapeBlock:
    kind: str
    segments: List[ParagraphSegment] = field(default_factory=list)
    level: Optional[int] = None
    list_explicit_none: bool = False

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
                if rendered:
                    rendered.append("\n")
                continue
            text = segment.text.strip()
            if not text:
                continue
            if segment.kind == "hyperlink" and segment.target:
                label = text.replace("\\", "\\\\").replace("[", "\\[").replace("]", "\\]")
                text = f"[{label}]({segment.target})"
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


class SourceDocument(BaseModel):
    """Stable source identity without machine-specific absolute paths."""

    model_config = ConfigDict(extra="forbid")

    name: str = Field(min_length=1)
    format: Literal["pptx", "ppt"]

    @model_validator(mode="after")
    def validate_name_is_basename(self) -> "SourceDocument":
        if Path(self.name).name != self.name:
            raise ValueError("source name must be a basename")
        return self


class BoundingBox(BaseModel):
    """Shape bounds in PowerPoint English Metric Units (EMU)."""

    model_config = ConfigDict(extra="forbid")

    x: int
    y: int
    width: int = Field(gt=0)
    height: int = Field(gt=0)
    unit: Literal["emu"] = "emu"

    @classmethod
    def from_corners(cls, corners: tuple[int, int, int, int]) -> "BoundingBox":
        x1, y1, x2, y2 = corners
        return cls(x=x1, y=y1, width=x2 - x1, height=y2 - y1)


class ContentBlock(BaseModel):
    """A rendered content unit in presentation reading order."""

    model_config = ConfigDict(extra="forbid")

    kind: Literal[
        "text",
        "heading",
        "list",
        "math",
        "image",
        "chart",
        "smartart",
        "table",
        "attachment",
        "unsupported",
    ]
    content: str = Field(min_length=1)
    shape_id: Optional[str] = None
    heading_level: Optional[int] = Field(default=None, ge=1, le=6)
    bbox: Optional[BoundingBox] = None
    source_part: Literal["slide", "layout", "master"] = "slide"

    @model_validator(mode="after")
    def validate_heading_contract(self) -> "ContentBlock":
        if not self.content.strip():
            raise ValueError("content must not be blank")
        if self.kind == "heading" and self.heading_level is None:
            raise ValueError("heading blocks require heading_level")
        if self.kind != "heading" and self.heading_level is not None:
            raise ValueError("heading_level is only valid for heading blocks")
        return self


class SlideDocument(BaseModel):
    """JSON-serializable intermediate representation for one slide."""

    model_config = ConfigDict(extra="forbid")

    page: int = Field(ge=1)
    hidden: bool = False
    blocks: List[ContentBlock] = Field(default_factory=list)
    notes: Optional[str] = None


class PresentationDocument(BaseModel):
    """Canonical intermediate representation shared by all output formats."""

    model_config = ConfigDict(extra="forbid")

    schema_version: Literal["1.0"] = "1.0"
    source: SourceDocument
    slides: List[SlideDocument] = Field(default_factory=list)

    @model_validator(mode="after")
    def validate_slide_order(self) -> "PresentationDocument":
        pages = [slide.page for slide in self.slides]
        if pages != sorted(set(pages)):
            raise ValueError("slide pages must be unique and strictly increasing")
        return self


def render_slide_markdown(slide: SlideDocument) -> str:
    parts = [f"[Page_{slide.page}]"]
    if slide.hidden:
        parts.append("<!-- hidden: true -->")
    for block in slide.blocks:
        content = block.content.strip()
        if not content:
            continue
        if block.kind == "heading" and block.heading_level is not None:
            content = f"{'#' * block.heading_level} {content}"
        parts.append(content)
    if slide.notes and slide.notes.strip():
        parts.extend(["[Speaker_Notes]", slide.notes.strip()])
    return "\n\n".join(parts).rstrip() + "\n"


def render_presentation_markdown(document: PresentationDocument) -> str:
    rendered_slides = [render_slide_markdown(slide).rstrip() for slide in document.slides]
    merged = "\n\n".join(rendered_slides).strip()
    return f"{merged}\n" if merged else ""


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
    attachment_blocks: int = 0
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
            "attachment_blocks": self.attachment_blocks,
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
    attachment_blocks: int = 0
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
        self.attachment_blocks += stats.attachment_blocks
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

    @property
    def output_json_name(self) -> str:
        return f"{self.source_stem}.json"

    def output_json_path(self, output_dir: Path) -> Path:
        return output_dir / self.name / self.output_json_name

    def with_package_dir(self, package_dir: Path) -> "PreparedPackage":
        return self.model_copy(update={"package_dir": package_dir})


# 전체 변환 파이프라인의 정규화된 실행 설정이다.
class ConverterConfig(BaseModel):
    model_config = ConfigDict(arbitrary_types_allowed=True, extra="forbid")

    cwd: Path
    output_dir: Path
    inputs: List[str] = Field(default_factory=list)
    output_format: str = "markdown"
    heading_mode: str = "auto"
    strict: bool = False
    pptx_inheritance: str = "style"
    inherited_shapes: str = "visible"
    ppt_converter: str = "auto"
