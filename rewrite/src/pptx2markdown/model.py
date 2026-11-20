"""Path-independent canonical document models."""

from __future__ import annotations

from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator


class SourceDocument(BaseModel):
    """Stable source identity without a machine-local path."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    name: str = Field(min_length=1)
    format: Literal["pptx"] = "pptx"
    sha256: str = Field(pattern=r"^[0-9a-f]{64}$")

    @model_validator(mode="after")
    def validate_basename(self) -> "SourceDocument":
        if "/" in self.name or "\\" in self.name or self.name in {".", ".."}:
            raise ValueError("source name must be a basename")
        return self


class BoundingBox(BaseModel):
    """Shape bounds in PowerPoint English Metric Units."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    x: int
    y: int
    width: int = Field(gt=0)
    height: int = Field(gt=0)
    unit: Literal["emu"] = "emu"


class AssetDescriptor(BaseModel):
    """Serializable metadata for one logical embedded asset."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    id: str = Field(pattern=r"^asset-[0-9a-f]{16}(?:-[0-9a-f]{8})?$")
    kind: Literal["image", "attachment"]
    original_name: str = Field(min_length=1)
    media_type: str = Field(min_length=1)
    size: int = Field(ge=0)
    sha256: str = Field(pattern=r"^[0-9a-f]{64}$")
    source_part: str = Field(min_length=1)
    payload_available: bool = True

    @model_validator(mode="after")
    def validate_names(self) -> "AssetDescriptor":
        if "/" in self.original_name or "\\" in self.original_name:
            raise ValueError("asset original_name must be a basename")
        if self.source_part.startswith("/") or ".." in self.source_part.split("/"):
            raise ValueError("asset source_part must be a safe package-relative path")
        return self


class ConversionWarning(BaseModel):
    """Structured non-fatal conversion diagnostic."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    code: str = Field(min_length=1)
    message: str = Field(min_length=1)
    slide: int | None = Field(default=None, ge=1)
    shape_id: str | None = None
    asset_id: str | None = None


BLOCK_ID_PATTERN = r"^s[1-9][0-9]*\.b[1-9][0-9]*$"


class BlockBase(BaseModel):
    """Fields shared by every block in reading order.

    ``id`` is ``s<page>.b<n>``, the block's position on its slide; recipes and
    reviews address blocks by it.
    """

    model_config = ConfigDict(extra="forbid", frozen=True)

    id: str | None = Field(default=None, pattern=BLOCK_ID_PATTERN)
    shape_id: str | None = None
    bbox: BoundingBox | None = None
    source_part: Literal["slide", "layout", "master"] = "slide"


class TextBlock(BlockBase):
    """Paragraph lines as inline Markdown.

    ``text`` of text, heading and list blocks is escaped inline Markdown: links,
    ``$inline math$`` and ``**bold**``/``*italic*`` are markup, and source characters
    that Markdown would interpret are backslash-escaped. Slide ``notes`` stay plain.
    """

    kind: Literal["text"] = "text"
    text: str = Field(min_length=1)


class HeadingBlock(BlockBase):
    """A heading; ``source`` says why: a title ``placeholder``, a placeholder at the
    layout's title ``position``, section ``numbering``, heading-like ``text``, or
    ``font`` size alone (the weakest signal)."""

    kind: Literal["heading"] = "heading"
    text: str = Field(min_length=1)
    level: int = Field(ge=1, le=6)
    source: Literal["placeholder", "position", "numbering", "text", "font"] | None = None


class ListItem(BaseModel):
    model_config = ConfigDict(extra="forbid", frozen=True)

    text: str = Field(min_length=1)
    level: int = Field(default=0, ge=0, le=8)


class ListBlock(BlockBase):
    """``start`` is the number of an ordered list's first top-level item."""

    kind: Literal["list"] = "list"
    items: tuple[ListItem, ...] = Field(min_length=1)
    ordered: bool = False
    start: int = Field(default=1, ge=1)


class MathBlock(BlockBase):
    kind: Literal["math"] = "math"
    latex: str = Field(min_length=1)


class ImageBlock(BlockBase):
    kind: Literal["image"] = "image"
    asset_id: str
    alt_text: str = "image"
    description: str | None = None


class MarkdownBlockBase(BlockBase):
    markdown: str = Field(min_length=1)
    asset_ids: tuple[str, ...] = ()


class ChartBlock(MarkdownBlockBase):
    kind: Literal["chart"] = "chart"


class SmartArtBlock(MarkdownBlockBase):
    kind: Literal["smartart"] = "smartart"


MergedCells = Literal["fill", "header", "blank"]


class TableCell(BaseModel):
    """One grid position of a table.

    A cell that starts a merge carries the text and its ``row_span``/``col_span``;
    the positions it covers have ``merged_from`` set to its ``(row, column)``.
    """

    model_config = ConfigDict(extra="forbid", frozen=True)

    text: str = ""
    row_span: int = Field(default=1, ge=1)
    col_span: int = Field(default=1, ge=1)
    merged_from: tuple[int, int] | None = None


class TableBlock(MarkdownBlockBase):
    """A native table: ``rows`` is the full grid, ``markdown`` its rendering.

    Cell ``text`` is inline Markdown with one line per paragraph. ``markdown``
    renders the grid: the first ``caption_rows`` rows as a title line, the next
    ``header_rows`` rows as the header, and merged cells filled according to
    ``merged_cells``.
    """

    kind: Literal["table"] = "table"
    rows: tuple[tuple[TableCell, ...], ...] = ()
    header_rows: int = Field(default=1, ge=0)
    caption_rows: int = Field(default=0, ge=0)
    merged_cells: MergedCells = "fill"

    @model_validator(mode="after")
    def validate_merges(self) -> "TableBlock":
        for row_index, row in enumerate(self.rows):
            for column, cell in enumerate(row):
                if cell.merged_from is None:
                    continue
                origin_row, origin_column = cell.merged_from
                if (
                    not (origin_row <= row_index and origin_column <= column)
                    or (origin_row, origin_column) == (row_index, column)
                    or origin_column >= len(self.rows[origin_row])
                    or self.rows[origin_row][origin_column].merged_from is not None
                ):
                    raise ValueError(
                        f"cell ({row_index}, {column}) must be merged from an earlier merge start"
                    )
        return self


class AttachmentBlock(BlockBase):
    """An embedded file; ``label`` is its file name, ``program`` its OLE ProgID."""

    kind: Literal["attachment"] = "attachment"
    asset_id: str
    label: str = Field(min_length=1)
    program: str | None = None


class UnsupportedBlock(BlockBase):
    kind: Literal["unsupported"] = "unsupported"
    message: str = Field(min_length=1)


ContentBlock = Annotated[
    TextBlock
    | HeadingBlock
    | ListBlock
    | MathBlock
    | ImageBlock
    | ChartBlock
    | SmartArtBlock
    | TableBlock
    | AttachmentBlock
    | UnsupportedBlock,
    Field(discriminator="kind"),
]


class SlideDocument(BaseModel):
    """Canonical content for one slide."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    page: int = Field(ge=1)
    hidden: bool = False
    blocks: tuple[ContentBlock, ...] = ()
    notes: str | None = None


class PresentationDocument(BaseModel):
    """Canonical, deterministic and path-independent presentation model."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    schema_version: Literal["2.0"] = "2.0"
    source: SourceDocument
    assets: tuple[AssetDescriptor, ...] = ()
    slides: tuple[SlideDocument, ...] = ()
    warnings: tuple[ConversionWarning, ...] = ()

    @model_validator(mode="after")
    def validate_document_graph(self) -> "PresentationDocument":
        pages = [slide.page for slide in self.slides]
        if pages != sorted(set(pages)):
            raise ValueError("slide pages must be unique and strictly increasing")

        block_ids = [block.id for slide in self.slides for block in slide.blocks if block.id]
        if len(block_ids) != len(set(block_ids)):
            raise ValueError("block ids must be unique")

        descriptors = {asset.id: asset for asset in self.assets}
        if len(descriptors) != len(self.assets):
            raise ValueError("asset ids must be unique")
        if tuple(asset.id for asset in self.assets) != tuple(sorted(descriptors)):
            raise ValueError("assets must be sorted by id")

        referenced: set[str] = set()
        for slide in self.slides:
            for block in slide.blocks:
                if isinstance(block, (ImageBlock, AttachmentBlock)):
                    referenced.add(block.asset_id)
                elif isinstance(block, MarkdownBlockBase):
                    referenced.update(block.asset_ids)
        missing = sorted(referenced - descriptors.keys())
        if missing:
            raise ValueError(f"blocks reference unknown assets: {', '.join(missing)}")

        warning_asset_ids = {
            warning.asset_id for warning in self.warnings if warning.asset_id is not None
        }
        missing_warning_assets = sorted(warning_asset_ids - descriptors.keys())
        if missing_warning_assets:
            raise ValueError(
                "warnings reference unknown assets: " + ", ".join(missing_warning_assets)
            )
        return self
