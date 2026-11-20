"""Extraction-free PPTX parser that builds the canonical v2 model."""

from __future__ import annotations

import html
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, replace
from pathlib import Path, PurePosixPath
from typing import BinaryIO, Literal

from .assets import AssetLimits, AssetStore, safe_basename
from .attachments import unpack_ole
from .charts import chart_markdown
from .errors import (
    AssetLimitError,
    PackageSecurityError,
    PresentationParseError,
    UnsupportedPresentationError,
)
from .headings import (
    HeadingDetector,
    HeadingHint,
    TextShape,
    heading_hints,
)
from .inheritance import (
    TemplateResolver,
    shape_text,
)
from .inline import (
    finalize,
    finalize_heading,
)
from .model import (
    AttachmentBlock,
    BoundingBox,
    ChartBlock,
    ContentBlock,
    ConversionWarning,
    HeadingBlock,
    ImageBlock,
    PresentationDocument,
    SlideDocument,
    SmartArtBlock,
    SourceDocument,
    TableBlock,
    TextBlock,
    UnsupportedBlock,
)
from .ooxml import (
    OFFICE_REL_NS,
    OOXMLPackage,
    PackageLimits,
    Relationship,
    SlideEntry,
    resolve_relationship_target,
)
from .paragraphs import (
    Paragraph,
    apply_inherited_lists,
    font_sections,
    paragraph_blocks,
    parse_paragraph,
    shape_markdown,
    table_cell_lines,
)
from .reading_order import LayoutNode, geometry_reading_order
from .shapes import (
    NS,
    Shape,
    bbox_model,
    element_shape_id,
    flatten_shapes,
    heading_signals,
    is_decorative,
    is_footer,
    local_name,
    table_overlays,
    with_template_style,
)
from .smartart import smartart_markdown
from .tables import (
    GridCell,
    MergedCells,
    caption_rows,
    cell_at,
    cell_bounds,
    cell_in,
    grid_rows,
    rows_markdown,
    table_grid,
)
from .vector import VectorImages, render_vector_images


@dataclass(frozen=True)
class ParseOptions:
    """Public parsing behavior with explicit resource limits."""

    include_assets: bool = True
    include_notes: bool = True
    include_hidden: bool = True
    strict: bool = True
    merged_cells: MergedCells = "fill"
    # "png" has LibreOffice draw EMF/WMF images as PNG (see ``vector``).
    vector_images: VectorImages = "keep"
    package_limits: PackageLimits = PackageLimits()
    asset_limits: AssetLimits = AssetLimits()


@dataclass(frozen=True)
class ParsedComponents:
    """Internal result consumed by the public ParsedPresentation wrapper."""

    document: PresentationDocument
    assets: AssetStore
    source_path: Path | None


class _WarningCollector:
    def __init__(self) -> None:
        self._warnings: list[ConversionWarning] = []

    def add(
        self,
        code: str,
        message: str,
        *,
        slide: int | None = None,
        shape_id: str | None = None,
        asset_id: str | None = None,
    ) -> None:
        self._warnings.append(
            ConversionWarning(
                code=code,
                message=message,
                slide=slide,
                shape_id=shape_id,
                asset_id=asset_id,
            )
        )

    def values(self) -> tuple[ConversionWarning, ...]:
        return tuple(
            sorted(
                self._warnings,
                key=lambda warning: (
                    warning.slide or 0,
                    warning.shape_id or "",
                    warning.code,
                    warning.asset_id or "",
                    warning.message,
                ),
            )
        )


def _image_markdown(alt_text: str, asset_id: str) -> str:
    alt = alt_text.replace("\\", "\\\\").replace("[", r"\[").replace("]", r"\]")
    return f"![{alt}](asset://{asset_id})"


def _picture_alt_text(element: ET.Element) -> str:
    properties = element.find(".//p:cNvPr", NS)
    if properties is None:
        return "image"
    return properties.attrib.get("descr") or properties.attrib.get("name") or "image"


class _PresentationParser:
    def __init__(
        self, source: Path | bytes, options: ParseOptions, *, name: str | None = None
    ) -> None:
        self.source = source
        self.name = name or (source.name if isinstance(source, Path) else "presentation.pptx")
        self.source_path = source if isinstance(source, Path) else None
        self.options = options
        self.assets = AssetStore(options.asset_limits)
        self.warnings = _WarningCollector()

    def parse(self) -> ParsedComponents:
        try:
            with OOXMLPackage(self.source, self.options.package_limits, name=self.name) as package:
                templates = TemplateResolver(package)
                slides: list[SlideDocument] = []
                for entry in package.slide_entries():
                    try:
                        slide = self._parse_slide(package, entry, templates)
                    except Exception as exc:
                        if self.options.strict:
                            raise PresentationParseError(
                                self.name,
                                str(exc),
                                slide_number=entry.page,
                            ) from exc
                        self.warnings.add(
                            "slide-parse-failed",
                            str(exc),
                            slide=entry.page,
                        )
                        slide = SlideDocument(page=entry.page)
                    if self.options.include_hidden or not slide.hidden:
                        slides.append(slide)

                if self.options.vector_images == "png" and self.options.include_assets:
                    _, failed = render_vector_images(self.assets)
                    for asset_id in failed:
                        self.warnings.add(
                            "vector-image-not-converted",
                            "LibreOffice could not draw this EMF/WMF image; it is kept as is",
                            asset_id=asset_id,
                        )
                document = PresentationDocument(
                    source=SourceDocument(
                        name=safe_basename(self.name, fallback="presentation.pptx"),
                        sha256=package.source_sha256,
                    ),
                    assets=self.assets.descriptors(),
                    slides=tuple(slides),
                    warnings=self.warnings.values(),
                )
                return ParsedComponents(document, self.assets, self.source_path)
        except Exception:
            self.assets.close()
            raise

    def _parse_slide(
        self, package: OOXMLPackage, entry: SlideEntry, templates: TemplateResolver
    ) -> SlideDocument:
        root = package.xml(entry.part_name)
        hidden = str(root.attrib.get("show", "1")).strip().casefold() in {
            "0",
            "false",
            "off",
        }
        tree = root.find("./p:cSld/p:spTree", NS)
        if tree is None:
            raise ValueError("missing p:cSld/p:spTree")
        relationships = package.relationships(entry.part_name)
        template = templates.for_slide(entry.part_name, root)
        shapes = flatten_shapes(tree)
        shapes.extend(
            Shape(
                element=inherited.element,
                tag="sp",
                shape_id=element_shape_id(inherited.element),
                bbox=inherited.bbox,
                source_order=len(shapes) + offset,
                source_part=inherited.source_part,
                template_part=inherited.part_name,
            )
            for offset, inherited in enumerate(template.inherited_text_shapes())
        )
        shapes.extend(
            Shape(
                element=inherited.element,
                tag="pic",
                shape_id=element_shape_id(inherited.element),
                bbox=inherited.bbox,
                source_order=len(shapes) + offset,
                source_part=inherited.source_part,
                template_part=inherited.part_name,
            )
            for offset, inherited in enumerate(template.inherited_pictures())
        )
        styled = [with_template_style(shape, template) for shape in shapes]
        ordered_shapes = geometry_reading_order(
            LayoutNode(
                shape,
                shape.source_order,
                shape.bbox,
                structure=not (
                    is_footer(shape) or is_decorative(shape, shape_text(shape.element))
                ),
            )
            for shape in styled
        )
        hints = heading_hints([heading_signals(shape) for shape in ordered_shapes])
        detector = HeadingDetector()
        overlays = table_overlays(ordered_shapes)
        in_tables = {id(picture.element) for pictures in overlays.values() for picture in pictures}
        blocks: list[ContentBlock] = []
        for shape, hint in zip(ordered_shapes, hints):
            if id(shape.element) in in_tables:
                continue
            if shape.tag in {"sp", "cxnSp"}:
                shape_relationships = (
                    package.relationships(shape.template_part)
                    if shape.template_part is not None
                    else relationships
                )
                if shape.element.find("./p:spPr/a:blipFill/a:blip", NS) is not None:
                    # A shape filled with a picture shows that picture.
                    blocks.extend(self._picture_blocks(package, entry, shape, shape_relationships))
                blocks.extend(
                    self._text_shape_blocks(entry, shape, shape_relationships, hint, detector)
                )
            else:
                blocks.extend(
                    self._shape_blocks(
                        package,
                        entry,
                        shape,
                        package.relationships(shape.template_part)
                        if shape.template_part is not None
                        else relationships,
                        overlays.get(id(shape.element), []),
                    )
                )
        heading_indexes = [i for i, block in enumerate(blocks) if isinstance(block, HeadingBlock)]
        if len(heading_indexes) == 1:
            # A slide's only heading is its title.
            index = heading_indexes[0]
            blocks[index] = blocks[index].model_copy(update={"level": 1})
        notes = self._notes(package, entry.part_name) if self.options.include_notes else None
        return SlideDocument(
            page=entry.page,
            hidden=hidden,
            blocks=tuple(
                block.model_copy(update={"id": f"s{entry.page}.b{index}"})
                for index, block in enumerate(blocks, 1)
            ),
            notes=notes,
        )

    def _shape_blocks(
        self,
        package: OOXMLPackage,
        entry: SlideEntry,
        shape: Shape,
        relationships: dict[str, Relationship],
        overlays: list[Shape],
    ) -> list[ContentBlock]:
        if shape.tag == "pic":
            return self._picture_blocks(package, entry, shape, relationships)
        if shape.tag == "graphicFrame":
            return self._graphic_frame_blocks(package, entry, shape, relationships, overlays)
        return []

    def _text_shape_blocks(
        self,
        entry: SlideEntry,
        shape: Shape,
        relationships: dict[str, Relationship],
        hint: HeadingHint,
        detector: HeadingDetector,
    ) -> list[ContentBlock]:
        body = shape.element.find("./p:txBody", NS)
        if body is None or is_footer(shape):
            return []
        paragraphs: list[Paragraph] = []
        unconverted_math = 0
        for element in body.findall("./a:p", NS):
            paragraph, unconverted = parse_paragraph(element, relationships)
            unconverted_math += unconverted
            if paragraph.text:
                paragraphs.append(paragraph)
        if unconverted_math:
            self.warnings.add(
                "math-unconverted",
                f"{unconverted_math} equation(s) could not be converted to LaTeX",
                slide=entry.page,
                shape_id=shape.shape_id,
            )
        if not paragraphs:
            return []

        paragraphs = apply_inherited_lists(paragraphs, shape.style)
        if all(p.uniform_emphasis for p in paragraphs) and (
            len(paragraphs) > 1 or sum(len(p.plain) for p in paragraphs) > 60
        ):
            # A whole text box set in bold is styled, not emphasised; a short bold
            # label on its own keeps its emphasis as a hint that it is a subheading.
            paragraphs = [replace(p, rich=p.text) for p in paragraphs]
        heading = detector.classify(
            TextShape(
                markdown=shape_markdown(paragraphs),
                plain=" ".join(p.plain for p in paragraphs if p.plain),
                placeholder_type=shape.style.placeholder_type,
                font_pt=shape.style.font_pt,
                has_list=any(p.is_list for p in paragraphs),
                has_math=any(p.math for p in paragraphs),
                font_sections=font_sections(paragraphs),
            ),
            hint,
        )
        bbox = bbox_model(shape.bbox)
        if heading == "duplicate":
            return []
        if heading is None:
            return paragraph_blocks(paragraphs, shape=shape, bbox=bbox)
        provenance = {"shape_id": shape.shape_id, "bbox": bbox, "source_part": shape.source_part}
        blocks: list[ContentBlock] = [
            HeadingBlock(
                text=finalize_heading(heading.text),
                level=heading.level,
                source=heading.source,
                **provenance,
            )
        ]
        if heading.trailing:
            blocks.append(TextBlock(text=finalize(heading.trailing), **provenance))
        return blocks

    def _register_related_asset(
        self,
        package: OOXMLPackage,
        *,
        source_part: str,
        relationship: Relationship,
        kind: Literal["image", "attachment"],
        slide: int,
        shape_id: str | None,
    ):
        if relationship.external:
            self.warnings.add(
                "external-asset-ignored",
                f"external asset was not downloaded: {relationship.target}",
                slide=slide,
                shape_id=shape_id,
            )
            return None
        target = resolve_relationship_target(source_part, relationship)
        if target is None or not package.exists(target):
            self.warnings.add(
                "asset-target-missing",
                f"asset relationship target is unavailable: {relationship.target}",
                slide=slide,
                shape_id=shape_id,
            )
            return None
        return self._register_part(package, target, kind=kind, slide=slide, shape_id=shape_id)

    def _register_part(
        self,
        package: OOXMLPackage,
        target: str,
        *,
        kind: Literal["image", "attachment"],
        slide: int,
        shape_id: str | None,
    ):
        try:
            with package.open_part(target) as stream:
                return self.assets.register_stream(
                    stream,
                    logical_key=target,
                    kind=kind,
                    original_name=PurePosixPath(target).name,
                    source_part=target,
                    preserve_payload=self.options.include_assets,
                )
        except AssetLimitError as exc:
            if self.options.strict:
                raise
            self.warnings.add(
                "asset-limit-exceeded",
                str(exc),
                slide=slide,
                shape_id=shape_id,
            )
            return None

    def _smartart_blocks(
        self,
        package: OOXMLPackage,
        entry: SlideEntry,
        shape: Shape,
        data_part: str,
        bbox: BoundingBox | None,
    ) -> list[ContentBlock]:
        provenance = {"shape_id": shape.shape_id, "bbox": bbox}
        try:
            conversion = smartart_markdown(package, data_part)
        except (PackageSecurityError, AssetLimitError):
            raise
        except Exception as exc:  # smartart2md on a malformed data model
            self.warnings.add(
                "smartart-conversion-failed",
                f"SmartArt could not be converted: {type(exc).__name__}: {exc}",
                slide=entry.page,
                shape_id=shape.shape_id,
            )
            return [UnsupportedBlock(message="SmartArt could not be converted", **provenance)]

        asset_ids: list[str] = []

        def picture(part: str) -> str:
            descriptor = self._register_part(
                package, part, kind="image", slide=entry.page, shape_id=shape.shape_id
            )
            if descriptor is None:
                return ""
            asset_ids.append(descriptor.id)
            return _image_markdown("image", descriptor.id)

        markdown = conversion.render(picture)
        if markdown:
            return [
                SmartArtBlock(
                    markdown=markdown, asset_ids=tuple(dict.fromkeys(asset_ids)), **provenance
                )
            ]
        if conversion.content_points:
            count = conversion.content_points
            label = f"{count} unlabeled node{'' if count == 1 else 's'}"
            return [SmartArtBlock(markdown=f"[smartart: {label}]", **provenance)]
        self.warnings.add(
            "smartart-empty",
            "SmartArt has no content nodes",
            slide=entry.page,
            shape_id=shape.shape_id,
        )
        return []

    def _ole_attachment(
        self,
        package: OOXMLPackage,
        entry: SlideEntry,
        shape: Shape,
        ole: ET.Element,
        relationships: dict[str, Relationship],
    ):
        """The embedded file, unpacked from its OLE compound file when it has one."""

        relationship = relationships.get(ole.attrib.get(f"{{{OFFICE_REL_NS}}}id", ""))
        if relationship is None:
            return None
        target = (
            None
            if relationship.external
            else resolve_relationship_target(entry.part_name, relationship)
        )
        if target is not None and target.casefold().endswith(".bin") and package.exists(target):
            limit = self.options.asset_limits.max_single_asset_bytes
            with package.open_part(target) as stream:
                data = stream.read(limit + 1)
            unpacked = None
            if len(data) <= limit:
                stem = safe_basename(
                    ole.attrib.get("name") or PurePosixPath(target).stem, fallback="attachment"
                )
                try:
                    unpacked = unpack_ole(data, prog_id=ole.attrib.get("progId", ""), stem=stem)
                except ValueError as exc:
                    self.warnings.add(
                        "ole-unpack-failed",
                        f"embedded OLE object kept as stored: {exc}",
                        slide=entry.page,
                        shape_id=shape.shape_id,
                    )
            if len(data) <= limit:
                if unpacked is not None:
                    payload, name, key = unpacked.payload, unpacked.name, f"{target}#payload"
                else:
                    # Kept as stored, it is still better named after the object
                    # ("Slide.bin") than after its part ("oleObject2.bin").
                    payload, name, key = data, f"{stem}.bin", target
                try:
                    return self.assets.register_bytes(
                        payload,
                        logical_key=key,
                        kind="attachment",
                        original_name=name,
                        source_part=target,
                        preserve_payload=self.options.include_assets,
                    )
                except AssetLimitError as exc:
                    if self.options.strict:
                        raise
                    self.warnings.add(
                        "asset-limit-exceeded", str(exc), slide=entry.page, shape_id=shape.shape_id
                    )
                    return None
        return self._register_related_asset(
            package,
            source_part=entry.part_name,
            relationship=relationship,
            kind="attachment",
            slide=entry.page,
            shape_id=shape.shape_id,
        )

    def _picture_blocks(
        self,
        package: OOXMLPackage,
        entry: SlideEntry,
        shape: Shape,
        relationships: dict[str, Relationship],
    ) -> list[ContentBlock]:
        blocks: list[ContentBlock] = []
        bbox = bbox_model(shape.bbox)
        # A picture drawn from the layout or master resolves its links from that part.
        part_name = shape.template_part or entry.part_name
        blip = shape.element.find(".//a:blip", NS)
        rid = blip.attrib.get(f"{{{OFFICE_REL_NS}}}embed", "") if blip is not None else ""
        relationship = relationships.get(rid)
        if relationship is not None:
            descriptor = self._register_related_asset(
                package,
                source_part=part_name,
                relationship=relationship,
                kind="image",
                slide=entry.page,
                shape_id=shape.shape_id,
            )
            if descriptor is not None:
                properties = shape.element.find(".//p:cNvPr", NS)
                alt_text = (
                    properties.attrib.get("descr") or properties.attrib.get("name") or "image"
                    if properties is not None
                    else "image"
                )
                blocks.append(
                    ImageBlock(
                        asset_id=descriptor.id,
                        alt_text=alt_text,
                        shape_id=shape.shape_id,
                        bbox=bbox,
                        source_part=shape.source_part,
                    )
                )
        elif rid:
            self.warnings.add(
                "image-relationship-missing",
                f"image relationship not found: {rid}",
                slide=entry.page,
                shape_id=shape.shape_id,
            )

        for tag in ("audioFile", "videoFile"):
            media = shape.element.find(f".//a:{tag}", NS)
            media_rid = (
                media.attrib.get(f"{{{OFFICE_REL_NS}}}link", "") if media is not None else ""
            )
            media_relationship = relationships.get(media_rid)
            if media_relationship is None:
                continue
            descriptor = self._register_related_asset(
                package,
                source_part=part_name,
                relationship=media_relationship,
                kind="attachment",
                slide=entry.page,
                shape_id=shape.shape_id,
            )
            if descriptor is not None:
                blocks.append(
                    AttachmentBlock(
                        asset_id=descriptor.id,
                        label=descriptor.original_name,
                        shape_id=shape.shape_id,
                        bbox=bbox,
                    )
                )
        return blocks

    def _chart_block(
        self,
        package: OOXMLPackage,
        entry: SlideEntry,
        shape: Shape,
        part_name: str,
        bbox: BoundingBox | None,
    ) -> ChartBlock | UnsupportedBlock | None:
        try:
            conversion = chart_markdown(package, part_name)
        except (PresentationParseError, PackageSecurityError):
            raise
        except Exception as exc:  # third-party converter: report and keep the slide
            self.warnings.add(
                "chart-conversion-failed",
                f"chart conversion failed: {type(exc).__name__}: {exc}",
                slide=entry.page,
                shape_id=shape.shape_id,
            )
            return UnsupportedBlock(
                message="chart conversion failed", shape_id=shape.shape_id, bbox=bbox
            )
        for message in conversion.rejected_workbooks:
            self.warnings.add(
                "chart-workbook-rejected",
                f"embedded chart workbook rejected; using cached chart values: {message}",
                slide=entry.page,
                shape_id=shape.shape_id,
            )
        if not conversion.markdown:
            return None
        return ChartBlock(markdown=conversion.markdown, shape_id=shape.shape_id, bbox=bbox)

    def _table_block(
        self,
        package: OOXMLPackage,
        entry: SlideEntry,
        shape: Shape,
        table: ET.Element,
        relationships: dict[str, Relationship],
        overlays: list[Shape],
    ) -> TableBlock | None:
        unconverted_math = 0

        def cell_lines(cell: ET.Element) -> list[str]:
            nonlocal unconverted_math
            lines, missing = table_cell_lines(cell, relationships)
            unconverted_math += missing
            return lines

        grid = table_grid(table, cell_lines)
        if not grid:
            return None
        if unconverted_math:
            self.warnings.add(
                "math-unconverted",
                f"{unconverted_math} equation(s) could not be converted to LaTeX",
                slide=entry.page,
                shape_id=shape.shape_id,
            )

        asset_ids: list[str] = []

        def add_image(cell: GridCell | None, rid: str, alt_text: str, owner: str | None) -> None:
            relationship = relationships.get(rid)
            if relationship is None:
                if rid:
                    self.warnings.add(
                        "image-relationship-missing",
                        f"image relationship not found: {rid}",
                        slide=entry.page,
                        shape_id=owner,
                    )
                return
            descriptor = self._register_related_asset(
                package,
                source_part=entry.part_name,
                relationship=relationship,
                kind="image",
                slide=entry.page,
                shape_id=owner,
            )
            if descriptor is None or cell is None:
                return
            cell.lines.append(_image_markdown(alt_text, descriptor.id))
            asset_ids.append(descriptor.id)

        bounds = cell_bounds(table, shape.bbox) if shape.bbox is not None else None
        for picture in overlays:
            left, top, right, bottom = picture.bbox  # type: ignore[misc]
            cell = (
                cell_at(grid, bounds, ((left + right) // 2, (top + bottom) // 2))
                if bounds
                else None
            )
            blip = picture.element.find(".//a:blip", NS)
            rid = blip.attrib.get(f"{{{OFFICE_REL_NS}}}embed", "") if blip is not None else ""
            add_image(cell, rid, _picture_alt_text(picture.element), picture.shape_id)
        for row_index, row in enumerate(table.findall("./a:tr", NS)):
            for column, element in enumerate(row.findall("./a:tc", NS)):
                blip = element.find("./a:tcPr/a:blipFill/a:blip", NS)
                if blip is not None:
                    rid = blip.attrib.get(f"{{{OFFICE_REL_NS}}}embed", "")
                    add_image(cell_in(grid, row_index, column), rid, "image", shape.shape_id)

        rows = grid_rows(grid)
        captions = caption_rows(rows)
        return TableBlock(
            markdown=rows_markdown(
                rows, merged_cells=self.options.merged_cells, caption_rows=captions
            ),
            rows=rows,
            caption_rows=captions,
            merged_cells=self.options.merged_cells,
            asset_ids=tuple(dict.fromkeys(asset_ids)),
            shape_id=shape.shape_id,
            bbox=bbox_model(shape.bbox),
        )

    def _graphic_frame_blocks(
        self,
        package: OOXMLPackage,
        entry: SlideEntry,
        shape: Shape,
        relationships: dict[str, Relationship],
        overlays: list[Shape],
    ) -> list[ContentBlock]:
        bbox = bbox_model(shape.bbox)
        table = shape.element.find(".//a:tbl", NS)
        if table is not None:
            block = self._table_block(package, entry, shape, table, relationships, overlays)
            return [block] if block is not None else []

        chart = shape.element.find(".//c:chart", NS)
        if chart is None:
            chart = shape.element.find(".//cx:chart", NS)
        if chart is not None:
            rid = chart.attrib.get(f"{{{OFFICE_REL_NS}}}id", "")
            relationship = relationships.get(rid)
            if relationship is not None and not relationship.external:
                target = resolve_relationship_target(entry.part_name, relationship)
                if target is not None and package.exists(target):
                    block = self._chart_block(package, entry, shape, target, bbox)
                    if block is not None:
                        return [block]
            self.warnings.add(
                "chart-data-unavailable",
                "chart relationship or cached data is unavailable",
                slide=entry.page,
                shape_id=shape.shape_id,
            )
            return [
                UnsupportedBlock(
                    message="chart data unavailable",
                    shape_id=shape.shape_id,
                    bbox=bbox,
                )
            ]

        diagram = shape.element.find(".//dgm:relIds", NS)
        if diagram is not None:
            rid = diagram.attrib.get(f"{{{OFFICE_REL_NS}}}dm", "")
            relationship = relationships.get(rid)
            if relationship is not None and not relationship.external:
                target = resolve_relationship_target(entry.part_name, relationship)
                if target is not None and package.exists(target):
                    return self._smartart_blocks(package, entry, shape, target, bbox)
            self.warnings.add(
                "smartart-data-unavailable",
                "SmartArt data relationship is unavailable",
                slide=entry.page,
                shape_id=shape.shape_id,
            )
            return [
                UnsupportedBlock(
                    message="SmartArt data unavailable",
                    shape_id=shape.shape_id,
                    bbox=bbox,
                )
            ]

        ole = shape.element.find(".//p:oleObj", NS)
        if ole is not None:
            descriptor = self._ole_attachment(package, entry, shape, ole, relationships)
            if descriptor is not None:
                return [
                    AttachmentBlock(
                        asset_id=descriptor.id,
                        label=descriptor.original_name,
                        program=ole.attrib.get("progId") or None,
                        shape_id=shape.shape_id,
                        bbox=bbox,
                    )
                ]

        model = next(
            (element for element in shape.element.iter() if local_name(element.tag) == "model3d"),
            None,
        )
        if model is not None:
            model_blocks: list[ContentBlock] = []
            preview = next(
                (element for element in model.iter() if local_name(element.tag) == "blip"),
                None,
            )
            preview_rid = (
                preview.attrib.get(f"{{{OFFICE_REL_NS}}}embed", "") if preview is not None else ""
            )
            preview_relationship = relationships.get(preview_rid)
            if preview_relationship is not None:
                preview_descriptor = self._register_related_asset(
                    package,
                    source_part=entry.part_name,
                    relationship=preview_relationship,
                    kind="image",
                    slide=entry.page,
                    shape_id=shape.shape_id,
                )
                if preview_descriptor is not None:
                    properties = shape.element.find(".//p:cNvPr", NS)
                    alt_text = (
                        properties.attrib.get("descr")
                        or properties.attrib.get("name")
                        or "3D model preview"
                        if properties is not None
                        else "3D model preview"
                    )
                    model_blocks.append(
                        ImageBlock(
                            asset_id=preview_descriptor.id,
                            alt_text=alt_text,
                            shape_id=shape.shape_id,
                            bbox=bbox,
                        )
                    )
            rid = model.attrib.get(f"{{{OFFICE_REL_NS}}}embed", "")
            relationship = relationships.get(rid)
            if relationship is not None:
                descriptor = self._register_related_asset(
                    package,
                    source_part=entry.part_name,
                    relationship=relationship,
                    kind="attachment",
                    slide=entry.page,
                    shape_id=shape.shape_id,
                )
                if descriptor is not None:
                    model_blocks.append(
                        AttachmentBlock(
                            asset_id=descriptor.id,
                            label=descriptor.original_name,
                            shape_id=shape.shape_id,
                            bbox=bbox,
                        )
                    )
            if model_blocks:
                return model_blocks

        self.warnings.add(
            "unsupported-graphic-frame",
            "graphic frame type is not supported",
            slide=entry.page,
            shape_id=shape.shape_id,
        )
        return [
            UnsupportedBlock(
                message="unsupported graphic frame",
                shape_id=shape.shape_id,
                bbox=bbox,
            )
        ]

    def _notes(self, package: OOXMLPackage, slide_part: str) -> str | None:
        relationships = package.relationships(slide_part)
        notes_relationship = next(
            (
                relationship
                for relationship in relationships.values()
                if relationship.type.endswith("/notesSlide") and not relationship.external
            ),
            None,
        )
        if notes_relationship is None:
            return None
        target = resolve_relationship_target(slide_part, notes_relationship)
        if target is None or not package.exists(target):
            return None
        root = package.xml(target)
        paragraphs: list[str] = []
        for shape in root.findall(".//p:sp", NS):
            placeholder = shape.find("./p:nvSpPr/p:nvPr/p:ph", NS)
            placeholder_type = (
                placeholder.attrib.get("type", "body") if placeholder is not None else ""
            )
            if placeholder_type in {"hdr", "ftr", "dt", "sldNum"}:
                continue
            # LibreOffice writes the page number as a plain text box on the notes page.
            if placeholder is None and shape_text(shape).strip().isdigit():
                continue
            for paragraph in shape.findall("./p:txBody/a:p", NS):
                text = "".join(node.text or "" for node in paragraph.findall(".//a:t", NS))
                text = re.sub(r"\s+", " ", html.unescape(text)).strip()
                if text:
                    paragraphs.append(text)
        return "\n".join(paragraphs) or None


def parse_components(
    source: str | Path | bytes | BinaryIO,
    *,
    options: ParseOptions | None = None,
    name: str | None = None,
) -> ParsedComponents:
    """Parse one PPTX, from a local path or in memory, into canonical components.

    ``source`` may be a path, the file's bytes, or a binary stream (read to its end,
    up to ``PackageLimits.max_package_bytes``). ``name`` is the file name recorded in
    the document for in-memory input.
    """

    options = options or ParseOptions()
    if isinstance(source, (str, Path)):
        path = Path(source).expanduser()
        if path.suffix.casefold() != ".pptx":
            raise UnsupportedPresentationError("only .pptx files are supported")
        if not path.is_file():
            raise FileNotFoundError(path)
        payload: Path | bytes = path.resolve()
        label = name or path.name
    else:
        limit = options.package_limits.max_package_bytes
        if isinstance(source, (bytes, bytearray, memoryview)):
            data = bytes(source)
        elif hasattr(source, "read"):
            data = source.read(limit + 1)
            if not isinstance(data, bytes):
                raise TypeError("the stream must be opened in binary mode")
        else:
            raise TypeError("source must be a path, bytes or a binary stream")
        if len(data) > limit:
            raise PackageSecurityError(f"presentation exceeds {limit} bytes")
        # A compound file goes on to the package, which explains encryption or .ppt.
        if not data.startswith((b"PK", b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")):
            raise UnsupportedPresentationError("input is not a .pptx (ZIP) package")
        payload = data
        label = name or "presentation.pptx"
    try:
        return _PresentationParser(payload, options, name=label).parse()
    except (PresentationParseError, PackageSecurityError, AssetLimitError):
        raise
    except Exception as exc:
        raise PresentationParseError(label, str(exc)) from exc
