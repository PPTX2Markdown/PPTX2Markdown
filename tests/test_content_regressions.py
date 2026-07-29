from __future__ import annotations

import subprocess
import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from unittest.mock import patch

from pptx2markdown.main_converter.asset_utils import convert_vector_assets_to_png
from pptx2markdown.main_converter.converter_models import (
    ContentBlock,
    PresentationDocument,
    SlideDocument,
    SourceDocument,
)
from pptx2markdown.main_converter.run_pptx_to_markdown import (
    _ooxml_part_from_relationship,
    _rewrite_converted_vector_links,
    extract_shape_blocks,
    extract_speaker_notes,
    format_markdown_image,
    graphic_frame_kind,
    load_effective_properties,
    load_heading_hints,
    paragraph_has_list_semantics,
    relativize_markdown_path,
    render_image_tag,
    render_shape_blocks,
)
from pptx2markdown.main_converter.slide_converter import (
    FlattenedShape,
    SlideConversionContext,
    _apply_effective_list_properties,
    _expanded_children,
    _is_hidden_slide,
    _ordered_flattened_shapes,
    _with_effective_bboxes,
)
from pptx2markdown.main_converter.table_overlay import inject_table_run_hyperlinks
from pptx2markdown.pptx_inheritance.resolver import (
    _is_placeholder_prompt_text,
    _is_synthetic_literal_date_anchor,
    _is_unpositioned_literal_date_placeholder,
    _is_visible_materialized_shape,
    is_decorative,
)
from pptx2markdown.structure_analyzer.constants import NS
from pptx2markdown.structure_analyzer.pipeline import materialize_tree_by_objects
from pptx2markdown.structure_analyzer.structure import SlideObject
from pptx2markdown.structure_analyzer.text_rules import is_numbered_heading_text, normalize_text
from pptx2markdown.structure_analyzer.xml_primitives import (
    extract_bbox_emu,
    has_slide_number_field,
    is_slide_number_only_shape,
    text_without_slide_number_fields,
)
from pptx2markdown.table_pipeline.parse import parse_table_element
from pptx2markdown.table_pipeline.render import render_parsed_table_to_markdown


def _slide_object(**overrides: object) -> SlideObject:
    values: dict[str, object] = {
        "shape_id": "master:7",
        "xml_index": 2_000_001,
        "tag": "sp",
        "name": "Master text",
        "ph_type": None,
        "ph_idx": None,
        "x": 100,
        "y": 200,
        "coord_source": "master",
        "text": "Inherited text",
        "normalized": "inherited text",
        "is_footer": False,
        "is_decorative": False,
        "is_heading": False,
        "is_title_placeholder": False,
        "font_pt": 18.0,
        "bbox": (100, 200, 1100, 700),
        "source_part": "master",
        "inheritance_kind": "materialized",
    }
    values.update(overrides)
    return SlideObject(**values)


def _empty_slide_tree() -> ET.ElementTree:
    xml = f"""
    <p:sld xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">
      <p:cSld>
        <p:spTree>
          <p:nvGrpSpPr/>
          <p:grpSpPr/>
        </p:spTree>
      </p:cSld>
    </p:sld>
    """
    return ET.ElementTree(ET.fromstring(xml))


class ContentRegressionTests(unittest.TestCase):
    def test_images_use_standard_markdown_syntax(self) -> None:
        self.assertEqual(render_image_tag("media/image9.wmf"), "![image](media/image9.wmf)")
        self.assertEqual(
            render_image_tag("media/diagram (final).png"),
            "![image](<media/diagram (final).png>)",
        )

    def test_vector_image_conversion_is_deferred_until_postprocessing(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "image9.wmf"
            source.write_bytes(b"wmf")
            output_dir = root / "output"

            rendered = format_markdown_image(
                str(source),
                output_dir=output_dir,
                media_dir=output_dir / "media",
            )

            self.assertTrue((output_dir / "media" / "image9.wmf").is_file())
            self.assertEqual(rendered, "![image](media/image9.wmf)")

    def test_vector_assets_are_converted_in_one_libreoffice_process(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            media_dir = root / "media"
            media_dir.mkdir()
            emf = media_dir / "image1.emf"
            wmf = media_dir / "image2.wmf"
            emf.write_bytes(b"emf")
            wmf.write_bytes(b"wmf")

            def create_converted_files(command: list[str], **_kwargs: object) -> None:
                output_index = command.index("--outdir") + 1
                output_dir = Path(command[output_index])
                for staged_path in command[output_index + 1 :]:
                    staged = Path(staged_path)
                    (output_dir / f"{staged.stem}.png").write_bytes(b"png")

            with (
                patch(
                    "pptx2markdown.main_converter.asset_utils._resolve_soffice_cmd",
                    return_value="/test/soffice",
                ),
                patch("pptx2markdown.main_converter.asset_utils.shutil.which", return_value=None),
                patch(
                    "pptx2markdown.main_converter.asset_utils.subprocess.run",
                    side_effect=create_converted_files,
                ) as run,
            ):
                converted = convert_vector_assets_to_png([emf, wmf, emf], media_dir)

            self.assertEqual(run.call_count, 1)
            self.assertEqual(set(converted), {emf.resolve(), wmf.resolve()})
            self.assertEqual({path.suffix for path in converted.values()}, {".png"})

    def test_failed_vector_batch_preserves_original_assets(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            media_dir = Path(temporary)
            source = media_dir / "image1.emf"
            source.write_bytes(b"emf")

            with (
                patch(
                    "pptx2markdown.main_converter.asset_utils._resolve_soffice_cmd",
                    return_value="/test/soffice",
                ),
                patch("pptx2markdown.main_converter.asset_utils.shutil.which", return_value=None),
                patch(
                    "pptx2markdown.main_converter.asset_utils.subprocess.run",
                    side_effect=subprocess.TimeoutExpired("soffice", 60),
                ),
                self.assertLogs("pptx2markdown.main_converter.asset_utils", level="WARNING"),
            ):
                converted = convert_vector_assets_to_png([source], media_dir)

            self.assertEqual(converted, {})
            self.assertTrue(source.is_file())

    def test_converted_vector_links_are_rewritten_after_batch(self) -> None:
        output_dir = Path("/output/deck")
        source = output_dir / "media" / "image1.wmf"
        converted = output_dir / "media" / "image1.png"
        document = PresentationDocument(
            source=SourceDocument(name="deck.pptx", format="pptx"),
            slides=[
                SlideDocument(
                    page=1,
                    blocks=[ContentBlock(kind="image", content="![image](media/image1.wmf)")],
                )
            ],
        )

        rewritten = _rewrite_converted_vector_links(
            document,
            output_dir,
            {source: converted},
        )

        self.assertEqual(rewritten.slides[0].blocks[0].content, "![image](media/image1.png)")

    def test_markdown_asset_paths_always_use_uri_separators(self) -> None:
        with patch(
            "pptx2markdown.main_converter.run_pptx_to_markdown.os.path.relpath",
            return_value=r"media\image1.png",
        ):
            relative = relativize_markdown_path("ignored", Path("output"))

        self.assertEqual(relative, "media/image1.png")

    def test_text_normalization_preserves_all_unicode_scripts(self) -> None:
        self.assertEqual(
            normalize_text("网格系统 · 日本語 · Русский · Café"),
            "网格系统 日本語 русский café",
        )

    def test_bullet_none_overrides_paragraph_level(self) -> None:
        paragraph = ET.fromstring(
            f"""
            <a:p xmlns:a="{NS["a"]}">
              <a:pPr lvl="0"><a:buNone/></a:pPr>
              <a:r><a:t>7</a:t></a:r>
            </a:p>
            """
        )
        self.assertFalse(paragraph_has_list_semantics(paragraph))

    def test_plain_paragraph_between_lists_is_not_promoted(self) -> None:
        shape = ET.fromstring(
            f"""
            <p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">
              <p:txBody>
                <a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:t>First</a:t></a:r></a:p>
                <a:p><a:pPr lvl="0"><a:buNone/></a:pPr><a:r><a:t>Section</a:t></a:r></a:p>
                <a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:t>Second</a:t></a:r></a:p>
              </p:txBody>
            </p:sp>
            """
        )
        rendered = render_shape_blocks(extract_shape_blocks(shape))
        self.assertEqual(rendered, "- First\nSection\n- Second")

    def test_shape_level_list_summary_does_not_promote_mixed_text_box(self) -> None:
        shape = ET.fromstring(
            f"""
            <p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">
              <p:txBody>
                <a:p><a:r><a:t>Static</a:t></a:r></a:p>
                <a:p><a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr>
                  <a:r><a:t>Numbered</a:t></a:r></a:p>
                <a:p><a:r><a:t>Final</a:t></a:r></a:p>
              </p:txBody>
            </p:sp>
            """
        )
        blocks = extract_shape_blocks(shape)

        effective = _apply_effective_list_properties(
            blocks,
            {"has_list_semantics": True, "list_kind": "ol", "ph_type": None},
        )

        self.assertEqual(render_shape_blocks(effective), "Static\n1. Numbered\nFinal")

    def test_numeric_text_is_not_implicitly_a_slide_number(self) -> None:
        shape = ET.fromstring(
            f"""
            <p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">
              <p:txBody><a:p><a:r><a:t>42</a:t></a:r></a:p></p:txBody>
            </p:sp>
            """
        )
        self.assertFalse(has_slide_number_field(shape))

    def test_explicit_slide_number_field_is_detected(self) -> None:
        shape = ET.fromstring(
            f"""
            <p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">
              <p:txBody><a:p><a:fld type="slidenum"><a:t>42</a:t></a:fld></a:p></p:txBody>
            </p:sp>
            """
        )
        self.assertTrue(has_slide_number_field(shape))
        self.assertTrue(is_slide_number_only_shape(shape))

    def test_mixed_slide_number_field_does_not_hide_shape_text(self) -> None:
        shape = ET.fromstring(
            f"""
            <p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">
              <p:txBody><a:p>
                <a:r><a:t>Visible before</a:t></a:r>
                <a:fld type="slidenum"><a:t>42</a:t></a:fld>
                <a:r><a:t>Visible after</a:t></a:r>
              </a:p></p:txBody>
            </p:sp>
            """
        )

        self.assertTrue(has_slide_number_field(shape))
        self.assertFalse(is_slide_number_only_shape(shape))
        self.assertEqual(text_without_slide_number_fields(shape), "Visible before Visible after")
        self.assertEqual(
            render_shape_blocks(extract_shape_blocks(shape)),
            "Visible beforeVisible after",
        )

    def test_consecutive_breaks_preserve_heading_body_section_boundary(self) -> None:
        shape = ET.fromstring(
            f"""
            <p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">
              <p:txBody><a:p>
                <a:r><a:t>Title</a:t></a:r><a:br/><a:br/>
                <a:r><a:t>Body line</a:t></a:r>
              </a:p></p:txBody>
            </p:sp>
            """
        )
        self.assertEqual(render_shape_blocks(extract_shape_blocks(shape)), "Title\n\nBody line")

    def test_font_change_inside_word_does_not_insert_space(self) -> None:
        shape = ET.fromstring(
            f"""
            <p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">
              <p:txBody><a:p>
                <a:r><a:rPr sz="2400"/><a:t>inter</a:t></a:r>
                <a:r><a:rPr sz="1200"/><a:t>national</a:t></a:r>
              </a:p></p:txBody>
            </p:sp>
            """
        )
        blocks = extract_shape_blocks(shape)
        self.assertEqual(render_shape_blocks(blocks), "international")
        self.assertEqual(blocks[0].segments[0].font_pt, 24.0)

    def test_speaker_notes_extracts_body_and_ignores_slide_number(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "ppt"
            slides = root / "slides"
            notes = root / "notesSlides"
            (slides / "_rels").mkdir(parents=True)
            notes.mkdir(parents=True)
            slide_xml = slides / "slide1.xml"
            slide_xml.write_text("<p:sld xmlns:p='urn:p'/>", encoding="utf-8")
            (slides / "_rels" / "slide1.xml.rels").write_text(
                '<Relationships xmlns="http://schemas.openxmlformats.org/'
                'package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
                'officeDocument/2006/relationships/notesSlide" '
                'Target="../notesSlides/notesSlide1.xml"/>'
                "</Relationships>",
                encoding="utf-8",
            )
            (notes / "notesSlide1.xml").write_text(
                f"""<p:notes xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">
                  <p:cSld><p:spTree>
                    <p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
                      <p:txBody><a:p><a:r><a:t>Explain this slide</a:t></a:r></a:p></p:txBody>
                    </p:sp>
                    <p:sp><p:nvSpPr><p:nvPr><p:ph type="sldNum"/></p:nvPr></p:nvSpPr>
                      <p:txBody><a:p><a:r><a:t>1</a:t></a:r></a:p></p:txBody>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:notes>""",
                encoding="utf-8",
            )

            self.assertEqual(extract_speaker_notes(slide_xml), "Explain this slide")

    def test_speaker_notes_resolves_package_absolute_relationship_target(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            package = Path(tmp)
            (package / "[Content_Types].xml").write_text("<Types/>", encoding="utf-8")
            slides = package / "ppt" / "slides"
            notes = package / "ppt" / "notesSlides"
            (slides / "_rels").mkdir(parents=True)
            notes.mkdir(parents=True)
            slide_xml = slides / "slide1.xml"
            slide_xml.write_text("<p:sld xmlns:p='urn:p'/>", encoding="utf-8")
            (slides / "_rels" / "slide1.xml.rels").write_text(
                '<Relationships xmlns="http://schemas.openxmlformats.org/'
                'package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
                'officeDocument/2006/relationships/notesSlide" '
                'Target="/ppt/notesSlides/notesSlide1.xml"/>'
                "</Relationships>",
                encoding="utf-8",
            )
            (notes / "notesSlide1.xml").write_text(
                f"""<p:notes xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
                    <p:txBody><a:p><a:r><a:t>Absolute target note</a:t></a:r></a:p></p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:notes>""",
                encoding="utf-8",
            )

            self.assertEqual(extract_speaker_notes(slide_xml), "Absolute target note")

    def test_slide_show_false_is_hidden_metadata(self) -> None:
        hidden_slide = ET.fromstring('<p:sld xmlns:p="urn:p" show="0"/>')
        visible_slide = ET.fromstring('<p:sld xmlns:p="urn:p"/>')
        self.assertTrue(_is_hidden_slide(hidden_slide))
        self.assertFalse(_is_hidden_slide(visible_slide))

    def test_unpositioned_literal_date_placeholder_gets_origin_anchor(self) -> None:
        shape = ET.fromstring(
            f"""
            <p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">
              <p:spPr/>
              <p:txBody><a:p><a:r><a:t>Visible malformed date</a:t></a:r></a:p></p:txBody>
            </p:sp>
            """
        )
        self.assertTrue(_is_unpositioned_literal_date_placeholder(shape, "dt", None))
        self.assertFalse(_is_unpositioned_literal_date_placeholder(shape, "dt", (0, 0, 1, 1)))
        self.assertTrue(_is_synthetic_literal_date_anchor("dt", (0, 0, 1, 1), "default"))
        self.assertFalse(_is_synthetic_literal_date_anchor("dt", (0, 0, 1, 1), "direct"))

    def test_ole_graphic_frame_is_identified_before_table_fallback(self) -> None:
        frame = ET.fromstring(
            f"""
            <p:graphicFrame xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">
              <a:graphic>
                <a:graphicData
                  uri="http://schemas.openxmlformats.org/presentationml/2006/ole">
                <p:oleObj/>
                </a:graphicData>
              </a:graphic>
            </p:graphicFrame>
            """
        )
        self.assertEqual(graphic_frame_kind(frame), "ole")

    def test_model3d_alternate_content_keeps_preview_and_source_model(self) -> None:
        alternate = ET.fromstring(
            f"""
            <p:spTree xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}"
                      xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
              <mc:AlternateContent>
                <mc:Choice Requires="am3d">
                  <p:graphicFrame>
                    <a:graphic><a:graphicData
                      uri="http://schemas.microsoft.com/office/drawing/2017/model3d"/>
                    </a:graphic>
                  </p:graphicFrame>
                </mc:Choice>
                <mc:Fallback><p:pic/></mc:Fallback>
              </mc:AlternateContent>
            </p:spTree>
            """
        )

        children = list(_expanded_children(alternate))

        self.assertEqual(
            [element.tag.rsplit("}", 1)[-1] for element in children],
            ["pic", "graphicFrame"],
        )
        self.assertEqual(graphic_frame_kind(children[1]), "model3d")

    def test_external_run_hyperlink_is_rendered_and_unsafe_scheme_is_plain_text(self) -> None:
        shape = ET.fromstring(
            f"""
            <p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}" xmlns:r="{NS["r"]}">
              <p:txBody>
                <a:p>
                  <a:r><a:rPr><a:hlinkClick r:id="rId1"/></a:rPr><a:t>Open docs</a:t></a:r>
                </a:p>
                <a:p>
                  <a:r><a:rPr><a:hlinkClick r:id="rId2"/></a:rPr><a:t>Unsafe</a:t></a:r>
                </a:p>
              </p:txBody>
            </p:sp>
            """
        )
        blocks = extract_shape_blocks(
            shape,
            {"rId1": "https://example.com/docs", "rId2": "javascript:alert(1)"},
        )

        self.assertEqual(
            render_shape_blocks(blocks),
            "[Open docs](https://example.com/docs)\nUnsafe",
        )
        self.assertEqual([block.plain_text for block in blocks], ["Open docs", "Unsafe"])

    def test_external_hyperlink_inside_table_cell_is_preserved(self) -> None:
        table = ET.fromstring(
            f"""
            <a:tbl xmlns:a="{NS["a"]}" xmlns:r="{NS["r"]}">
              <a:tblGrid><a:gridCol w="100"/></a:tblGrid>
              <a:tr h="100"><a:tc><a:txBody><a:p>
                <a:r><a:rPr><a:hlinkClick r:id="rId1"/></a:rPr>
                  <a:t>Linked cell</a:t></a:r>
              </a:p></a:txBody></a:tc></a:tr>
            </a:tbl>
            """
        )
        parsed = parse_table_element(table)

        inject_table_run_hyperlinks(
            parsed,
            table,
            {"rId1": "https://example.com/cell"},
            ns=NS,
        )
        markdown = render_parsed_table_to_markdown(parsed)

        self.assertIn("[Linked cell](https://example.com/cell)", markdown)

    def test_merged_table_body_keeps_content_only_at_merge_origin(self) -> None:
        table = ET.fromstring(
            f"""
            <a:tbl xmlns:a="{NS["a"]}">
              <a:tblGrid><a:gridCol w="100"/><a:gridCol w="100"/></a:tblGrid>
              <a:tr h="100">
                <a:tc gridSpan="2"><a:txBody><a:p><a:r><a:t>Header</a:t></a:r></a:p>
                </a:txBody></a:tc>
                <a:tc hMerge="1"><a:txBody><a:p/></a:txBody></a:tc>
              </a:tr>
              <a:tr h="100">
                <a:tc gridSpan="2"><a:txBody><a:p><a:r><a:t>Body</a:t></a:r></a:p>
                </a:txBody></a:tc>
                <a:tc hMerge="1"><a:txBody><a:p/></a:txBody></a:tc>
              </a:tr>
            </a:tbl>
            """
        )

        markdown = render_parsed_table_to_markdown(parse_table_element(table))

        self.assertEqual(markdown, "| Header | Header |\n|---|---|\n| Body |  |\n")

    def test_table_cell_preserves_paragraphs_and_nested_list_semantics(self) -> None:
        table = ET.fromstring(
            f"""
            <a:tbl xmlns:a="{NS["a"]}">
              <a:tblGrid><a:gridCol w="100"/></a:tblGrid>
              <a:tr h="100"><a:tc><a:txBody><a:lstStyle/>
                <a:p><a:pPr><a:buChar char="•"/></a:pPr>
                  <a:r><a:t>Parent</a:t></a:r></a:p>
                <a:p><a:pPr lvl="1"><a:buChar char="•"/></a:pPr>
                  <a:r><a:t>Child</a:t></a:r></a:p>
                <a:p><a:pPr><a:buNone/></a:pPr>
                  <a:r><a:t>Plain</a:t></a:r></a:p>
              </a:txBody></a:tc></a:tr>
            </a:tbl>
            """
        )

        markdown = render_parsed_table_to_markdown(parse_table_element(table))

        self.assertIn("| - Parent<br>  - Child<br>Plain |", markdown)

    def test_materialized_shape_keeps_sidecar_order_and_bbox_after_id_remap(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            slide_xml = root / "slide1.reordered.xml"
            slide_xml.write_text("<slide/>", encoding="utf-8")
            sidecar = root / "slide1.structure_analysis.json"
            sidecar.write_text(
                """{
                  "structure_order": [
                    {"shape_id":"1","xml_index":1,"bbox":[0,0,100,100]},
                    {"shape_id":"layout:7","xml_index":1000007,
                     "inheritance_kind":"materialized","source_part":"layout",
                     "bbox":[0,110,100,210]},
                    {"shape_id":"2","xml_index":2,"bbox":[0,220,100,320]}
                  ]
                }""",
                encoding="utf-8",
            )
            context = SlideConversionContext(
                slide_xml=slide_xml,
                page_no=1,
                ns=NS,
                heading_hints=load_heading_hints(slide_xml),
                effective_properties=load_effective_properties(slide_xml),
            )
            elem = ET.Element("shape")
            items = [
                FlattenedShape(elem, "sp", "1", "", (), (1,), (0, 0, 100, 100)),
                FlattenedShape(elem, "sp", "1000007", "", (), (2,), None),
                FlattenedShape(elem, "sp", "2", "", (), (3,), (0, 220, 100, 320)),
            ]

            enriched = _with_effective_bboxes(items, context)
            ordered = _ordered_flattened_shapes(enriched, context)

        self.assertEqual([item.shape_id for item in ordered], ["1", "1000007", "2"])
        self.assertEqual(ordered[1].bbox, (0, 110, 100, 210))

    def test_partial_sidecar_match_preserves_rewritten_xml_order(self) -> None:
        elem = ET.Element("shape")
        context = SlideConversionContext(
            slide_xml=Path("slide.xml"),
            page_no=1,
            ns=NS,
            heading_hints={
                "1": {"order_index": 0},
                "2": {"order_index": 1},
            },
        )
        items = [
            FlattenedShape(elem, "sp", "1", "", (), (1,), None),
            FlattenedShape(elem, "sp", "inherited", "", (), (2,), None),
            FlattenedShape(elem, "sp", "2", "", (), (3,), None),
        ]

        self.assertEqual(_ordered_flattened_shapes(items, context), items)

    def test_relationship_target_recovers_part_inside_reordered_path(self) -> None:
        rels_path = Path("work/structure/sample/_rels/slide9.reordered.xml.rels")
        target = "../../target_slides/sample/ppt/charts/chart1.xml"

        self.assertEqual(
            _ooxml_part_from_relationship(rels_path, target),
            "ppt/charts/chart1.xml",
        )

    def test_master_only_text_is_inserted_into_empty_slide(self) -> None:
        master_xml = f"""
        <p:sldMaster xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">
          <p:cSld>
            <p:spTree>
              <p:nvGrpSpPr/>
              <p:grpSpPr/>
              <p:sp>
                <p:nvSpPr><p:cNvPr id="7" name="Master text"/></p:nvSpPr>
                <p:spPr/>
                <p:txBody><a:p><a:r><a:t>Inherited text</a:t></a:r></a:p></p:txBody>
              </p:sp>
            </p:spTree>
          </p:cSld>
        </p:sldMaster>
        """

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source_slide = root / "slide.xml"
            source_master = root / "master.xml"
            source_slide.write_text(ET.tostring(_empty_slide_tree().getroot(), encoding="unicode"))
            source_master.write_text(master_xml)
            tree = _empty_slide_tree()

            materialize_tree_by_objects(
                tree,
                [_slide_object()],
                {"master_xml": str(source_master)},
                root,
                root / "_rels" / "slide.reordered.xml.rels",
                source_slide,
            )

        texts = [node.text for node in tree.findall(".//a:t", NS)]
        self.assertEqual(texts, ["Inherited text"])

    def test_thin_master_picture_is_filtered_as_decoration(self) -> None:
        thin_picture = ET.fromstring(
            f"""
            <p:pic xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">
              <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10000" cy="100"/></a:xfrm></p:spPr>
            </p:pic>
            """
        )
        regular_picture = ET.fromstring(
            f"""
            <p:pic xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">
              <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10000" cy="5000"/></a:xfrm></p:spPr>
            </p:pic>
            """
        )

        self.assertFalse(_is_visible_materialized_shape(thin_picture, "pic", None))
        self.assertTrue(_is_visible_materialized_shape(regular_picture, "pic", None))

    def test_layout_placeholder_prompts_are_not_materialized(self) -> None:
        self.assertTrue(_is_placeholder_prompt_text("‹#›"))
        self.assertTrue(_is_placeholder_prompt_text("Click to edit Master title style"))
        self.assertTrue(_is_placeholder_prompt_text("<date/time>"))
        self.assertFalse(_is_placeholder_prompt_text("Q3 < Q4"))

    def test_zero_extent_picture_is_not_emitted_as_visible_content(self) -> None:
        picture = ET.fromstring(
            f"""
            <p:pic xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">
              <p:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>
              </p:spPr>
            </p:pic>
            """
        )
        self.assertTrue(is_decorative("pic", "", picture))

    def test_image_filled_autoshape_is_visible_content(self) -> None:
        shape = ET.fromstring(
            f"""
            <p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}" xmlns:r="{NS["r"]}">
              <p:spPr>
                <a:xfrm><a:off x="10" y="20"/><a:ext cx="100" cy="200"/></a:xfrm>
                <a:blipFill><a:blip r:embed="rId1"/></a:blipFill>
              </p:spPr>
            </p:sp>
            """
        )

        self.assertFalse(is_decorative("sp", "", shape))
        self.assertFalse(_is_visible_materialized_shape(shape, "sp", None))

    def test_zero_height_autofit_text_keeps_anchor_bbox(self) -> None:
        shape = ET.fromstring(
            f"""
            <p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">
              <p:spPr>
                <a:xfrm><a:off x="10" y="20"/><a:ext cx="100" cy="0"/></a:xfrm>
              </p:spPr>
              <p:txBody><a:bodyPr><a:spAutoFit/></a:bodyPr></p:txBody>
            </p:sp>
            """
        )
        self.assertEqual(extract_bbox_emu(shape), (10, 20, 110, 21))

    def test_numbered_heading_accepts_space_before_period(self) -> None:
        self.assertTrue(is_numbered_heading_text("3 . Text Box Heading"))
        self.assertTrue(is_numbered_heading_text("2 . 1 Nested Heading"))
        self.assertFalse(is_numbered_heading_text("3 items in a sentence"))


if __name__ == "__main__":
    unittest.main()
