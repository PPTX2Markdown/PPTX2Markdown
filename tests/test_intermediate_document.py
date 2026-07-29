from __future__ import annotations

import inspect
import json
import unittest
from pathlib import Path
from unittest.mock import patch

from pydantic import ValidationError

import pptx2markdown
from pptx2markdown.main_converter.converter_models import (
    BoundingBox,
    ContentBlock,
    PresentationDocument,
    SlideDocument,
    SourceDocument,
    render_presentation_markdown,
)

SCHEMA_PATH = (
    Path(__file__).parents[1] / "src" / "pptx2markdown" / "presentation_document.schema.json"
)


class IntermediateDocumentTests(unittest.TestCase):
    def test_public_convert_api_passes_vector_conversion_option(self) -> None:
        with (
            patch(
                "pptx2markdown.main_converter.run_pptx_to_markdown._build_config",
                side_effect=lambda args: args,
            ),
            patch("pptx2markdown.main_converter.run_pptx_to_markdown._configure_logging"),
            patch(
                "pptx2markdown.main_converter.run_pptx_to_markdown.run",
                return_value=0,
            ) as run,
        ):
            status = pptx2markdown.convert([], convert_vector_images=True)

        self.assertEqual(status, 0)
        self.assertTrue(run.call_args.args[0].convert_vector_images)

    def test_public_convert_api_exposes_static_pipeline_only(self) -> None:
        parameters = inspect.signature(pptx2markdown.convert).parameters

        self.assertIn("convert_vector_images", parameters)

        self.assertNotIn("reading_order", parameters)
        self.assertNotIn("reuse_surya_cache", parameters)
        self.assertNotIn("strict_headings", parameters)
        self.assertNotIn("image_vlm_provider", parameters)
        self.assertNotIn("image_vlm_model", parameters)
        self.assertNotIn("ignore_image_vlm_cache", parameters)

        with self.assertRaisesRegex(ValueError, "headings must be"):
            pptx2markdown.convert([], headings="surya")
        with self.assertRaisesRegex(ValueError, "placeholder_inheritance must be"):
            pptx2markdown.convert([], placeholder_inheritance="semantic")
        with self.assertRaisesRegex(ValueError, "inherited_shapes must be"):
            pptx2markdown.convert([], inherited_shapes="semantic")

    def test_json_round_trip_preserves_markdown_rendering(self) -> None:
        document = PresentationDocument(
            source=SourceDocument(name="deck.pptx", format="pptx"),
            slides=[
                SlideDocument(
                    page=1,
                    blocks=[
                        ContentBlock(
                            kind="heading",
                            content="Title",
                            shape_id="2",
                            heading_level=1,
                        ),
                        ContentBlock(kind="text", content="Body", shape_id="3"),
                    ],
                ),
                SlideDocument(
                    page=2,
                    hidden=True,
                    notes="Explain the table.",
                    blocks=[ContentBlock(kind="table", content="| A |\n| - |")],
                ),
            ],
        )

        restored = PresentationDocument.model_validate_json(document.model_dump_json())

        self.assertEqual(
            render_presentation_markdown(restored),
            (
                "[Page_1]\n\n# Title\n\nBody\n\n[Page_2]\n\n"
                "<!-- hidden: true -->\n\n| A |\n| - |\n\n"
                "[Speaker_Notes]\n\nExplain the table.\n"
            ),
        )

    def test_schema_contract_rejects_invalid_documents(self) -> None:
        with self.assertRaisesRegex(ValidationError, "source name must be a basename"):
            SourceDocument(name="/tmp/deck.pptx", format="pptx")
        with self.assertRaisesRegex(ValidationError, "heading blocks require heading_level"):
            ContentBlock(kind="heading", content="Title")
        with self.assertRaisesRegex(ValidationError, "only valid for heading blocks"):
            ContentBlock(kind="text", content="Body", heading_level=2)
        with self.assertRaisesRegex(ValidationError, "strictly increasing"):
            PresentationDocument(
                source=SourceDocument(name="deck.pptx", format="pptx"),
                slides=[SlideDocument(page=2), SlideDocument(page=1)],
            )

    def test_schema_contract_includes_geometry_and_provenance(self) -> None:
        block = ContentBlock(
            kind="text",
            content="Inherited body",
            shape_id="layout:7",
            bbox=BoundingBox(x=100, y=200, width=300, height=400),
            source_part="layout",
        )

        self.assertEqual(block.bbox.unit, "emu")
        self.assertEqual(block.source_part, "layout")

        hidden_slide = SlideDocument(page=1, hidden=True)
        self.assertTrue(hidden_slide.hidden)

    def test_checked_in_json_schema_matches_pydantic_model(self) -> None:
        checked_in = json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))
        self.assertEqual(checked_in, PresentationDocument.model_json_schema())


if __name__ == "__main__":
    unittest.main()
