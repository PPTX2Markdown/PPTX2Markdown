from __future__ import annotations

import inspect
import unittest

import pptx2markdown
from pptx2markdown.main_converter.converter_models import (
    ContentBlock,
    PresentationDocument,
    SlideDocument,
    render_presentation_markdown,
)


class IntermediateDocumentTests(unittest.TestCase):
    def test_public_convert_api_exposes_static_pipeline_only(self) -> None:
        parameters = inspect.signature(pptx2markdown.convert).parameters

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
            source="deck.pptx",
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
                    blocks=[ContentBlock(kind="table", content="| A |\n| - |")],
                ),
            ],
        )

        restored = PresentationDocument.model_validate_json(document.model_dump_json())

        self.assertEqual(
            render_presentation_markdown(restored),
            "[Page_1]\n\n# Title\n\nBody\n\n[Page_2]\n\n| A |\n| - |\n",
        )


if __name__ == "__main__":
    unittest.main()
