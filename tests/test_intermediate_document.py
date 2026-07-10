from __future__ import annotations

import unittest

from pptx2markdown.main_converter.converter_models import (
    ContentBlock,
    PresentationDocument,
    SlideDocument,
    render_presentation_markdown,
)


class IntermediateDocumentTests(unittest.TestCase):
    def test_json_round_trip_preserves_markdown_rendering(self) -> None:
        document = PresentationDocument(
            source="deck.pptx",
            reading_order="xycut",
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
