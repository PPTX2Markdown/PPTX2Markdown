"""Slide text that looks like HTML is written so it renders as text."""

from __future__ import annotations

import unittest

from pptx2markdown.main_converter.converter_models import (
    ContentBlock,
    SlideDocument,
    render_slide_markdown,
)


class RawHtmlTests(unittest.TestCase):
    def test_raw_html_in_slide_text_is_escaped(self) -> None:
        slide = SlideDocument(
            page=1,
            blocks=[
                ContentBlock(kind="heading", content="<b>Plan</b>", heading_level=1),
                ContentBlock(kind="text", content="<script>alert(1)</script> and $x<y$"),
                ContentBlock(kind="table", content="| a<br>b | <img src=x> |\n|---|---|"),
                ContentBlock(kind="math", content="$$\na<b\n$$"),
            ],
            notes="<iframe>",
        )
        markdown = render_slide_markdown(slide)
        self.assertIn("# \\<b>Plan\\</b>", markdown)
        self.assertIn("\\<script>alert(1)\\</script> and $x<y$", markdown)
        self.assertIn("| a<br>b | \\<img src=x> |", markdown)
        self.assertIn("$$\na<b\n$$", markdown)
        self.assertIn("\\<iframe>", markdown)
        self.assertNotIn("<script>", markdown.replace("\\<script>", ""))


if __name__ == "__main__":
    unittest.main()
