from __future__ import annotations

import unittest

from pptx2markdown.model import (
    AssetDescriptor,
    HeadingBlock,
    ImageBlock,
    ListBlock,
    ListItem,
    PresentationDocument,
    SlideDocument,
    SourceDocument,
    TableBlock,
    TableCell,
    TextBlock,
)
from pptx2markdown.recipe import Recipe, apply_recipe
from pptx2markdown.review import review
from pptx2markdown.tables import rows_markdown

ASSET = "asset-" + "b" * 16


def _table(block_id: str, rows, caption_rows: int = 0) -> TableBlock:
    return TableBlock(
        id=block_id,
        rows=rows,
        caption_rows=caption_rows,
        markdown=rows_markdown(rows, caption_rows=caption_rows),
    )


C = TableCell
GROUPED = (
    (C(text="Item", row_span=2), C(text="2024", col_span=2), C(merged_from=(0, 1))),
    (C(merged_from=(0, 0)), C(text="H1"), C(text="H2")),
    (C(text="Sales"), C(text="10"), C(text="12")),
)
NUMERIC = ((C(text="1"), C(text="2")), (C(text="3"), C(text="4")))
TITLED = (
    (C(text="Summary", col_span=2), C(merged_from=(0, 0))),
    (C(text="A"), C(text="B")),
    (C(text="1"), C(text="2")),
)


def _document() -> PresentationDocument:
    return PresentationDocument(
        source=SourceDocument(name="deck.pptx", sha256="0" * 64),
        assets=(
            AssetDescriptor(
                id=ASSET,
                kind="image",
                original_name="a.png",
                media_type="image/png",
                size=1,
                sha256="1" * 64,
                source_part="ppt/media/a.png",
            ),
        ),
        slides=(
            SlideDocument(
                page=1,
                blocks=(
                    HeadingBlock(id="s1.b1", text="Results", level=1),
                    TextBlock(id="s1.b2", text="**Key points**", shape_id="5"),
                    ListBlock(id="s1.b3", items=(ListItem(text="Up"),), shape_id="5"),
                    TextBlock(id="s1.b4", text="**Standalone bold**", shape_id="6"),
                    TextBlock(id="s1.b5", text="Other shape", shape_id="7"),
                    _table("s1.b6", GROUPED),
                    _table("s1.b7", NUMERIC),
                    _table("s1.b8", TITLED, caption_rows=1),
                    ImageBlock(id="s1.b9", asset_id=ASSET, alt_text="Picture 4"),
                    ImageBlock(id="s1.b10", asset_id=ASSET, alt_text="Org chart"),
                ),
            ),
        ),
    )


class ReviewTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.items = {(item.target, item.topic): item for item in review(_document())}

    def test_only_uncertain_decisions_are_asked(self) -> None:
        self.assertEqual(
            sorted(self.items),
            [
                ("s1.b2", "subheading"),
                ("s1.b6", "table-header-rows"),
                ("s1.b7", "table-header-rows"),
                ("s1.b8", "table-caption"),
                ("s1.b9", "image-description"),
            ],
        )

    def test_grouped_header_suggests_two_header_rows(self) -> None:
        item = self.items[("s1.b6", "table-header-rows")]
        self.assertTrue(item.evidence["header_spans_rows"])
        self.assertEqual(item.options[item.suggested].rule.header_rows, 2)

    def test_numeric_first_row_suggests_no_header(self) -> None:
        item = self.items[("s1.b7", "table-header-rows")]
        self.assertEqual(item.options[item.suggested].rule.header_rows, 0)

    def test_suggested_options_form_a_recipe(self) -> None:
        rules = [
            item.options[item.suggested].rule
            for item in self.items.values()
            if item.options[item.suggested].rule is not None
        ]
        result = apply_recipe(_document(), Recipe(rules=tuple(rules)))
        blocks = {block.id: block for block in result.slides[0].blocks}
        self.assertEqual(
            blocks["s1.b2"], HeadingBlock(id="s1.b2", text="Key points", level=2, shape_id="5")
        )
        self.assertEqual(
            blocks["s1.b6"].markdown.splitlines()[0], "| Item | 2024 / H1 | 2024 / H2 |"
        )
        self.assertEqual(blocks["s1.b7"].markdown.splitlines()[0], "| col_1 | col_2 |")


if __name__ == "__main__":
    unittest.main()


class CliReviewTests(unittest.TestCase):
    def test_review_prints_questions_with_rules_and_writes_nothing(self) -> None:
        import io
        import json
        import tempfile
        from contextlib import redirect_stdout
        from pathlib import Path

        from pptx2markdown import cli

        from .pptx_factory import build_test_pptx

        with tempfile.TemporaryDirectory() as directory:
            source = build_test_pptx(Path(directory) / "deck.pptx")
            output = io.StringIO()
            with redirect_stdout(output):
                code = cli.main([str(source), "--review", "-o", str(Path(directory) / "out")])
            self.assertEqual(code, 0)
            self.assertFalse((Path(directory) / "out").exists())
        report = json.loads(output.getvalue())
        self.assertEqual(report[0]["source"], "deck.pptx")
        for item in report[0]["items"]:
            self.assertIn(item["suggested"], range(len(item["options"])))
