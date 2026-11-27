from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from pydantic import ValidationError

from pptx2markdown import cli
from pptx2markdown.chunks import slide_chunks
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
from pptx2markdown.recipe import Recipe, Rule, Selector, apply_recipe, explain_recipe
from pptx2markdown.render import render_presentation_markdown
from pptx2markdown.tables import rows_markdown

from .pptx_factory import build_test_pptx

ASSET = "asset-" + "a" * 16
ROWS = (
    (TableCell(text="Region"), TableCell(text="Sales")),
    (TableCell(text="Seoul", row_span=2), TableCell(text="10")),
    (TableCell(merged_from=(1, 0)), TableCell(text="20")),
)


def _document() -> PresentationDocument:
    table = TableBlock(id="s2.b2", rows=ROWS, markdown=rows_markdown(ROWS), shape_id="7")
    return PresentationDocument(
        source=SourceDocument(name="deck.pptx", sha256="0" * 64),
        assets=(
            AssetDescriptor(
                id=ASSET,
                kind="image",
                original_name="chart.png",
                media_type="image/png",
                size=1,
                sha256="1" * 64,
                source_part="ppt/media/chart.png",
            ),
        ),
        slides=(
            SlideDocument(
                page=1,
                blocks=(
                    HeadingBlock(id="s1.b1", text="Quarterly review", level=1),
                    TextBlock(id="s1.b2", text="Prepared by finance"),
                ),
                notes="Welcome everyone",
            ),
            SlideDocument(
                page=2,
                blocks=(
                    HeadingBlock(id="s2.b1", text="Sales by region", level=2),
                    table,
                    ImageBlock(id="s2.b3", asset_id=ASSET, alt_text="Sales chart"),
                    ListBlock(
                        id="s2.b4", items=(ListItem(text="Up 10%"), ListItem(text="Seoul leads"))
                    ),
                ),
                notes="Mention the Seoul office",
            ),
            SlideDocument(
                page=3,
                hidden=True,
                blocks=(TextBlock(id="s3.b1", text="Backup data"),),
            ),
        ),
    )


class SelectorTests(unittest.TestCase):
    def test_selector_syntax(self) -> None:
        self.assertEqual(Selector.parse("s2.b3").block_id, "s2.b3")
        self.assertEqual(Selector.parse("slide[2-]").first_page, 2)
        self.assertIsNone(Selector.parse("slide[2-]").last_page)
        self.assertTrue(Selector.parse("slide[hidden]").slides_only)
        selector = Selector.parse("slide[2]/table[-1]")
        self.assertEqual((selector.first_page, selector.kind, selector.index), (2, "table", -1))

    def test_invalid_selectors_are_rejected_when_the_rule_is_built(self) -> None:
        for target in ("slide[0]", "slide[5-2]", "chart[x]", "tables", "slide[2]/foo", "a/b"):
            with self.subTest(target=target), self.assertRaises(ValidationError):
                Rule(target=target, exclude=True)

    def test_explain_lists_what_each_rule_selects(self) -> None:
        recipe = Recipe(
            rules=(
                Rule(target="heading", heading_level=2),
                Rule(target="slide[hidden]", exclude=True),
                Rule(target="slide[9]/table", exclude=True),
            )
        )
        matches = explain_recipe(_document(), recipe)
        self.assertEqual(matches[0].block_ids, ("s1.b1", "s2.b1"))
        self.assertEqual(matches[1].pages, (3,))
        self.assertEqual(matches[2].block_ids, ())


class ApplyRecipeTests(unittest.TestCase):
    def test_empty_recipe_changes_nothing(self) -> None:
        document = _document()
        self.assertEqual(apply_recipe(document, Recipe()), document)

    def test_document_settings(self) -> None:
        result = apply_recipe(
            _document(), Recipe(include_notes=False, include_hidden=False, merged_cells="blank")
        )
        self.assertEqual([slide.page for slide in result.slides], [1, 2])
        self.assertTrue(all(slide.notes is None for slide in result.slides))
        table = result.slides[1].blocks[1]
        self.assertEqual(table.merged_cells, "blank")
        self.assertEqual(table.markdown.splitlines()[-1], "|  | 20 |")

    def test_rules_target_blocks_and_slides(self) -> None:
        recipe = Recipe(
            rules=(
                Rule(target="slide[1]", notes=False),
                Rule(target="s1.b2", exclude=True),
                Rule(target="table[0]", header_rows=2, merged_cells="fill", caption_rows=0),
                Rule(target="image", description="Bar chart of sales by region"),
                Rule(target="slide[2]/list", as_text=True),
                Rule(target="slide[3]", exclude=True),
            )
        )
        result = apply_recipe(_document(), recipe)
        first, second = result.slides
        self.assertIsNone(first.notes)
        self.assertEqual([block.id for block in first.blocks], ["s1.b1"])
        self.assertEqual(second.notes, "Mention the Seoul office")
        table = second.blocks[1]
        self.assertEqual(table.header_rows, 2)
        self.assertEqual(table.markdown.splitlines()[0], "| Region / Seoul | Sales / 10 |")
        self.assertEqual(second.blocks[2].description, "Bar chart of sales by region")
        self.assertEqual(second.blocks[3], TextBlock(id="s2.b4", text="Up 10%\nSeoul leads"))

    def test_later_rules_win_and_selectors_see_the_parsed_slide(self) -> None:
        recipe = Recipe(
            rules=(
                Rule(target="s1.b2", heading_level=3),
                Rule(target="heading", heading_level=4),
                Rule(target="text", exclude=True),
            )
        )
        result = apply_recipe(_document(), recipe)
        # s1.b2 was text when the recipe started, so "text" still removes it and
        # "heading" does not reach it.
        self.assertEqual([block.id for block in result.slides[0].blocks], ["s1.b1"])
        self.assertEqual(result.slides[0].blocks[0].level, 4)

    def test_text_becomes_a_heading_and_ids_stay_stable(self) -> None:
        result = apply_recipe(
            _document(),
            Recipe(
                rules=(Rule(target="s2.b1", exclude=True), Rule(target="s1.b2", heading_level=2))
            ),
        )
        self.assertEqual(
            result.slides[0].blocks[1],
            HeadingBlock(id="s1.b2", text="Prepared by finance", level=2),
        )
        self.assertEqual(
            [block.id for block in result.slides[1].blocks], ["s2.b2", "s2.b3", "s2.b4"]
        )

    def test_recipe_round_trips_through_json(self) -> None:
        recipe = Recipe(rules=(Rule(target="table", merged_cells="blank", reason="audit view"),))
        self.assertEqual(Recipe.model_validate_json(recipe.model_dump_json()), recipe)
        self.assertIn("rules", Recipe.model_json_schema()["properties"])


class AnnotatedMarkdownTests(unittest.TestCase):
    def test_block_ids_are_visible_for_review(self) -> None:
        markdown = render_presentation_markdown(_document(), asset_mode="omit", block_ids=True)
        self.assertIn("<!-- s2.b2 table -->\n| Region | Sales |", markdown)
        plain = render_presentation_markdown(_document(), asset_mode="omit")
        self.assertNotIn("<!-- s", plain)


class ChunkTests(unittest.TestCase):
    def test_one_chunk_per_slide_with_citation_metadata(self) -> None:
        chunks = slide_chunks(_document())
        self.assertEqual([chunk.metadata["page"] for chunk in chunks], [1, 2, 3])
        second = chunks[1]
        self.assertTrue(
            second.text.startswith(
                "<!-- deck.pptx, slide 2: Quarterly review — Sales by region -->\n\n"
            )
        )
        self.assertIn("Speaker notes: Mention the Seoul office", second.text)
        self.assertNotIn("![", second.text)
        self.assertIn("[image: Sales chart]", second.text)
        self.assertEqual(second.metadata["title"], "Sales by region")
        self.assertEqual(second.metadata["block_ids"], ["s2.b1", "s2.b2", "s2.b3", "s2.b4"])
        self.assertTrue(chunks[2].metadata["hidden"])

    def test_long_slides_split_at_block_boundaries(self) -> None:
        chunks = slide_chunks(_document(), max_chars=60, include_notes=False, context=False)
        second = [chunk for chunk in chunks if chunk.metadata["page"] == 2]
        self.assertGreater(len(second), 1)
        self.assertTrue(all(chunk.metadata["parts"] == len(second) for chunk in second))
        self.assertEqual(
            [block for chunk in second for block in chunk.metadata["block_ids"]],
            ["s2.b1", "s2.b2", "s2.b3", "s2.b4"],
        )

    def test_default_picture_names_are_not_content(self) -> None:
        document = _document()
        slide = document.slides[1]
        image = slide.blocks[2].model_copy(update={"alt_text": "Picture 3"})
        slide = slide.model_copy(update={"blocks": (*slide.blocks[:2], image, slide.blocks[3])})
        chunks = slide_chunks(document.model_copy(update={"slides": (slide,)}))
        self.assertNotIn("[image:", chunks[0].text)
        self.assertNotIn("s2.b3", chunks[0].metadata["block_ids"])

    def test_comment_cannot_be_closed_by_a_title(self) -> None:
        document = _document()
        slide = document.slides[0].model_copy(
            update={"blocks": (HeadingBlock(id="s1.b1", text="a --> b", level=1),)}
        )
        chunks = slide_chunks(document.model_copy(update={"slides": (slide,)}))
        self.assertEqual(chunks[0].text.splitlines()[0].count("-->"), 1)


class CliRecipeTests(unittest.TestCase):
    def test_cli_applies_a_recipe_file(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = build_test_pptx(root / "deck.pptx")
            recipe = root / "recipe.json"
            recipe.write_text(
                json.dumps(
                    {"include_notes": False, "rules": [{"target": "image", "exclude": True}]}
                ),
                encoding="utf-8",
            )
            code = cli.main(
                [str(source), "-o", str(root / "out"), "--recipe", str(recipe), "--no-assets"]
            )
            self.assertEqual(code, 0)
            markdown = (root / "out" / "deck" / "deck.md").read_text(encoding="utf-8")
            self.assertNotIn("[Speaker_Notes]", markdown)
            self.assertNotIn("![", markdown)


if __name__ == "__main__":
    unittest.main()


class HeadingPolicyTests(unittest.TestCase):
    def _document(self) -> PresentationDocument:
        return PresentationDocument(
            source=SourceDocument(name="deck.pptx", sha256="0" * 64),
            slides=(
                SlideDocument(
                    page=1,
                    blocks=(
                        HeadingBlock(id="s1.b1", text="Title", level=1, source="placeholder"),
                        HeadingBlock(id="s1.b2", text="2. Results", level=2, source="numbering"),
                        HeadingBlock(id="s1.b3", text="Big claim", level=2, source="font"),
                    ),
                ),
            ),
        )

    def test_placeholder_policy_keeps_only_placeholder_headings(self) -> None:
        result = apply_recipe(self._document(), Recipe(headings="placeholder"))
        self.assertEqual(
            [block.kind for block in result.slides[0].blocks], ["heading", "text", "text"]
        )
        self.assertEqual(result.slides[0].blocks[1], TextBlock(id="s1.b2", text="2. Results"))

    def test_font_only_heading_is_a_review_question(self) -> None:
        from pptx2markdown.review import review

        items = [item for item in review(self._document()) if item.topic == "heading"]
        self.assertEqual([item.target for item in items], ["s1.b3"])
        self.assertTrue(items[0].options[1].rule.as_text)


class CliHeadingsTests(unittest.TestCase):
    def test_strict_headings_keep_only_placeholder_headings(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = build_test_pptx(root / "deck.pptx")
            code = cli.main([str(source), "-o", str(root / "out"), "--headings", "strict"])
            self.assertEqual(code, 0)
            markdown = (root / "out" / "deck" / "deck.md").read_text(encoding="utf-8")
        self.assertIn("# Quarterly Report", markdown)
