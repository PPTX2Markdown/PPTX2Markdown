from __future__ import annotations

import asyncio
import importlib.util
import json
import tempfile
import unittest
from pathlib import Path

from pptx2markdown import mcp_server
from pptx2markdown.libreoffice import find_soffice

from .pptx_factory import build_test_pptx


class ToolFunctionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.directory = tempfile.TemporaryDirectory()
        cls.path = str(build_test_pptx(Path(cls.directory.name) / "deck.pptx"))

    @classmethod
    def tearDownClass(cls) -> None:
        cls.directory.cleanup()

    def test_convert_shows_block_ids_and_applies_a_recipe(self) -> None:
        markdown = mcp_server.convert_presentation(self.path)
        self.assertIn("<!-- s1.b1 heading -->", markdown)
        trimmed = mcp_server.convert_presentation(
            self.path, {"rules": [{"target": "s1.b1", "exclude": True}]}
        )
        self.assertNotIn("<!-- s1.b1 heading -->", trimmed)

    def test_check_recipe_reports_rules_that_select_nothing(self) -> None:
        matches = mcp_server.check_recipe(
            self.path,
            {
                "rules": [
                    {"target": "heading", "heading_level": 2},
                    {"target": "s9.b1", "exclude": True},
                ]
            },
        )
        self.assertEqual(matches[0]["block_ids"], ["s1.b1"])
        self.assertEqual(matches[1]["block_ids"], [])

    def test_invalid_input_is_a_clear_error(self) -> None:
        with self.assertRaisesRegex(ValueError, "invalid recipe"):
            mcp_server.check_recipe(self.path, {"rules": [{"target": "tables[x]"}]})
        with self.assertRaisesRegex(ValueError, "not a .pptx"):
            mcp_server.review_presentation(self.path + ".txt")
        with self.assertRaisesRegex(ValueError, "file not found"):
            mcp_server.review_presentation(self.path.replace("deck", "missing"))

    def test_review_answers_become_a_recipe(self) -> None:
        report = mcp_server.review_presentation(self.path)
        self.assertEqual(report["schema_version"], 1)
        recipe = mcp_server.answers_to_recipe(self.path, {}, unanswered="suggested")
        suggested = [
            option["rule"]
            for item in report["items"]
            for option in item["options"]
            if option["id"] == item["suggested"] and option["rule"]
        ]
        self.assertEqual(recipe.get("rules", []), suggested)
        with self.assertRaisesRegex(ValueError, "no review item"):
            mcp_server.answers_to_recipe(self.path, {"heading@s9.b1": "keep"})

    def test_convert_can_show_only_some_slides(self) -> None:
        self.assertEqual(mcp_server.convert_presentation(self.path, pages=[99]), "")
        self.assertIn("[Page_1]", mcp_server.convert_presentation(self.path, pages=[1]))

    def test_chunks_and_schema(self) -> None:
        chunks = mcp_server.presentation_chunks(self.path, max_chars=500)
        self.assertEqual(chunks[0]["metadata"]["page"], 1)
        self.assertIn("rules", mcp_server.recipe_schema()["properties"])


@unittest.skipUnless(importlib.util.find_spec("mcp"), "the mcp extra is not installed")
class SlideTests(unittest.TestCase):
    def setUp(self) -> None:
        from .test_ordering import build_tracks

        self.directory = tempfile.TemporaryDirectory()
        self.path = str(build_tracks(Path(self.directory.name) / "tracks.pptx"))

    def tearDown(self) -> None:
        self.directory.cleanup()

    def test_review_filters_narrow_the_list_but_not_the_counts(self) -> None:
        full = mcp_server.review_presentation(self.path)
        self.assertEqual(
            full["counts"],
            {
                "total": 1,
                "topic": {"reading-order": 1},
                "look": {"slide": 1},
                "impact": {"high": 1},
            },
        )
        for narrowed in (
            {"pages": [2]},
            {"looks": ["text"]},
            {"impact": "low"},
            {"topics": ["heading"]},
        ):
            report = mcp_server.review_presentation(self.path, **narrowed)
            self.assertEqual((report["items"], report["counts"]), ([], full["counts"]))
        shown = mcp_server.review_presentation(self.path, looks=["slide"], pages=[1])
        self.assertEqual([item["id"] for item in shown["items"]], ["reading-order@slide[1]"])

    @unittest.skipUnless(find_soffice(), "LibreOffice is not installed")
    def test_slides_review_asks_to_look_at_are_rendered(self) -> None:
        label, image = mcp_server.presentation_slides(self.path)
        self.assertEqual(label, "Slide 1:")
        self.assertTrue(image.data.startswith(b"\x89PNG"))

    def test_without_libreoffice_the_tool_says_so(self) -> None:
        from unittest import mock

        with mock.patch.object(mcp_server, "find_soffice", return_value=None):
            (message,) = mcp_server.presentation_slides(self.path)
        self.assertIn("LibreOffice is not installed", message)


class BoundaryTests(unittest.TestCase):
    def test_core_and_server_never_load_the_model_layer(self) -> None:
        import subprocess
        import sys

        code = (
            "import sys, pptx2markdown, pptx2markdown.mcp_server, pptx2markdown.ordering; "
            "assert 'anthropic' not in sys.modules"
        )
        subprocess.run([sys.executable, "-c", code], check=True)


@unittest.skipUnless(importlib.util.find_spec("mcp"), "the mcp extra is not installed")
class ServerTests(unittest.TestCase):
    def test_every_tool_is_registered_and_callable(self) -> None:
        server = mcp_server.build_server()
        tools = asyncio.run(server.list_tools())
        self.assertEqual(
            sorted(tool.name for tool in tools),
            sorted(tool.__name__ for tool in mcp_server.TOOLS),
        )
        result = asyncio.run(server.call_tool("recipe_schema", {}))
        self.assertFalse(result.is_error)
        text = result.content[0].text
        self.assertIn("rules", json.loads(text)["properties"])

    def test_workflow_prompts_are_registered_with_the_path(self) -> None:
        server = mcp_server.build_server()
        prompts = asyncio.run(server.list_prompts())
        self.assertEqual(
            sorted(prompt.name for prompt in prompts),
            sorted(prompt.__name__ for prompt in mcp_server.PROMPTS),
        )
        result = asyncio.run(
            server.get_prompt("prepare_for_retrieval", {"path": "/decks/q3.pptx"})
        )
        text = result.messages[0].content.text
        self.assertIn("/decks/q3.pptx", text)
        self.assertIn("audit_presentation", text)

    def test_pictures_needing_a_description_are_returned_as_images(self) -> None:
        from .pptx_factory import SLIDE

        undescribed = SLIDE.replace('name="Picture 1" descr="Sales chart"', 'name="Picture 1"')
        with tempfile.TemporaryDirectory() as directory:
            path = str(build_test_pptx(Path(directory) / "deck.pptx", undescribed))
            result = asyncio.run(
                mcp_server.build_server().call_tool("presentation_images", {"path": path})
            )
        self.assertFalse(result.is_error)
        self.assertEqual([item.type for item in result.content], ["text", "image"])
        self.assertTrue(result.content[0].text.startswith("s1.b"))


if __name__ == "__main__":
    unittest.main()
