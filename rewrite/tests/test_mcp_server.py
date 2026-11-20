from __future__ import annotations

import asyncio
import importlib.util
import json
import tempfile
import unittest
from pathlib import Path

from pptx2markdown import mcp_server

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

    def test_chunks_and_schema(self) -> None:
        chunks = mcp_server.presentation_chunks(self.path, max_chars=500)
        self.assertEqual(chunks[0]["metadata"]["page"], 1)
        self.assertIn("rules", mcp_server.recipe_schema()["properties"])


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


if __name__ == "__main__":
    unittest.main()
