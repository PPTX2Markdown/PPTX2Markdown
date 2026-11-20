from __future__ import annotations

import importlib.util
import json
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT = REPO_ROOT / "scripts" / "parity.py"

_spec = importlib.util.spec_from_file_location("parity_harness", SCRIPT)
parity = importlib.util.module_from_spec(_spec)
sys.modules[_spec.name] = parity
_spec.loader.exec_module(parity)


def block(kind: str, text: str = "", key: str | None = None, level: int | None = None):
    return {"kind": kind, "text": text, "key": key, "level": level}


def slide(page: int, blocks, notes=None, hidden=False):
    return {"page": page, "hidden": hidden, "notes": notes, "blocks": list(blocks)}


class TokenTests(unittest.TestCase):
    def test_tokens_ignore_markdown_syntax_and_asset_paths(self) -> None:
        text = "![image](media/a.png) [label](https://x.test/y)\n| a | b |\n|---|---|\n**bold**"
        self.assertEqual(parity.tokens(text), ["label", "a", "b", "bold"])

    def test_tokens_ignore_emphasis_and_escapes_inside_words(self) -> None:
        self.assertEqual(
            parity.tokens("**도구**의 snake\\_case \\# tag"), ["도구의", "snake_case", "tag"]
        )

    def test_display_lines_canonicalize_table_padding(self) -> None:
        self.assertEqual(
            parity.display_lines("| a   | b |\n| --- | :-: |\n![x](m.png)"),
            ["| a | b |", "|---|", "<image>"],
        )


class NormalizationTests(unittest.TestCase):
    def test_both_engines_normalize_to_the_same_shape(self) -> None:
        legacy = {
            "slides": [
                {
                    "page": 1,
                    "hidden": False,
                    "notes": "n",
                    "blocks": [
                        {
                            "kind": "heading",
                            "content": "T",
                            "shape_id": "2",
                            "heading_level": 1,
                            "source_part": "slide",
                        },
                        {
                            "kind": "list",
                            "content": "- a\n- b",
                            "shape_id": "3",
                            "heading_level": None,
                            "source_part": "slide",
                        },
                        {
                            "kind": "image",
                            "content": "![image](media/i.png)",
                            "shape_id": "4",
                            "heading_level": None,
                            "source_part": "slide",
                        },
                    ],
                }
            ]
        }
        rewrite = {
            "slides": [
                {
                    "page": 1,
                    "hidden": False,
                    "notes": "n",
                    "blocks": [
                        {
                            "kind": "heading",
                            "text": "T",
                            "level": 1,
                            "shape_id": "2",
                            "source_part": "slide",
                        },
                        {
                            "kind": "list",
                            "items": [{"text": "a", "level": 0}, {"text": "b", "level": 0}],
                            "ordered": False,
                            "shape_id": "3",
                            "source_part": "slide",
                        },
                        {
                            "kind": "image",
                            "asset_id": "asset-1",
                            "alt_text": "Pic",
                            "shape_id": "4",
                            "source_part": "slide",
                        },
                    ],
                }
            ]
        }
        self.assertEqual(parity.normalize_legacy(legacy), parity.normalize_rewrite(rewrite))


class CompareSlideTests(unittest.TestCase):
    def test_identical_slides_pass(self) -> None:
        blocks = [block("heading", "Title", "slide:1", 1), block("text", "Body", "slide:2")]
        result = parity.compare_slide(slide(1, blocks), slide(1, blocks))
        self.assertTrue(result["passed"])
        self.assertEqual(result["recall"], 1.0)

    def test_heading_demoted_to_text_is_reported(self) -> None:
        result = parity.compare_slide(
            slide(1, [block("heading", "Title", "slide:1", 1)]),
            slide(1, [block("text", "Title", "slide:1")]),
        )
        self.assertFalse(result["passed"])
        self.assertEqual(result["recall"], 1.0)
        self.assertFalse(result["headings_equal"])
        self.assertEqual(result["shapes"][0]["status"], "differs")

    def test_missing_and_extra_shapes_affect_recall_and_precision(self) -> None:
        result = parity.compare_slide(
            slide(1, [block("text", "keep", "slide:1"), block("math", "x y", "slide:2")]),
            slide(1, [block("text", "keep", "slide:1"), block("text", "date", "slide:3")]),
        )
        self.assertAlmostEqual(result["recall"], 1 / 3)
        self.assertAlmostEqual(result["precision"], 1 / 2)
        statuses = {row["key"]: row["status"] for row in result["shapes"]}
        self.assertEqual(statuses, {"slide:1": "same", "slide:2": "missing", "slide:3": "extra"})

    def test_reordered_shapes_lower_the_order_score(self) -> None:
        a, b, c = (block("text", t, f"slide:{t}") for t in "abc")
        result = parity.compare_slide(slide(1, [a, b, c]), slide(1, [c, a, b]))
        self.assertLess(result["order"], 1.0)
        self.assertFalse(result["passed"])

    def test_notes_and_hidden_flags_are_compared(self) -> None:
        result = parity.compare_slide(
            slide(1, [], notes="one", hidden=True), slide(1, [], notes="two")
        )
        self.assertFalse(result["notes_equal"])
        self.assertFalse(result["hidden_equal"])
        self.assertFalse(result["passed"])


class CompareDocumentTests(unittest.TestCase):
    def test_slide_missing_from_rewrite_counts_against_recall(self) -> None:
        legacy = {
            "slides": [
                slide(1, [block("text", "a", "slide:1")]),
                slide(2, [block("text", "b", "slide:1")]),
            ]
        }
        rewrite = {"slides": [slide(1, [block("text", "a", "slide:1")])]}
        result = parity.compare_documents(legacy, rewrite)
        self.assertEqual(result["summary"]["slides_passed"], 1)
        self.assertEqual(result["slides"][1]["status"], "missing_in_rewrite")
        self.assertEqual(result["summary"]["recall"], 0.5)
        self.assertEqual(result["summary"]["shape_parity"]["text"], {"shapes": 2, "same": 1})


@unittest.skipUnless(
    (REPO_ROOT / "rewrite" / "src" / "pptx2markdown").is_dir(), "rewrite sources not present"
)
class EndToEndTests(unittest.TestCase):
    def test_harness_runs_both_engines_on_a_fixture(self) -> None:
        fixture = REPO_ROOT / "tests" / "fixtures" / "golden" / "python-pptx_txt-text.pptx"
        with tempfile.TemporaryDirectory() as temp:
            with redirect_stdout(StringIO()):
                exit_code = parity.main([str(fixture), "--report-dir", temp, "--jobs", "2"])
            self.assertEqual(exit_code, 0)
            report = json.loads((Path(temp) / "report.json").read_text(encoding="utf-8"))
            deck = report["decks"][0]
            self.assertEqual(deck["engines"]["legacy"]["status"], "ok")
            self.assertEqual(deck["engines"]["rewrite"]["status"], "ok")
            self.assertIsNotNone(deck["comparison"])
            run_dir = Path(temp) / "runs" / deck["id"]
            self.assertTrue((run_dir / "legacy.txt").is_file())
            self.assertTrue((run_dir / "rewrite.txt").is_file())


class BaselineTests(unittest.TestCase):
    BASE = {
        "slides_passed": 10,
        "slides_compared": 12,
        "recall": 0.9,
        "precision": 0.95,
        "same_shapes": {"table": 3, "text": 8},
        "failed_runs": ["old.pptx:rewrite"],
    }

    def test_equal_or_better_numbers_pass(self) -> None:
        better = {**self.BASE, "slides_passed": 11, "same_shapes": {"table": 4, "text": 8}}
        self.assertEqual(parity.compare_baseline(self.BASE, better), [])

    def test_every_worse_number_is_reported(self) -> None:
        worse = {
            **self.BASE,
            "slides_passed": 9,
            "recall": 0.89,
            "same_shapes": {"text": 8},
            "failed_runs": ["new.pptx:rewrite"],
        }
        self.assertEqual(
            parity.compare_baseline(self.BASE, worse),
            [
                "slides_passed: 10 -> 9",
                "recall: 0.9 -> 0.89",
                "identical table shapes: 3 -> 0",
                "runs that now fail: new.pptx:rewrite",
            ],
        )


if __name__ == "__main__":
    unittest.main()
