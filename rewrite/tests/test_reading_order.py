from __future__ import annotations

import json
import unittest
from pathlib import Path

import pptx2markdown
from pptx2markdown.reading_order import LayoutNode, geometry_reading_order


class ReadingOrderTests(unittest.TestCase):
    def test_columns_are_read_left_to_right_then_top_to_bottom(self) -> None:
        nodes = [
            LayoutNode("right-bottom", 3, (600, 600, 900, 900)),
            LayoutNode("left-bottom", 1, (0, 600, 300, 900)),
            LayoutNode("right-top", 2, (600, 0, 900, 300)),
            LayoutNode("left-top", 0, (0, 0, 300, 300)),
        ]
        self.assertEqual(
            geometry_reading_order(nodes),
            ["left-top", "right-top", "left-bottom", "right-bottom"],
        )

    def test_unpositioned_nodes_are_appended_in_source_order(self) -> None:
        nodes = [
            LayoutNode("unpositioned-2", 2, None),
            LayoutNode("positioned", 1, (0, 0, 10, 10)),
            LayoutNode("unpositioned-1", 0, None),
        ]
        self.assertEqual(
            geometry_reading_order(nodes),
            ["positioned", "unpositioned-1", "unpositioned-2"],
        )


if __name__ == "__main__":
    unittest.main()


class StructureTests(unittest.TestCase):
    def test_a_background_frame_does_not_seal_the_column_gap(self) -> None:
        inch = 914_400
        nodes = [
            LayoutNode("frame", 0, (0, 0, 10 * inch, 7 * inch), structure=False),
            LayoutNode("right-top", 1, (6 * inch, 1 * inch, 9 * inch, 2 * inch)),
            LayoutNode("left-bottom", 2, (1 * inch, 3 * inch, 4 * inch, 4 * inch)),
            LayoutNode("left-top", 3, (1 * inch, 1 * inch, 4 * inch, 2 * inch)),
            LayoutNode("right-bottom", 4, (6 * inch, 2_500_000, 9 * inch, 3_400_000)),
        ]
        self.assertEqual(
            geometry_reading_order(nodes),
            ["left-top", "left-bottom", "right-top", "right-bottom", "frame"],
        )

    def test_a_stack_of_offset_boxes_is_read_top_to_bottom(self) -> None:
        nodes = [
            LayoutNode("second", 1, (5000, 400, 9000, 700)),
            LayoutNode("first", 2, (0, 0, 4000, 300)),
            LayoutNode("third", 3, (10000, 800, 14000, 1100)),
        ]
        self.assertEqual(geometry_reading_order(nodes), ["first", "second", "third"])


FIXTURES = Path(__file__).resolve().parents[2] / "tests" / "fixtures"


def _first_lines(slide) -> list[str]:
    lines = []
    for block in slide.blocks:
        if block.kind == "list":
            text = block.items[0].text
        else:
            text = getattr(block, "text", None) or getattr(block, "markdown", "")
        lines.append(text.replace("**", "").replace("\\", "").split("\n")[0])
    return lines


@unittest.skipUnless(FIXTURES.is_dir(), "repository fixtures are not available")
class ExpectedReadingOrderTests(unittest.TestCase):
    """The repository's reading-order cases: columns, dividers, jitter and overlaps."""

    def _check(self, name: str, key: str) -> None:
        expected = json.loads((FIXTURES / f"{name}.expected.json").read_text(encoding="utf-8"))
        with pptx2markdown.parse(FIXTURES / f"{name}.pptx") as presentation:
            slides = {slide.page: slide for slide in presentation.document.slides}
        for case in expected["cases"]:
            want = [item[1] if isinstance(item, list) else item for item in case[key]]
            with self.subTest(slide=case["slide"], policy=case["policy"]):
                self.assertEqual(
                    _first_lines(slides[case["slide"]]), [text.split("\n")[0] for text in want]
                )

    def test_layout_cases(self) -> None:
        self._check("xycut_layout_cases", "blocks")

    def test_tolerance_cases(self) -> None:
        self._check("xycut_tolerance_cases", "expected_order")
