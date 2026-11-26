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


def _box(name: str, x: float, y: float, w: float, h: float, **flags) -> LayoutNode[str]:
    inch = 914_400
    bbox = (int(x * inch), int(y * inch), int((x + w) * inch), int((y + h) * inch))
    return LayoutNode(name, _box.order.__next__(), bbox, **flags)


_box.order = iter(range(10_000))  # type: ignore[attr-defined]


class ContainerTests(unittest.TestCase):
    def test_cards_on_a_slide_covering_picture_are_read_card_by_card(self) -> None:
        nodes = [_box("background", 0, 0, 13.3, 7.5, structure=False)]
        for column, x in enumerate((0.5, 4.5, 8.5)):
            nodes.append(_box(f"frame{column}", x, 1, 3.6, 1.5, structure=False, container=True))
            nodes += [
                _box(f"head{column}", x + 0.2, 1.1, 3, 0.4),
                _box(f"body{column}", x + 0.2, 1.6, 3, 0.4),
            ]
        order = [n for n in geometry_reading_order(nodes) if n.startswith(("head", "body"))]
        self.assertEqual(order, ["head0", "body0", "head1", "body1", "head2", "body2"])

    def test_a_group_is_read_as_one_unit(self) -> None:
        # Staggered rows leave no column gap; each group still reads as a whole.
        a = [_box(f"a{i}", 0.5, 1 + 0.5 * i, 4, 0.4, groups=(1,)) for i in range(2)]
        b = [_box(f"b{i}", 5, 1 + 0.5 * i, 4, 0.4, groups=(2,)) for i in range(2)]
        c = [_box(f"c{i}", 2.5, 3 + 0.5 * i, 4, 0.4, groups=(3,)) for i in range(2)]
        order = geometry_reading_order([*c, *b, *a])
        self.assertEqual(order, ["a0", "a1", "b0", "b1", "c0", "c1"])

    def test_a_labelled_shape_reads_its_own_text_before_what_it_holds(self) -> None:
        nodes = [
            _box("inner", 1, 2, 3, 0.4),
            _box("label", 0.5, 1, 5, 3, container=True),
            _box("outside", 7, 1, 3, 0.4),
        ]
        self.assertEqual(geometry_reading_order(nodes), ["label", "inner", "outside"])

    def test_labels_on_a_picture_are_read_top_to_bottom(self) -> None:
        nodes = [
            _box("picture", 1, 1, 9, 5, container=True, annotated=True),
            _box("low-left", 1.5, 4, 2, 0.4),
            _box("high-left", 1.5, 1.5, 2, 0.4),
            _box("middle-right", 7, 2.5, 2, 0.4),
        ]
        self.assertEqual(
            geometry_reading_order(nodes), ["picture", "high-left", "middle-right", "low-left"]
        )

    def test_decoration_never_splits_columns(self) -> None:
        # An empty decorative group across the slide, and a bracket inside a container.
        nodes = [
            _box("panel", 0.2, 0.8, 12, 4, structure=False, groups=(1,)),
            _box("panel-edge", 0.2, 0.8, 2, 2, structure=False, groups=(1,)),
            _box("left-head", 0.5, 1, 3, 0.4),
            _box("right-head", 6, 1, 3, 0.4),
            _box("left-body", 0.5, 1.6, 3, 1),
            _box("right-body", 6, 1.6, 3, 1),
        ]
        order = [n for n in geometry_reading_order(nodes) if "panel" not in n]
        self.assertEqual(order, ["left-head", "left-body", "right-head", "right-body"])

    def test_a_tiny_box_barely_outside_is_not_held(self) -> None:
        # A 100-EMU-high line just above a title placeholder stays outside it.
        title = LayoutNode("title", 1, (100, 200, 8_000_100, 900_200), container=True)
        line = LayoutNode("line", 0, (100, 50, 3_000_100, 150))
        self.assertEqual(geometry_reading_order([title, line]), ["line", "title"])


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
