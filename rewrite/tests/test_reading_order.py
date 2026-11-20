from __future__ import annotations

import unittest

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
