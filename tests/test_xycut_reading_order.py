from __future__ import annotations

import unittest

from pptx2markdown.structure_analyzer.structure import SlideObject, order_objects


def make_object(
    text: str,
    bbox: tuple[int, int, int, int],
    xml_index: int,
    *,
    is_heading: bool = False,
    is_title_placeholder: bool = False,
) -> SlideObject:
    return SlideObject(
        shape_id=str(xml_index),
        xml_index=xml_index,
        tag="sp",
        name=f"Shape {xml_index}",
        ph_type="title" if is_title_placeholder else None,
        ph_idx=None,
        x=bbox[0],
        y=bbox[1],
        coord_source="direct",
        text=text,
        normalized=text.casefold(),
        is_footer=False,
        is_decorative=False,
        is_heading=is_heading,
        is_title_placeholder=is_title_placeholder,
        font_pt=18.0,
        bbox=bbox,
    )


class XycutReadingOrderTests(unittest.TestCase):
    def test_contiguous_numbering_preserves_semantic_order(self) -> None:
        objects = [
            make_object("2. XML mode", (100_000, 300_000, 250_000, 340_000), 1),
            make_object("3. Surya mode", (100_000, 500_000, 250_000, 540_000), 2),
            make_object("4. Structure analyzer", (500_000, 200_000, 700_000, 240_000), 3),
            make_object("5. Surya pipeline", (500_000, 400_000, 700_000, 440_000), 4),
            make_object("6. Table pipeline", (500_000, 600_000, 700_000, 640_000), 5),
            make_object("1. Main converter", (100_000, 100_000, 300_000, 140_000), 6),
        ]

        ordered = order_objects(objects, mode="xycut")

        self.assertEqual(
            [obj.text for obj in ordered],
            [
                "1. Main converter",
                "2. XML mode",
                "3. Surya mode",
                "4. Structure analyzer",
                "5. Surya pipeline",
                "6. Table pipeline",
            ],
        )

    def test_comparison_columns_keep_heading_with_its_body(self) -> None:
        objects = [
            make_object(
                "AS-IS",
                (100_000, 100_000, 350_000, 220_000),
                1,
                is_heading=True,
                is_title_placeholder=True,
            ),
            make_object("TO-BE", (650_000, 100_000, 900_000, 220_000), 2),
            make_object("Left column body", (50_000, 300_000, 450_000, 700_000), 3),
            make_object("Right column body", (550_000, 300_000, 950_000, 700_000), 4),
        ]

        ordered = order_objects(objects, mode="xycut")

        self.assertEqual(
            [obj.text for obj in ordered],
            ["AS-IS", "Left column body", "TO-BE", "Right column body"],
        )

    def test_regular_rows_keep_top_left_order(self) -> None:
        objects = [
            make_object("Bottom right", (500_000, 400_000, 700_000, 450_000), 1),
            make_object("Top right", (500_000, 100_000, 700_000, 150_000), 2),
            make_object("Bottom left", (100_000, 400_000, 300_000, 450_000), 3),
            make_object("Top left", (100_000, 100_000, 300_000, 150_000), 4),
        ]

        ordered = order_objects(objects, mode="xycut")

        self.assertEqual(
            [obj.text for obj in ordered],
            ["Top left", "Top right", "Bottom left", "Bottom right"],
        )


if __name__ == "__main__":
    unittest.main()
