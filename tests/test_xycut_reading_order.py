from __future__ import annotations

import unittest

from pptx2markdown.structure_analyzer.structure import SlideObject, order_objects


def make_object(
    text: str,
    bbox: tuple[int, int, int, int],
    xml_index: int,
) -> SlideObject:
    return SlideObject(
        shape_id=str(xml_index),
        xml_index=xml_index,
        tag="sp",
        name=f"Shape {xml_index}",
        ph_type=None,
        ph_idx=None,
        x=bbox[0],
        y=bbox[1],
        coord_source="direct",
        text=text,
        normalized=text.casefold(),
        is_footer=False,
        is_decorative=False,
        is_heading=False,
        is_title_placeholder=False,
        font_pt=18.0,
        bbox=bbox,
    )


class XycutReadingOrderTests(unittest.TestCase):
    def test_numbered_text_does_not_override_geometry(self) -> None:
        objects = [
            make_object("2. XML mode", (100_000, 300_000, 250_000, 340_000), 1),
            make_object("3. Surya mode", (100_000, 500_000, 250_000, 540_000), 2),
            make_object("4. Structure analyzer", (500_000, 200_000, 700_000, 240_000), 3),
            make_object("5. Surya pipeline", (500_000, 400_000, 700_000, 440_000), 4),
            make_object("6. Table pipeline", (500_000, 600_000, 700_000, 640_000), 5),
            make_object("1. Main converter", (100_000, 100_000, 300_000, 140_000), 6),
        ]

        ordered = order_objects(objects)

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

    def test_column_labels_do_not_override_geometry(self) -> None:
        objects = [
            make_object("AS-IS", (100_000, 100_000, 350_000, 220_000), 1),
            make_object("TO-BE", (650_000, 100_000, 900_000, 220_000), 2),
            make_object("Left column body", (50_000, 300_000, 450_000, 700_000), 3),
            make_object("Right column body", (550_000, 300_000, 950_000, 700_000), 4),
        ]

        ordered = order_objects(objects)

        self.assertEqual(
            [obj.text for obj in ordered],
            ["AS-IS", "Left column body", "TO-BE", "Right column body"],
        )

    def test_grid_uses_column_order_when_vertical_cut_crosses_region(self) -> None:
        objects = [
            make_object("Bottom right", (500_000, 400_000, 700_000, 450_000), 1),
            make_object("Top right", (500_000, 100_000, 700_000, 150_000), 2),
            make_object("Bottom left", (100_000, 400_000, 300_000, 450_000), 3),
            make_object("Top left", (100_000, 100_000, 300_000, 150_000), 4),
        ]

        ordered = order_objects(objects)

        self.assertEqual(
            [obj.text for obj in ordered],
            ["Top left", "Bottom left", "Top right", "Bottom right"],
        )

    def test_tight_staggered_stream_uses_vertical_order(self) -> None:
        objects = [
            make_object("First", (100_000, 100_000, 400_000, 200_000), 1),
            make_object("Second", (600_000, 210_000, 900_000, 310_000), 2),
            make_object("Third", (100_000, 320_000, 400_000, 420_000), 3),
            make_object("Fourth", (600_000, 430_000, 900_000, 530_000), 4),
        ]

        ordered = order_objects(objects)

        self.assertEqual([obj.text for obj in ordered], ["First", "Second", "Third", "Fourth"])

    def test_full_width_top_region_keeps_columns_below_it(self) -> None:
        objects = [
            make_object("Top region", (50_000, 50_000, 950_000, 150_000), 1),
            make_object("Left upper", (100_000, 300_000, 400_000, 400_000), 2),
            make_object("Right upper", (600_000, 300_000, 900_000, 400_000), 3),
            make_object("Left lower", (100_000, 500_000, 400_000, 600_000), 4),
            make_object("Right lower", (600_000, 500_000, 900_000, 600_000), 5),
        ]

        ordered = order_objects(objects)

        self.assertEqual(
            [obj.text for obj in ordered],
            ["Top region", "Left upper", "Left lower", "Right upper", "Right lower"],
        )

    def test_single_column_uses_top_to_bottom_order(self) -> None:
        objects = [
            make_object("Bottom", (100_000, 500_000, 400_000, 600_000), 1),
            make_object("Top", (100_000, 100_000, 400_000, 200_000), 2),
            make_object("Middle", (100_000, 300_000, 400_000, 400_000), 3),
        ]

        ordered = order_objects(objects)

        self.assertEqual([obj.text for obj in ordered], ["Top", "Middle", "Bottom"])

    def test_text_and_heading_signals_do_not_change_order(self) -> None:
        bboxes = [
            (100_000, 100_000, 400_000, 200_000),
            (100_000, 300_000, 400_000, 400_000),
            (600_000, 100_000, 900_000, 200_000),
            (600_000, 300_000, 900_000, 400_000),
        ]
        semantic = [
            make_object("1. First", bbox, index) for index, bbox in enumerate(bboxes, start=1)
        ]
        generic = [
            make_object(f"Block {index}", bbox, index)
            for index, bbox in enumerate(bboxes, start=1)
        ]
        semantic[0].is_heading = True
        semantic[0].is_title_placeholder = True
        semantic[1].is_footer = True
        semantic[2].is_decorative = True

        semantic_order = [obj.shape_id for obj in order_objects(semantic)]
        generic_order = [obj.shape_id for obj in order_objects(generic)]

        self.assertEqual(semantic_order, generic_order)

    def test_nearly_touching_column_boundaries_are_separated(self) -> None:
        objects = [
            make_object("Left upper", (100_000, 100_000, 500_001, 200_000), 1),
            make_object("Left lower", (100_000, 300_000, 500_001, 400_000), 2),
            make_object("Right tall", (500_000, 100_000, 900_000, 400_000), 3),
        ]

        ordered = order_objects(objects)

        self.assertEqual(
            [obj.text for obj in ordered],
            ["Left upper", "Left lower", "Right tall"],
        )

    def test_missing_bbox_does_not_use_text_to_invent_geometry(self) -> None:
        positioned = make_object("Positioned", (100_000, 100_000, 300_000, 200_000), 3)
        missing_bbox = make_object("A very long numbered 1. semantic label", (0, 0, 1, 1), 1)
        missing_bbox.bbox = None

        ordered = order_objects([missing_bbox, positioned])

        self.assertEqual([obj.text for obj in ordered], ["Positioned", missing_bbox.text])


if __name__ == "__main__":
    unittest.main()
