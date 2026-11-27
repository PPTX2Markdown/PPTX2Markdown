from __future__ import annotations

import unittest

from pptx2markdown.headings import (
    Heading,
    HeadingDetector,
    HeadingHint,
    ShapeSignals,
    TextShape,
    heading_hints,
)


def text_shape(
    text: str,
    *,
    placeholder_type: str | None = None,
    font_pt: float | None = None,
    has_list: bool = False,
    has_math: bool = False,
) -> TextShape:
    return TextShape(
        markdown=text,
        plain=text,
        placeholder_type=placeholder_type,
        font_pt=font_pt,
        has_list=has_list,
        has_math=has_math,
    )


class HeadingHintTests(unittest.TestCase):
    def test_title_placeholder_and_numbered_sections_get_depths(self) -> None:
        hints = heading_hints(
            [
                ShapeSignals("Report", "title", 40.0, (0, 0, 9_000_000, 800_000)),
                ShapeSignals("1. Scope", None, 28.0, (0, 1_000_000, 9_000_000, 1_500_000)),
                ShapeSignals("Body text", None, 18.0, (0, 2_000_000, 9_000_000, 3_000_000)),
            ]
        )
        self.assertEqual([hint.depth for hint in hints], [1, 2, None])
        self.assertEqual([hint.candidate for hint in hints], [True, True, False])

    def test_bullet_prefixed_text_is_not_a_candidate(self) -> None:
        (hint,) = heading_hints([ShapeSignals("- 1. Scope", None, 28.0, None)])
        self.assertFalse(hint.candidate)


class HeadingDetectorTests(unittest.TestCase):
    def test_placeholders_map_to_levels(self) -> None:
        detector = HeadingDetector()
        self.assertEqual(
            detector.classify(text_shape("Report", placeholder_type="title"), HeadingHint()),
            Heading(level=1, text="Report", source="placeholder"),
        )
        self.assertEqual(
            detector.classify(text_shape("Q3", placeholder_type="subTitle"), HeadingHint()),
            Heading(level=2, text="Q3", source="placeholder"),
        )

    def test_large_first_text_block_falls_back_to_h1(self) -> None:
        heading = HeadingDetector().classify(text_shape("Agenda", font_pt=28.0), HeadingHint())
        self.assertEqual(heading, Heading(level=1, text="Agenda", source="font"))

    def test_a_list_of_several_numbered_items_is_not_a_heading(self) -> None:
        shape = TextShape(
            markdown="1. PDF parsing\n2. Table loss\n3. Formula errors",
            plain="PDF parsing Table loss Formula errors",
            placeholder_type=None,
            font_pt=28.0,
            has_list=True,
            has_math=False,
        )
        self.assertIsNone(HeadingDetector().classify(shape, HeadingHint()))

    def test_numbered_fallback_is_numbering_not_font(self) -> None:
        heading = HeadingDetector().classify(text_shape("2. Results", font_pt=28.0), HeadingHint())
        self.assertEqual(heading.source, "numbering")

    def test_small_or_later_text_blocks_are_not_headings(self) -> None:
        self.assertIsNone(
            HeadingDetector().classify(text_shape("Agenda", font_pt=18.0), HeadingHint())
        )
        detector = HeadingDetector()
        detector.classify(text_shape("Intro paragraph", font_pt=18.0), HeadingHint())
        self.assertIsNone(detector.classify(text_shape("Agenda", font_pt=28.0), HeadingHint()))

    def test_long_sentences_and_numbered_lists_are_rejected(self) -> None:
        hint = HeadingHint(depth=2, score=0.75)
        long_text = "This sentence is clearly body text, not a heading."
        self.assertIsNone(HeadingDetector().classify(text_shape(long_text), HeadingHint()))
        self.assertIsNone(HeadingDetector().classify(text_shape("1. a 2. b"), hint))

    def test_equations_are_never_headings(self) -> None:
        shape = text_shape("$x$", placeholder_type="title", has_math=True)
        self.assertIsNone(HeadingDetector().classify(shape, HeadingHint(depth=1, score=0.95)))

    def test_repeated_heading_is_reported_as_duplicate(self) -> None:
        detector = HeadingDetector()
        title = text_shape("Report", placeholder_type="title")
        self.assertIsInstance(detector.classify(title, HeadingHint()), Heading)
        self.assertEqual(detector.classify(title, HeadingHint()), "duplicate")

    def test_large_first_line_splits_title_from_body(self) -> None:
        shape = TextShape(
            markdown="Report\nPrepared by the team",
            plain="Report Prepared by the team",
            placeholder_type="title",
            font_pt=40.0,
            has_list=False,
            has_math=False,
            font_sections=("Report", "Prepared by the team"),
        )
        self.assertEqual(
            HeadingDetector().classify(shape, HeadingHint()),
            Heading(level=1, text="Report", trailing="Prepared by the team", source="placeholder"),
        )


if __name__ == "__main__":
    unittest.main()
