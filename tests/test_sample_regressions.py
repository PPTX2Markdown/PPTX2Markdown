from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

import pptx2markdown

SAMPLE_NAMES = (
    "Heading_Test.pptx",
    "formula_demo.pptx",
    "inheritance_chain_sample.pptx",
    "reading_order_test.pptx",
    "sample1.pptx",
    "smartArt.pptx",
    "table_demo.pptx",
    "문제점 목록 발표.pptx",
)


class SampleRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.repo_root = Path(__file__).resolve().parents[1]
        cls.samples_dir = cls.repo_root / "pptx_samples"
        missing = [name for name in SAMPLE_NAMES if not (cls.samples_dir / name).exists()]
        if missing:
            raise unittest.SkipTest(
                "local regression samples are unavailable: " + ", ".join(missing)
            )

        cls.temp_dir = tempfile.TemporaryDirectory()
        temp_root = Path(cls.temp_dir.name)
        cls.output_dir = temp_root / "output"
        exit_code = pptx2markdown.convert(
            [cls.samples_dir / name for name in SAMPLE_NAMES],
            output_dir=cls.output_dir,
            work_dir=temp_root / "work",
            reading_order="xycut",
            headings="auto",
        )
        if exit_code != 0:
            raise AssertionError(f"sample conversion failed with exit code {exit_code}")
        cls.mode_dir = cls.output_dir / "xycut"

    @classmethod
    def tearDownClass(cls) -> None:
        if hasattr(cls, "temp_dir"):
            cls.temp_dir.cleanup()

    @classmethod
    def _markdown(cls, stem: str) -> str:
        return (cls.mode_dir / stem / f"{stem}.md").read_text(encoding="utf-8")

    @staticmethod
    def _page(markdown: str, number: int) -> str:
        marker = f"[Page_{number}]"
        next_marker = f"[Page_{number + 1}]"
        _, page_and_rest = markdown.split(marker, maxsplit=1)
        return page_and_rest.split(next_marker, maxsplit=1)[0]

    def test_manifest_has_expected_pipeline_counts(self) -> None:
        manifest = json.loads(
            (self.mode_dir / "convert_manifest.json").read_text(encoding="utf-8")
        )
        summary = manifest["summary"]

        self.assertEqual(summary["processed_packages"], 8)
        self.assertEqual(summary["processed_slides"], 69)
        self.assertEqual(summary["failed"], 0)
        self.assertEqual(summary["math_blocks"], 50)
        self.assertEqual(summary["chart_blocks"], 1)
        self.assertEqual(summary["smartart_blocks"], 8)
        self.assertEqual(summary["table_blocks"], 23)
        self.assertEqual(summary["unresolved_images"], 0)

    def test_xycut_preserves_numbered_and_comparison_order(self) -> None:
        reading_order = self._page(self._markdown("reading_order_test"), 3)
        expected = (
            "1. Main converter",
            "2. Xml mode",
            "3. Surya mode",
            "4. Sturucture_analyzer",
            "5. Surya_pipeline",
            "6. table_pipeline",
        )
        positions = [reading_order.index(text) for text in expected]
        self.assertEqual(positions, sorted(positions))

        comparison = self._page(self._markdown("문제점 목록 발표"), 15)
        positions = [
            comparison.index("AS – IS"),
            comparison.index("PDF → 이미지"),
            comparison.index("TO-BE"),
            comparison.index("원본 파일"),
        ]
        self.assertEqual(positions, sorted(positions))

    def test_chart_and_master_content_are_preserved_without_stripe_image(self) -> None:
        markdown = self._markdown("sample1")

        self.assertIn("**Chart type:** Scatter Chart", markdown)
        self.assertIn("End of The Document", self._page(markdown, 11))
        self.assertNotIn("media/image1.png", markdown)

    def test_math_table_and_smartart_outputs_are_present(self) -> None:
        self.assertIn(r"\sigma = E \epsilon", self._markdown("formula_demo"))
        self.assertIn("| CustomerID | CustomerName | Orders", self._markdown("table_demo"))
        self.assertIn("- 1\n  - 2\n  - 3", self._markdown("smartArt"))

    def test_inherited_shapes_are_materialized(self) -> None:
        markdown = self._markdown("inheritance_chain_sample")

        self.assertIn("MASTER_ONLY_SHAPE", markdown)
        self.assertIn("LAYOUT1_ONLY_SHAPE", markdown)
        self.assertIn("SLIDE2_BODY_USES_LAYOUT_OFFSET_MASTER_EXTENT_MASTER_BULLET", markdown)

    def test_auto_heading_mode_keeps_expected_hierarchy(self) -> None:
        markdown = self._markdown("Heading_Test")

        self.assertIn("## 2. Text Box Heading – but font size look like Heading", markdown)
        self.assertIn("### 2. Text Box Heading", self._page(markdown, 3))


if __name__ == "__main__":
    unittest.main()
