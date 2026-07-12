from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

import pptx2markdown
from pptx2markdown.main_converter.converter_models import PresentationDocument

FIXTURES_DIR = Path(__file__).parent / "fixtures"
PPTX_FIXTURE = FIXTURES_DIR / "xycut_layout_cases.pptx"
EXPECTED_FIXTURE = FIXTURES_DIR / "xycut_layout_cases.expected.json"


class XycutLayoutPptxFixtureTests(unittest.TestCase):
    def _convert(self, root: Path) -> tuple[dict[str, object], bytes]:
        output_dir = root / "output"
        work_dir = root / "work"
        exit_code = pptx2markdown.convert(
            PPTX_FIXTURE,
            output_dir=output_dir,
            work_dir=work_dir,
            output_format="json",
            placeholder_inheritance="none",
            inherited_shapes="none",
        )
        self.assertEqual(exit_code, 0)
        output_path = output_dir / "xycut_layout_cases" / "xycut_layout_cases.json"
        payload_bytes = output_path.read_bytes()
        return json.loads(payload_bytes), payload_bytes

    def test_real_pptx_matches_layout_golden_order_and_block_kinds(self) -> None:
        expected = json.loads(EXPECTED_FIXTURE.read_text(encoding="utf-8"))
        with tempfile.TemporaryDirectory() as temp_dir:
            actual, _ = self._convert(Path(temp_dir))

        self.assertEqual(actual["schema_version"], expected["schema_version"])
        self.assertEqual(actual["source"], expected["source"])
        self.assertEqual(len(actual["slides"]), len(expected["cases"]))

        for slide, case in zip(actual["slides"], expected["cases"]):
            actual_blocks = [[block["kind"], block["content"]] for block in slide["blocks"]]
            self.assertEqual(slide["page"], case["slide"])
            self.assertEqual(actual_blocks, case["blocks"], msg=case["policy"])

    def test_real_pptx_emits_schema_valid_geometry_and_provenance(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            actual, _ = self._convert(Path(temp_dir))

        document = PresentationDocument.model_validate(actual)
        for slide in document.slides:
            for block in slide.blocks:
                self.assertIsNotNone(block.bbox)
                self.assertEqual(block.bbox.unit, "emu")
                self.assertGreater(block.bbox.width, 0)
                self.assertGreater(block.bbox.height, 0)
                self.assertEqual(block.source_part, "slide")

        overlap_blocks = {block.content: block for block in document.slides[3].blocks}
        self.assertEqual(overlap_blocks["L4 BACK"].bbox, overlap_blocks["L4 FRONT"].bbox)

    def test_json_output_is_byte_deterministic_across_work_directories(self) -> None:
        with (
            tempfile.TemporaryDirectory() as first_dir,
            tempfile.TemporaryDirectory() as second_dir,
        ):
            _, first = self._convert(Path(first_dir))
            _, second = self._convert(Path(second_dir))

        self.assertEqual(first, second)


if __name__ == "__main__":
    unittest.main()
