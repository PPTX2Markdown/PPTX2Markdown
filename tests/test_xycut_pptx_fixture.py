from __future__ import annotations

import json
import tempfile
import unittest
import zipfile
from pathlib import Path

from pptx2markdown.structure_analyzer.extractor import extract_slide_objects_xml
from pptx2markdown.structure_analyzer.pipeline import write_outputs
from pptx2markdown.structure_analyzer.structure import (
    XYCUT_MAX_OVERLAP_EMU,
    _projection_chunks,
)

FIXTURES_DIR = Path(__file__).parent / "fixtures"
PPTX_FIXTURE = FIXTURES_DIR / "xycut_tolerance_cases.pptx"
EXPECTED_FIXTURE = FIXTURES_DIR / "xycut_tolerance_cases.expected.json"


class XycutPptxFixtureTests(unittest.TestCase):
    def _analyze_fixture(self, root: Path) -> list[dict[str, object]]:
        package_dir = root / "package"
        with zipfile.ZipFile(PPTX_FIXTURE) as archive:
            archive.extractall(package_dir)

        expected = json.loads(EXPECTED_FIXTURE.read_text(encoding="utf-8"))
        reports: list[dict[str, object]] = []
        for case in expected["cases"]:
            slide_number = int(case["slide"])
            output_dir = root / "analysis" / f"slide{slide_number}"
            output_dir.mkdir(parents=True, exist_ok=True)
            result = write_outputs(
                package_dir / "ppt" / "slides" / f"slide{slide_number}.xml",
                output_dir,
                pptx_inheritance="none",
                inherited_shapes="none",
            )
            reports.append(json.loads(Path(result["output_json"]).read_text(encoding="utf-8")))
        return reports

    @staticmethod
    def _objects_by_text(report: dict[str, object]) -> dict[str, dict[str, object]]:
        return {
            str(row["text"]): row for row in report["structure_order"] if isinstance(row, dict)
        }

    def test_actual_pptx_preserves_expected_xycut_order(self) -> None:
        expected = json.loads(EXPECTED_FIXTURE.read_text(encoding="utf-8"))
        with tempfile.TemporaryDirectory() as temp_dir:
            reports = self._analyze_fixture(Path(temp_dir))

        self.assertEqual(len(reports), 6)
        for case, report in zip(expected["cases"], reports):
            actual_order = [str(row["text"]) for row in report["structure_order"]]
            self.assertEqual(actual_order, case["expected_order"], msg=case["policy"])

    def test_actual_pptx_retains_threshold_boundary_geometry(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            reports = self._analyze_fixture(Path(temp_dir))

        slide1 = self._objects_by_text(reports[0])
        slide2 = self._objects_by_text(reports[1])
        slide3 = self._objects_by_text(reports[2])
        slide4 = self._objects_by_text(reports[3])
        slide5 = self._objects_by_text(reports[4])
        slide6 = self._objects_by_text(reports[5])

        self.assertEqual(
            slide1["C1 RIGHT — y=243"]["bbox"][1] - slide1["C1 LEFT — y=240"]["bbox"][1],
            3 * XYCUT_MAX_OVERLAP_EMU,
        )

        column_x_positions = [
            slide2[label]["bbox"][0]
            for label in ("C2 TOP — x=400", "C2 MIDDLE — x=403", "C2 BOTTOM — x=401")
        ]
        self.assertEqual(
            max(column_x_positions) - min(column_x_positions),
            3 * XYCUT_MAX_OVERLAP_EMU,
        )

        one_px_overlap = slide3["C3 RIGHT TOP"]["bbox"][0] - slide3["C3 LEFT TOP"]["bbox"][2]
        two_px_overlap = slide4["C4 RIGHT TOP"]["bbox"][0] - slide4["C4 LEFT TOP"]["bbox"][2]
        tiny_positive_gap = slide5["C5 RIGHT TOP"]["bbox"][0] - slide5["C5 LEFT TOP"]["bbox"][2]
        self.assertEqual(one_px_overlap, -XYCUT_MAX_OVERLAP_EMU)
        self.assertEqual(two_px_overlap, -(2 * XYCUT_MAX_OVERLAP_EMU))
        self.assertGreater(tiny_positive_gap, 0)
        self.assertLess(tiny_positive_gap, XYCUT_MAX_OVERLAP_EMU)

        left_height = (
            slide6["C6 LEFT TOP — h=180"]["bbox"][3] - slide6["C6 LEFT TOP — h=180"]["bbox"][1]
        )
        right_height = (
            slide6["C6 RIGHT TOP — h=70"]["bbox"][3] - slide6["C6 RIGHT TOP — h=70"]["bbox"][1]
        )
        self.assertEqual(left_height, 180 * XYCUT_MAX_OVERLAP_EMU)
        self.assertEqual(right_height, 70 * XYCUT_MAX_OVERLAP_EMU)

    def test_actual_pptx_groups_small_alignment_jitter_in_projection(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            package_dir = Path(temp_dir) / "package"
            with zipfile.ZipFile(PPTX_FIXTURE) as archive:
                archive.extractall(package_dir)

            def case_objects(slide_number: int, prefix: str):
                objects, _ = extract_slide_objects_xml(
                    package_dir / "ppt" / "slides" / f"slide{slide_number}.xml",
                    pptx_inheritance="none",
                    inherited_shapes="none",
                )
                return [obj for obj in objects if obj.text.startswith(prefix)]

            same_row = case_objects(1, "C1 ")
            same_column = case_objects(2, "C2 ")
            unequal_height_top_row = [obj for obj in case_objects(6, "C6 ") if " TOP " in obj.text]

        self.assertEqual(
            _projection_chunks(same_row, "y", max_overlap=XYCUT_MAX_OVERLAP_EMU),
            [],
        )
        self.assertEqual(
            _projection_chunks(same_column, "x", max_overlap=XYCUT_MAX_OVERLAP_EMU),
            [],
        )
        self.assertEqual(
            _projection_chunks(
                unequal_height_top_row,
                "y",
                max_overlap=XYCUT_MAX_OVERLAP_EMU,
            ),
            [],
        )


if __name__ == "__main__":
    unittest.main()
