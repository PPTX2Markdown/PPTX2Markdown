from __future__ import annotations

import io
import re
import time
import unittest
import zipfile
from pathlib import Path

import pptx2markdown
from pptx2markdown.model import UnsupportedBlock

FIXTURE = (
    Path(__file__).resolve().parents[2]
    / "tests"
    / "fixtures"
    / "golden"
    / "shapecrawler_table-chart-ole.pptx"
)


@unittest.skipUnless(FIXTURE.is_file(), "repository fixtures are not available")
class ChartPointLimitTests(unittest.TestCase):
    def test_forged_point_count_fails_the_chart_fast_and_keeps_the_slide(self) -> None:
        output = io.BytesIO()
        with zipfile.ZipFile(FIXTURE) as source, zipfile.ZipFile(output, "w") as target:
            for info in source.infolist():
                data = source.read(info)
                if re.fullmatch(r"ppt/charts/chart\d+\.xml", info.filename):
                    data = re.sub(rb'ptCount val="\d+"', b'ptCount val="999999999999"', data)
                target.writestr(info, data)
        started = time.perf_counter()
        with pptx2markdown.parse(output.getvalue()) as presentation:
            document = presentation.document
        self.assertLess(time.perf_counter() - started, 10)
        codes = {warning.code for warning in document.warnings}
        self.assertIn("chart-conversion-failed", codes)
        self.assertTrue(
            any(isinstance(b, UnsupportedBlock) for s in document.slides for b in s.blocks)
        )


if __name__ == "__main__":
    unittest.main()
