"""A table without rows is skipped; it does not take its slide down."""

from __future__ import annotations

import json
import re
import tempfile
import unittest
import zipfile
from pathlib import Path

import pptx2markdown

FIXTURE = Path(__file__).parent / "fixtures" / "xycut_layout_cases.pptx"


class EmptyTableTests(unittest.TestCase):
    def test_row_less_table_is_skipped_and_the_slide_kept(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "deck.pptx"
            with zipfile.ZipFile(FIXTURE) as original, zipfile.ZipFile(source, "w") as copy:
                for info in original.infolist():
                    data = original.read(info.filename)
                    if info.filename == "ppt/slides/slide6.xml":
                        text = re.sub(r"<a:tr[ >].*?</a:tr>", "", data.decode(), flags=re.S)
                        data = text.encode()
                    copy.writestr(info, data)
            exit_code = pptx2markdown.convert(
                source, output_dir=root / "out", work_dir=root / "work", output_format="json"
            )
            document = json.loads((root / "out" / "deck" / "deck.json").read_text("utf-8"))
            manifest = json.loads((root / "out" / "convert_manifest.json").read_text("utf-8"))

        self.assertEqual(exit_code, 0)
        slide = document["slides"][5]
        self.assertEqual([block["kind"] for block in slide["blocks"]], ["heading", "text"])
        self.assertEqual(manifest["summary"]["table_skipped_blocks"], 1)


if __name__ == "__main__":
    unittest.main()
