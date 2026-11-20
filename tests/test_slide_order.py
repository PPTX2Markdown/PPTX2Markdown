"""Slides follow the presentation's slide list, not their part names."""

from __future__ import annotations

import json
import re
import tempfile
import unittest
import zipfile
from pathlib import Path

import pptx2markdown

FIXTURE = Path(__file__).parent / "fixtures" / "xycut_layout_cases.pptx"


def _reordered_copy(destination: Path) -> list[str]:
    """Copy the fixture with its slide list reversed and its first slide left out."""
    with zipfile.ZipFile(FIXTURE) as source, zipfile.ZipFile(destination, "w") as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "ppt/presentation.xml":
                text = data.decode("utf-8")
                entries = re.findall(r"<p:sldId [^>]*/>", text)
                kept = list(reversed(entries[1:]))
                block = re.search(r"<p:sldIdLst>.*?</p:sldIdLst>", text, re.S).group(0)
                text = text.replace(block, "<p:sldIdLst>" + "".join(kept) + "</p:sldIdLst>")
                data = text.encode("utf-8")
            target.writestr(info, data)
    return [f"L{number} TITLE" for number in range(len(entries), 1, -1)]


class SlideOrderTests(unittest.TestCase):
    def test_slides_follow_the_slide_list_and_unlisted_parts_are_skipped(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "reordered.pptx"
            expected_titles = _reordered_copy(source)
            exit_code = pptx2markdown.convert(
                source, output_dir=root / "out", work_dir=root / "work", output_format="json"
            )
            document = json.loads(
                (root / "out" / "reordered" / "reordered.json").read_text(encoding="utf-8")
            )

        self.assertEqual(exit_code, 0)
        self.assertEqual([slide["page"] for slide in document["slides"]], [1, 2, 3, 4, 5])
        titles = [slide["blocks"][0]["content"] for slide in document["slides"]]
        self.assertEqual(
            [title[: len(prefix)] for title, prefix in zip(titles, expected_titles)],
            expected_titles,
        )


if __name__ == "__main__":
    unittest.main()
