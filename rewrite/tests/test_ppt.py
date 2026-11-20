from __future__ import annotations

import os
import stat
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import pptx2markdown
from pptx2markdown.errors import PresentationParseError, UnsupportedPresentationError
from pptx2markdown.libreoffice import find_soffice

from .pptx_factory import NOTES, build_test_pptx

# Stands in for soffice: copies $FAKE_PPTX to <outdir>/<input stem>.pptx, or fails.
FAKE_SOFFICE = """#!/bin/sh
out=""; last=""
while [ $# -gt 0 ]; do
  case "$1" in --outdir) out="$2"; shift ;; *) last="$1" ;; esac
  shift
done
if [ -n "$FAKE_FAIL" ]; then echo "source file could not be loaded" >&2; exit 1; fi
mkdir -p "$out" && cp "$FAKE_PPTX" "$out/$(basename "$last" .ppt).pptx"
"""


class PptInputTests(unittest.TestCase):
    def setUp(self) -> None:
        self.directory = tempfile.TemporaryDirectory()
        self.root = Path(self.directory.name)
        self.pptx = build_test_pptx(self.root / "source.pptx")
        self.ppt = self.root / "legacy.ppt"
        self.ppt.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1not really a ppt")

    def tearDown(self) -> None:
        self.directory.cleanup()

    def test_ppt_needs_the_converter_to_be_asked_for(self) -> None:
        with self.assertRaisesRegex(UnsupportedPresentationError, "--ppt-converter"):
            pptx2markdown.parse(self.ppt)

    @unittest.skipIf(sys.platform.startswith("win"), "the stand-in converter is a shell script")
    def test_converted_bytes_are_parsed_under_the_ppt_name(self) -> None:
        script = self.root / "soffice"
        script.write_text(FAKE_SOFFICE)
        script.chmod(script.stat().st_mode | stat.S_IEXEC)
        environment = {"SOFFICE_PATH": str(script), "FAKE_PPTX": str(self.pptx)}
        with mock.patch.dict(os.environ, environment):
            with pptx2markdown.parse(self.ppt, ppt_converter="libreoffice") as presentation:
                self.assertEqual(presentation.document.source.name, "legacy.pptx")
                self.assertIn("Quarterly Report", presentation.render_markdown(assets="omit"))
            with mock.patch.dict(os.environ, {"FAKE_FAIL": "1"}):
                with self.assertRaisesRegex(PresentationParseError, "could not be loaded"):
                    pptx2markdown.parse(self.ppt, ppt_converter="libreoffice")

    def test_missing_libreoffice_is_a_clear_error(self) -> None:
        with mock.patch("pptx2markdown.libreoffice.find_soffice", return_value=None):
            with self.assertRaisesRegex(UnsupportedPresentationError, "LibreOffice"):
                pptx2markdown.parse(self.ppt, ppt_converter="libreoffice")

    @unittest.skipUnless(find_soffice(), "LibreOffice is not installed")
    def test_real_libreoffice_round_trip(self) -> None:
        import subprocess

        subprocess.run(
            [find_soffice(), "--headless", "--convert-to", "ppt", "--outdir", str(self.root)]
            + [str(self.pptx)],
            check=True,
            capture_output=True,
            timeout=300,
        )
        with pptx2markdown.parse(
            self.root / "source.ppt", ppt_converter="libreoffice"
        ) as presentation:
            markdown = presentation.render_markdown(assets="omit")
        self.assertIn("Quarterly Report", markdown)
        self.assertIn("Revenue grew", markdown)
        self.assertEqual(presentation.document.slides[0].notes, "Presenter note")


class NotesPageNumberTests(unittest.TestCase):
    def test_a_plain_text_box_holding_only_a_number_is_not_a_note(self) -> None:
        page_number = (
            '<p:sp><p:nvSpPr><p:cNvPr id="20" name=""/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>1</a:t></a:r></a:p></p:txBody>"
            "</p:sp>"
        )
        notes = NOTES.replace("</p:spTree>", page_number + "</p:spTree>")
        with tempfile.TemporaryDirectory() as directory:
            source = build_test_pptx(Path(directory) / "deck.pptx", notes_xml=notes)
            with pptx2markdown.parse(source) as presentation:
                self.assertEqual(presentation.document.slides[0].notes, "Presenter note")


if __name__ == "__main__":
    unittest.main()
