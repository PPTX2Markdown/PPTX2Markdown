from __future__ import annotations

import io
import json
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path

import pptx2markdown
from pptx2markdown.cli import main
from pptx2markdown.errors import UnsupportedPresentationError

from .pptx_factory import build_test_pptx


class CliTests(unittest.TestCase):
    def test_cli_writes_default_bundle_shape(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = build_test_pptx(root / "deck.pptx")
            output = root / "output"
            exit_code = main([str(source), "--output-dir", str(output)])
            self.assertEqual(exit_code, 0)
            self.assertTrue((output / "deck" / "deck.md").is_file())
            self.assertTrue((output / "deck" / "media" / "image1.png").is_file())


class LockFileTests(unittest.TestCase):
    """PowerPoint writes "~$deck.pptx" while a deck is open; globs pick it up."""

    def setUp(self) -> None:
        self.directory = tempfile.TemporaryDirectory()
        self.root = Path(self.directory.name)
        self.deck = build_test_pptx(self.root / "deck.pptx")
        self.lock = self.root / "~$deck.pptx"
        self.lock.write_bytes(b"\x0bmoonyoung" + b"\x00" * 40)

    def tearDown(self) -> None:
        self.directory.cleanup()

    def test_cli_skips_lock_files_with_a_warning_and_succeeds(self) -> None:
        errors = io.StringIO()
        with redirect_stderr(errors):
            code = main([str(self.deck), str(self.lock), "-o", str(self.root / "out")])
        self.assertEqual(code, 0)
        self.assertIn("lock file", errors.getvalue())
        self.assertEqual([p.name for p in (self.root / "out").iterdir()], ["deck"])

    def test_review_and_audit_skip_them_too(self) -> None:
        for flag in ("--review", "--audit"):
            output = io.StringIO()
            with redirect_stdout(output), redirect_stderr(io.StringIO()):
                code = main([str(self.deck), str(self.lock), flag])
            self.assertEqual(code, 0, flag)
            self.assertEqual(
                [entry["source"] for entry in json.loads(output.getvalue())], ["deck.pptx"]
            )

    def test_api_reports_skipped_lock_files_and_explains_a_named_one(self) -> None:
        result = pptx2markdown.convert_many([self.lock, self.deck], output_dir=self.root / "api")
        self.assertEqual((len(result.succeeded), result.failed), (1, ()))
        self.assertEqual(result.skipped, (self.lock,))
        self.assertEqual(result.exit_code, 0)
        with self.assertRaisesRegex(UnsupportedPresentationError, "lock file"):
            pptx2markdown.parse(self.lock)


if __name__ == "__main__":
    unittest.main()
