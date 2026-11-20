from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from pptx2markdown.cli import main

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


if __name__ == "__main__":
    unittest.main()
