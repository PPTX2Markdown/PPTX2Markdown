"""The repository PPTX suite converts to the accepted Markdown and JSON snapshots."""

from __future__ import annotations

import importlib.util
import unittest
from pathlib import Path

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "update_goldens.py"


@unittest.skipUnless(SCRIPT.is_file(), "the repository scripts are not available")
class GoldenSnapshotTests(unittest.TestCase):
    def test_every_deck_matches_its_snapshots_and_asset_hashes(self) -> None:
        spec = importlib.util.spec_from_file_location("update_goldens", SCRIPT)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        problems = module.check()
        self.assertEqual(
            problems, [], "run `python scripts/update_goldens.py --accept` after review"
        )


if __name__ == "__main__":
    unittest.main()
