from __future__ import annotations

import filecmp
import pickle
import tempfile
import unittest
from pathlib import Path

import pptx2markdown
from pptx2markdown.errors import PresentationParseError

from .pptx_factory import build_test_pptx


def _same_tree(left: Path, right: Path) -> bool:
    comparison = filecmp.dircmp(left, right)
    if comparison.left_only or comparison.right_only or comparison.diff_files:
        return False
    return (
        all(_same_tree(left / name, right / name) for name in comparison.common_dirs)
        and not filecmp.cmpfiles(left, right, comparison.common_files, shallow=False)[1]
    )


class ParallelBatchTests(unittest.TestCase):
    def test_jobs_give_the_same_bundles_results_and_errors(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            decks = [build_test_pptx(root / f"in{i}" / "deck.pptx") for i in range(3)]
            broken = root / "broken.pptx"
            broken.write_bytes(b"PK\x03\x04 not really a zip")
            inputs = [decks[0], broken, decks[1], decks[2]]
            serial = pptx2markdown.convert_many(inputs, output_dir=root / "one")
            parallel = pptx2markdown.convert_many(inputs, output_dir=root / "two", jobs=2)
            self.assertEqual(
                [r.bundle_dir.name for r in serial.succeeded],
                [r.bundle_dir.name for r in parallel.succeeded],
            )
            self.assertEqual([f.source for f in parallel.failed], [broken])
            self.assertIs(type(parallel.failed[0].error), type(serial.failed[0].error))
            self.assertEqual(str(parallel.failed[0].error), str(serial.failed[0].error))
            self.assertTrue(_same_tree(root / "one", root / "two"))

    def test_jobs_must_be_positive(self) -> None:
        with self.assertRaises(ValueError):
            pptx2markdown.convert_many([], jobs=0)

    def test_parse_errors_survive_pickling(self) -> None:
        error = PresentationParseError("deck.pptx", "bad table", slide_number=3)
        copy = pickle.loads(pickle.dumps(error))
        self.assertEqual((str(copy), copy.slide_number, copy.detail), (str(error), 3, "bad table"))


if __name__ == "__main__":
    unittest.main()
