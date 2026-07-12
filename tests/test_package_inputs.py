from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from pptx2markdown.main_converter.package_inputs import prepare_package_inputs


class PackageInputTests(unittest.TestCase):
    def test_office_lock_file_is_silently_ignored(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            lock_file = root / "~$sample.pptx"
            lock_file.write_bytes(b"office lock metadata")

            prepared, missing = prepare_package_inputs(root, [str(lock_file)])

        self.assertEqual(prepared, [])
        self.assertEqual(missing, [])

    def test_ignored_lock_file_does_not_hide_genuinely_missing_input(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            lock_file = root / "~$sample.pptx"
            lock_file.write_bytes(b"office lock metadata")
            missing_file = root / "missing.pptx"

            prepared, missing = prepare_package_inputs(
                root,
                [str(lock_file), str(missing_file)],
            )

        self.assertEqual(prepared, [])
        self.assertEqual([row["input"] for row in missing], [str(missing_file)])


if __name__ == "__main__":
    unittest.main()
