from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path

import pptx2markdown
from pptx2markdown.ooxml import _STRICT_NAMESPACES

from .pptx_factory import build_test_pptx


def _to_strict(source: Path, target: Path) -> Path:
    """The same package with every Transitional namespace URI replaced by its Strict one."""

    transitional = {value: key for key, value in _STRICT_NAMESPACES.items()}
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(target, "w") as strict:
        for info in original.infolist():
            data = original.read(info)
            if info.filename.endswith((".xml", ".rels")):
                for uri in sorted(transitional, key=len, reverse=True):
                    data = data.replace(uri, transitional[uri])
            strict.writestr(info, data)
    return target


class StrictOOXMLTests(unittest.TestCase):
    def test_strict_package_parses_like_its_transitional_original(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            original = build_test_pptx(Path(directory) / "transitional.pptx")
            strict = _to_strict(original, Path(directory) / "strict.pptx")
            self.assertIn(b"purl.oclc.org", zipfile.ZipFile(strict).read("ppt/slides/slide1.xml"))
            with pptx2markdown.parse(original) as expected, pptx2markdown.parse(strict) as actual:
                self.assertTrue(expected.document.slides[0].blocks)
                self.assertEqual(
                    actual.document.model_dump(exclude={"source"}),
                    expected.document.model_dump(exclude={"source"}),
                )


if __name__ == "__main__":
    unittest.main()
