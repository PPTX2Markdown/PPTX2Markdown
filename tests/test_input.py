from __future__ import annotations

import io
import tempfile
import unittest
from pathlib import Path

import pptx2markdown
from pptx2markdown.errors import PackageSecurityError, UnsupportedPresentationError
from pptx2markdown.ooxml import PackageLimits
from pptx2markdown.parser import ParseOptions, parse_components

from .pptx_factory import build_test_pptx


class InMemoryInputTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.directory = tempfile.TemporaryDirectory()
        cls.path = build_test_pptx(Path(cls.directory.name) / "deck.pptx")
        cls.data = cls.path.read_bytes()

    @classmethod
    def tearDownClass(cls) -> None:
        cls.directory.cleanup()

    def test_bytes_and_streams_parse_like_the_file(self) -> None:
        with pptx2markdown.parse(self.path) as from_file:
            expected = from_file.document
        with pptx2markdown.parse(self.data, name="deck.pptx") as from_bytes:
            self.assertEqual(from_bytes.document, expected)
            self.assertIsNone(from_bytes.source_path)
        with pptx2markdown.parse(io.BytesIO(self.data), name="deck.pptx") as from_stream:
            self.assertEqual(from_stream.document, expected)

    def test_in_memory_input_saves_a_bundle(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            with pptx2markdown.parse(self.data, name="upload.pptx") as presentation:
                result = presentation.save(Path(directory) / "upload")
            self.assertIsNone(result.source)
            self.assertTrue(result.document_path.is_file())
            self.assertEqual(presentation.document.source.name, "upload.pptx")

    def test_name_defaults_and_is_reduced_to_a_basename(self) -> None:
        with pptx2markdown.parse(self.data) as unnamed:
            self.assertEqual(unnamed.document.source.name, "presentation.pptx")
        with pptx2markdown.parse(self.data, name="../../etc/deck.pptx") as named:
            self.assertEqual(named.document.source.name, "deck.pptx")

    def test_bad_input_is_rejected(self) -> None:
        with self.assertRaises(UnsupportedPresentationError):
            pptx2markdown.parse(b"%PDF-1.7 not a deck")
        with self.assertRaises(TypeError):
            pptx2markdown.parse(io.StringIO("text"))
        with self.assertRaises(PackageSecurityError):
            parse_components(
                io.BytesIO(self.data),
                options=ParseOptions(package_limits=PackageLimits(max_package_bytes=100)),
            )


if __name__ == "__main__":
    unittest.main()


class NotAPackageTests(unittest.TestCase):
    def test_compound_file_gets_an_actionable_error(self) -> None:
        from pptx2markdown.errors import PresentationParseError

        ole = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 600
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "locked.pptx"
            path.write_bytes(ole)
            for source in (path, ole):
                with self.subTest(source=type(source).__name__):
                    with self.assertRaisesRegex(PresentationParseError, "password"):
                        pptx2markdown.parse(source)

    def test_small_parts_may_compress_far_beyond_the_ratio_limit(self) -> None:
        import zipfile

        with tempfile.TemporaryDirectory() as directory:
            path = build_test_pptx(Path(directory) / "deck.pptx")
            with zipfile.ZipFile(path, "a", compression=zipfile.ZIP_DEFLATED) as archive:
                archive.writestr("ppt/media/vector.emf", b"\x00" * 1_500_000)
            with pptx2markdown.parse(path) as presentation:
                self.assertTrue(presentation.document.slides)
