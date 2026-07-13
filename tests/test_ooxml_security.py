from __future__ import annotations

import stat
import tempfile
import unittest
import zipfile
from pathlib import Path

from pptx2markdown.ooxml_security import (
    ArchiveSafetyLimits,
    resolve_package_part,
    resolve_relationship_target,
    safe_extract_ooxml_archive,
)


class OoxmlArchiveSecurityTests(unittest.TestCase):
    def _extract(self, entries: list[tuple[str, bytes]], limits=None) -> Path:
        temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(temp_dir.cleanup)
        root = Path(temp_dir.name)
        source = root / "input.pptx"
        with zipfile.ZipFile(source, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for name, payload in entries:
                archive.writestr(name, payload)
        destination = root / "output"
        kwargs = {"limits": limits} if limits is not None else {}
        safe_extract_ooxml_archive(source, destination, **kwargs)
        return destination

    def test_extracts_normal_package(self) -> None:
        destination = self._extract(
            [
                ("[Content_Types].xml", b"<Types/>"),
                ("ppt/slides/slide1.xml", b"<p:sld/>"),
            ]
        )
        self.assertEqual((destination / "ppt/slides/slide1.xml").read_bytes(), b"<p:sld/>")

    def test_rejects_parent_traversal_and_windows_separator_variants(self) -> None:
        for malicious_name in ("../escaped.txt", "ppt\\..\\..\\escaped.txt"):
            with self.subTest(malicious_name=malicious_name):
                with self.assertRaisesRegex(ValueError, "unsafe archive entry"):
                    self._extract([(malicious_name, b"owned")])

    def test_rejects_absolute_and_drive_qualified_paths(self) -> None:
        for malicious_name in ("/tmp/escaped.txt", "C:/escaped.txt"):
            with self.subTest(malicious_name=malicious_name):
                with self.assertRaisesRegex(ValueError, "unsafe archive entry"):
                    self._extract([(malicious_name, b"owned")])

    def test_rejects_symbolic_link_entries(self) -> None:
        temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(temp_dir.cleanup)
        root = Path(temp_dir.name)
        source = root / "symlink.pptx"
        link = zipfile.ZipInfo("ppt/media/link")
        link.create_system = 3
        link.external_attr = (stat.S_IFLNK | 0o777) << 16
        with zipfile.ZipFile(source, "w") as archive:
            archive.writestr(link, "../../../outside")

        with self.assertRaisesRegex(ValueError, "symbolic-link"):
            safe_extract_ooxml_archive(source, root / "output")

    def test_rejects_entry_count_and_uncompressed_size_limits(self) -> None:
        count_limits = ArchiveSafetyLimits(max_entries=1)
        with self.assertRaisesRegex(ValueError, "too many entries"):
            self._extract([("one", b"1"), ("two", b"2")], count_limits)

        size_limits = ArchiveSafetyLimits(
            max_member_uncompressed_bytes=3,
            max_total_uncompressed_bytes=10,
        )
        with self.assertRaisesRegex(ValueError, "uncompressed size limit"):
            self._extract([("large.xml", b"1234")], size_limits)

    def test_rejects_extreme_compression_ratio(self) -> None:
        limits = ArchiveSafetyLimits(
            max_member_uncompressed_bytes=10_000,
            max_total_uncompressed_bytes=10_000,
            max_compression_ratio=2.0,
            compression_ratio_min_bytes=100,
        )
        with self.assertRaisesRegex(ValueError, "compression-ratio"):
            self._extract([("bomb.xml", b"A" * 1_000)], limits)


class OoxmlRelationshipSecurityTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        temp_root = Path(self.temp_dir.name)
        self.package = temp_root / "package"
        self.package.mkdir()
        (self.package / "[Content_Types].xml").write_text("<Types/>", encoding="utf-8")
        self.slide = self.package / "ppt/slides/slide1.xml"
        self.rels = self.package / "ppt/slides/_rels/slide1.xml.rels"
        self.media = self.package / "ppt/media/image1.png"
        self.slide.parent.mkdir(parents=True)
        self.rels.parent.mkdir(parents=True)
        self.media.parent.mkdir(parents=True)
        self.slide.write_text("<p:sld/>", encoding="utf-8")
        self.rels.write_text("<Relationships/>", encoding="utf-8")
        self.media.write_bytes(b"png")
        self.outside = temp_root / "outside.txt"
        self.outside.write_text("secret", encoding="utf-8")

    def test_resolves_normal_relative_and_package_absolute_targets(self) -> None:
        self.assertEqual(
            resolve_package_part(self.slide, "../media/image1.png"), self.media.resolve()
        )
        self.assertEqual(
            resolve_relationship_target(self.rels, "/ppt/media/image1.png"),
            self.media.resolve(),
        )

    def test_rejects_relationship_target_that_escapes_package(self) -> None:
        for target in (
            "../../../outside.txt",
            "%2e%2e/%2e%2e/%2e%2e/outside.txt",
            "file:///etc/passwd",
            "https://example.com/image.png",
        ):
            with self.subTest(target=target):
                self.assertIsNone(resolve_relationship_target(self.rels, target))

    def test_allows_only_matching_generated_structure_analysis_sidecar(self) -> None:
        work_root = Path(self.temp_dir.name) / "work"
        extracted = work_root / "target_slides/sample"
        extracted.mkdir(parents=True)
        (extracted / "[Content_Types].xml").write_text("<Types/>", encoding="utf-8")
        media = extracted / "ppt/media/image.png"
        media.parent.mkdir(parents=True)
        media.write_bytes(b"png")

        rels = work_root / "structure_analysis/sample/_rels/slide1.xml.rels"
        rels.parent.mkdir(parents=True)
        rels.write_text("<Relationships/>", encoding="utf-8")
        target = "../../target_slides/sample/ppt/media/image.png"
        self.assertEqual(resolve_relationship_target(rels, target), media.resolve())

        other_rels = work_root / "structure_analysis/other/_rels/slide1.xml.rels"
        other_rels.parent.mkdir(parents=True)
        other_rels.write_text("<Relationships/>", encoding="utf-8")
        self.assertIsNone(resolve_relationship_target(other_rels, target))


if __name__ == "__main__":
    unittest.main()
