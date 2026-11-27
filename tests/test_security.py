from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path

from pptx2markdown.errors import PackageSecurityError
from pptx2markdown.ooxml import OOXMLPackage, Relationship, resolve_relationship_target

from .pptx_factory import CONTENT_TYPES, PRESENTATION


class SecurityTests(unittest.TestCase):
    def test_archive_traversal_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "unsafe.pptx"
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("[Content_Types].xml", CONTENT_TYPES)
                archive.writestr("ppt/presentation.xml", PRESENTATION)
                archive.writestr("../escape.bin", b"bad")
            with self.assertRaises(PackageSecurityError):
                OOXMLPackage(source).open()

    def test_relationship_traversal_is_rejected(self) -> None:
        relationship = Relationship(
            id="rId1",
            type="image",
            target="../../../outside.png",
        )
        with self.assertRaises(PackageSecurityError):
            resolve_relationship_target("ppt/slides/slide1.xml", relationship)

    def test_external_relationship_is_not_resolved_as_a_part(self) -> None:
        relationship = Relationship(
            id="rId1",
            type="hyperlink",
            target="https://example.com",
            external=True,
        )
        self.assertIsNone(resolve_relationship_target("ppt/slides/slide1.xml", relationship))

    def test_xml_dtd_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "unsafe-xml.pptx"
            malicious = (
                b'<?xml version="1.0"?>'
                b'<!DOCTYPE p:presentation [<!ENTITY x "boom">]>'
                b'<p:presentation xmlns:p="'
                b"http://schemas.openxmlformats.org/presentationml/2006/main"
                b'">&x;</p:presentation>'
            )
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("[Content_Types].xml", CONTENT_TYPES)
                archive.writestr("ppt/presentation.xml", malicious)
            with OOXMLPackage(source) as package:
                with self.assertRaises(PackageSecurityError):
                    package.xml("ppt/presentation.xml")


if __name__ == "__main__":
    unittest.main()
