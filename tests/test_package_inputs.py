from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

import pptx2markdown
from pptx2markdown.main_converter.package_inputs import (
    OLE_COMPOUND_FILE_SIGNATURE,
    normalize_strict_ooxml_package,
    prepare_package_inputs,
)


class PackageInputTests(unittest.TestCase):
    @staticmethod
    def _minimal_pptx(path: Path) -> None:
        import zipfile

        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("ppt/slides/slide1.xml", "<slide/>")

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

    def test_corrupt_file_is_reported_without_hiding_valid_input(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            corrupt = root / "corrupt.pptx"
            corrupt.write_bytes(b"not a zip")
            valid = root / "valid.pptx"
            self._minimal_pptx(valid)

            prepared, failures = prepare_package_inputs(
                root,
                [str(corrupt), str(valid)],
            )

        self.assertEqual([package.source_pptx_path.name for package in prepared], ["valid.pptx"])
        self.assertEqual([row["input"] for row in failures], [str(corrupt)])
        self.assertEqual(failures[0]["error_type"], "BadZipFile")

    def test_encrypted_ooxml_container_has_actionable_error(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            encrypted = root / "encrypted.pptx"
            encrypted.write_bytes(OLE_COMPOUND_FILE_SIGNATURE + b"encrypted payload")

            prepared, failures = prepare_package_inputs(root, [str(encrypted)])

        self.assertEqual(prepared, [])
        self.assertEqual(failures[0]["error_type"], "EncryptedPresentationError")
        self.assertIn("remove the password", failures[0]["error"])

    def test_strict_ooxml_namespaces_are_normalized_after_extraction(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            slide = root / "ppt" / "slides" / "slide1.xml"
            rels = root / "ppt" / "slides" / "_rels" / "slide1.xml.rels"
            slide.parent.mkdir(parents=True)
            rels.parent.mkdir(parents=True)
            slide.write_text(
                '<p:sld xmlns:p="http://purl.oclc.org/ooxml/presentationml/main" '
                'xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"/>',
                encoding="utf-8",
            )
            rels.write_text(
                '<Relationships xmlns="http://purl.oclc.org/ooxml/package/relationships">'
                '<Relationship Type="http://purl.oclc.org/ooxml/'
                'officeDocument/relationships/slideLayout"/>'
                "</Relationships>",
                encoding="utf-8",
            )

            changed = normalize_strict_ooxml_package(root)
            slide_text = slide.read_text(encoding="utf-8")
            rels_text = rels.read_text(encoding="utf-8")

        self.assertEqual(changed, 2)
        self.assertIn("presentationml/2006/main", slide_text)
        self.assertIn("drawingml/2006/main", slide_text)
        self.assertIn("package/2006/relationships", rels_text)
        self.assertIn("officeDocument/2006/relationships/slideLayout", rels_text)

    def test_mixed_batch_converts_valid_pptx_and_records_corrupt_input(self) -> None:
        fixture = Path(__file__).parent / "fixtures" / "xycut_layout_cases.pptx"
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            corrupt = root / "corrupt.pptx"
            corrupt.write_bytes(b"not a zip")
            output_dir = root / "output"

            exit_code = pptx2markdown.convert(
                [corrupt, fixture],
                output_dir=output_dir,
                work_dir=root / "work",
                output_format="json",
            )
            manifest = json.loads(
                (output_dir / "convert_manifest.json").read_text(encoding="utf-8")
            )

        self.assertEqual(exit_code, 1)
        self.assertEqual(manifest["summary"]["processed_packages"], 1)
        self.assertEqual(manifest["summary"]["processed_slides"], 6)
        self.assertEqual(manifest["summary"]["failed"], 1)
        failed = next(row for row in manifest["packages"] if row.get("status") == "failed")
        self.assertEqual(failed["stage"], "input")
        self.assertEqual(failed["error_type"], "BadZipFile")

    def test_zero_slide_presentation_converts_to_empty_document(self) -> None:
        import zipfile

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "empty.pptx"
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr(
                    "ppt/presentation.xml",
                    '<p:presentation xmlns:p="http://schemas.openxmlformats.org/'
                    'presentationml/2006/main"><p:sldIdLst/></p:presentation>',
                )
            output_dir = root / "output"
            exit_code = pptx2markdown.convert(
                source,
                output_dir=output_dir,
                work_dir=root / "work",
                output_format="json",
            )
            document = json.loads(
                (output_dir / "empty" / "empty.json").read_text(encoding="utf-8")
            )
            manifest = json.loads(
                (output_dir / "convert_manifest.json").read_text(encoding="utf-8")
            )

        self.assertEqual(exit_code, 0)
        self.assertEqual(document["slides"], [])
        self.assertEqual(manifest["summary"]["processed_packages"], 1)
        self.assertEqual(manifest["summary"]["processed_slides"], 0)


if __name__ == "__main__":
    unittest.main()
