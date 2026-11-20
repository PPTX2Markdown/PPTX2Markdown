"""Regression tests for data-loss and package-boundary safety."""

from __future__ import annotations

import json
import shutil
import tempfile
import unicodedata
import unittest
import zipfile
from pathlib import Path

import pptx2markdown
from pptx2markdown.main_converter.embedded_attachments import (
    _sanitize_filename,
    _unique_attachment_path,
)
from pptx2markdown.main_converter.package_inputs import prepare_package_inputs
from pptx2markdown.main_converter.run_pptx_to_markdown import extract_speaker_notes
from pptx2markdown.ooxml_security import safe_extract_ooxml_archive

FIXTURES = Path(__file__).parent / "fixtures"
LAYOUT_FIXTURE = FIXTURES / "xycut_layout_cases.pptx"
MEDIA_FIXTURE = FIXTURES / "golden" / "pptx-automizer_media.pptx"
NOTES_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
NOTES_XML = (
    '<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    '<p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>'
    "<p:txBody><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"
    "</p:spTree></p:cSld></p:notes>"
)


class WorkDirDeletionTests(unittest.TestCase):
    def test_dot_only_file_names_never_delete_the_work_dir(self) -> None:
        for name in ("...pptx", "..pptx"):
            with self.subTest(name=name), tempfile.TemporaryDirectory() as tmp:
                root = Path(tmp)
                inputs = root / "in"
                inputs.mkdir()
                deck = inputs / name
                shutil.copy(LAYOUT_FIXTURE, deck)
                work = root / "work"
                (work / "target_slides").mkdir(parents=True)
                keep = work / "user_notes.txt"
                keep.write_text("keep me", encoding="utf-8")
                sibling = work / "target_slides" / "other"
                sibling.mkdir()
                (sibling / ".pptx_source.json").write_text("{}", encoding="utf-8")

                exit_code = pptx2markdown.convert(deck, output_dir=root / "out", work_dir=work)

                self.assertEqual(exit_code, 0)
                self.assertTrue(keep.is_file())
                self.assertTrue(sibling.is_dir())
                self.assertTrue((root / "out" / "presentation" / "presentation.md").is_file())


class SameNameInputTests(unittest.TestCase):
    def test_same_named_decks_in_different_folders_both_convert(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "a").mkdir()
            (root / "b").mkdir()
            shutil.copy(LAYOUT_FIXTURE, root / "a" / "deck.pptx")
            shutil.copy(MEDIA_FIXTURE, root / "b" / "Deck.pptx")
            out = root / "out"

            exit_code = pptx2markdown.convert(
                [root / "a" / "deck.pptx", root / "b" / "Deck.pptx"],
                output_dir=out,
                work_dir=root / "work",
            )
            manifest = json.loads((out / "convert_manifest.json").read_text("utf-8"))
            first = (out / "deck" / "deck.md").read_text("utf-8")
            second = (out / "Deck-2" / "Deck.md").read_text("utf-8")

        self.assertEqual(exit_code, 0)
        self.assertEqual(manifest["summary"]["processed_packages"], 2)
        self.assertEqual([row["name"] for row in manifest["packages"]], ["deck", "Deck-2"])
        self.assertNotEqual(first, second)

    def test_same_file_listed_twice_is_converted_once(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            deck = root / "deck.pptx"
            shutil.copy(LAYOUT_FIXTURE, deck)

            prepared, failures = prepare_package_inputs(root, [str(deck), str(deck)])

        self.assertEqual(failures, [])
        self.assertEqual(len(prepared), 1)


class SpeakerNotesBoundaryTests(unittest.TestCase):
    def _package(self, root: Path, target: str, mode: str = "") -> Path:
        package = root / "work" / "target_slides" / "deck"
        slides = package / "ppt" / "slides"
        (slides / "_rels").mkdir(parents=True)
        (package / "[Content_Types].xml").write_text("<Types/>", encoding="utf-8")
        slide_xml = slides / "slide1.xml"
        slide_xml.write_text("<p:sld xmlns:p='urn:p'/>", encoding="utf-8")
        mode_attr = f' TargetMode="{mode}"' if mode else ""
        (slides / "_rels" / "slide1.xml.rels").write_text(
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
            f'relationships"><Relationship Id="rId1" Type="{NOTES_REL}" '
            f'Target="{target}"{mode_attr}/></Relationships>',
            encoding="utf-8",
        )
        return slide_xml

    def test_notes_target_outside_package_is_ignored(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "secret.xml").write_text(NOTES_XML.format(text="SECRET"), "utf-8")
            slide_xml = self._package(root, "../../../../../secret.xml")

            self.assertIsNone(extract_speaker_notes(slide_xml))

    def test_external_notes_target_is_ignored(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            slide_xml = self._package(root, "file:///etc/passwd", mode="External")

            self.assertIsNone(extract_speaker_notes(slide_xml))


class ArchiveCollisionTests(unittest.TestCase):
    def _extract(self, names: list[str]) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "input.pptx"
            with zipfile.ZipFile(source, "w") as archive:
                for name in names:
                    archive.writestr(name, b"x")
            safe_extract_ooxml_archive(source, Path(tmp) / "out")

    def test_rejects_entries_that_collide_on_case_insensitive_filesystems(self) -> None:
        with self.assertRaisesRegex(ValueError, "duplicate archive entry"):
            self._extract(["ppt/slides/slide1.xml", "PPT/Slides/Slide1.xml"])

    def test_rejects_entries_that_collide_after_unicode_normalization(self) -> None:
        nfc = unicodedata.normalize("NFC", "ppt/media/café.png")
        nfd = unicodedata.normalize("NFD", "ppt/media/café.png")
        with self.assertRaisesRegex(ValueError, "duplicate archive entry"):
            self._extract([nfc, nfd])

    def test_rejects_windows_aliasing_names(self) -> None:
        for name in ("ppt/a:stream.xml", "ppt/slide1.xml.", "ppt/slide1.xml ", "ppt/CON.xml"):
            with self.subTest(name=name), self.assertRaisesRegex(ValueError, "unsafe"):
                self._extract([name])


class AttachmentNameTests(unittest.TestCase):
    def test_case_variants_get_distinct_paths(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            first = directory / "Report.pdf"
            second = _unique_attachment_path(directory, "report.pdf", {first})

        self.assertEqual(second.name, "report-2.pdf")

    def test_existing_files_are_not_overwritten(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            (directory / "report.pdf").write_bytes(b"user file")
            chosen = _unique_attachment_path(directory, "report.pdf", set())

        self.assertEqual(chosen.name, "report-2.pdf")

    def test_long_and_reserved_names_are_made_writable(self) -> None:
        long_name = _sanitize_filename("가" * 200 + ".pdf", "attachment.bin")
        self.assertTrue(long_name.endswith(".pdf"))
        self.assertLessEqual(len(long_name.encode("utf-8")), 200)
        self.assertEqual(_sanitize_filename("CON.pdf", "attachment.bin"), "_CON.pdf")


class StaleOutputTests(unittest.TestCase):
    def test_rerun_removes_previous_media_and_attachments(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            deck = root / "deck.pptx"
            out = root / "out"
            shutil.copy(MEDIA_FIXTURE, deck)
            pptx2markdown.convert(deck, output_dir=out, work_dir=root / "work")
            self.assertTrue(any((out / "deck" / "media").iterdir()))

            shutil.copy(LAYOUT_FIXTURE, deck)
            pptx2markdown.convert(deck, output_dir=out, work_dir=root / "work")

            self.assertFalse((out / "deck" / "media").exists())
            self.assertFalse((out / "deck" / "attachments").exists())
            self.assertTrue((out / "deck" / "deck.md").is_file())

    def test_folder_with_user_files_is_left_alone(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            deck = root / "deck.pptx"
            out = root / "out"
            shutil.copy(LAYOUT_FIXTURE, deck)
            (out / "deck" / "media").mkdir(parents=True)
            (out / "deck" / "media" / "old.png").write_bytes(b"old")
            (out / "deck" / "my_notes.txt").write_text("mine", encoding="utf-8")

            pptx2markdown.convert(deck, output_dir=out, work_dir=root / "work")

            self.assertTrue((out / "deck" / "my_notes.txt").is_file())
            self.assertTrue((out / "deck" / "media" / "old.png").is_file())


if __name__ == "__main__":
    unittest.main()
