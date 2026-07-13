from __future__ import annotations

import hashlib
import json
import tempfile
import unittest
from pathlib import Path

import pptx2markdown


class RealWorldGoldenTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.repo_root = Path(__file__).resolve().parents[1]
        expected_path = cls.repo_root / "tests" / "fixtures" / "real_world_cases.expected.json"
        cls.expected = json.loads(expected_path.read_text(encoding="utf-8"))
        cls.corpus_root = cls.repo_root / cls.expected["corpus_root"]
        paths = [cls.corpus_root / case["path"] for case in cls.expected["cases"]]
        missing = [str(path) for path in paths if not path.exists()]
        if missing:
            raise unittest.SkipTest("external real-world PPTX corpus is unavailable")

        for case, path in zip(cls.expected["cases"], paths):
            digest = hashlib.sha256(path.read_bytes()).hexdigest()
            if digest != case["sha256"]:
                raise AssertionError(f"corpus hash mismatch: {path}")

        cls.temp_dir = tempfile.TemporaryDirectory()
        temp_root = Path(cls.temp_dir.name)
        cls.output_dir = temp_root / "output"
        exit_code = pptx2markdown.convert(
            paths,
            output_dir=cls.output_dir,
            work_dir=temp_root / "work",
            output_format="json",
        )
        if exit_code != 0:
            raise AssertionError(f"real-world golden conversion failed: {exit_code}")

        cls.documents = {
            path.name: json.loads(
                (cls.output_dir / path.stem / f"{path.stem}.json").read_text(encoding="utf-8")
            )
            for path in paths
        }

    @classmethod
    def tearDownClass(cls) -> None:
        if hasattr(cls, "temp_dir"):
            cls.temp_dir.cleanup()

    def test_real_world_expected_content_and_geometry(self) -> None:
        for case in self.expected["cases"]:
            document = self.documents[Path(case["path"]).name]
            serialized_document = json.dumps(document, ensure_ascii=False)
            all_blocks = [block for slide in document["slides"] for block in slide["blocks"]]
            all_content = [block["content"] for block in all_blocks]

            if "slide_count" in case:
                self.assertEqual(len(document["slides"]), case["slide_count"], case["name"])

            for forbidden in case.get("forbidden_content", []):
                self.assertNotIn(forbidden, all_content, case["name"])
            for forbidden in case.get("forbidden_document_substrings", []):
                self.assertNotIn(forbidden, serialized_document, case["name"])
            merged_content = "\n".join(all_content)
            for required in case.get("required_content_substrings", []):
                self.assertIn(required, merged_content, case["name"])

            if "hidden_pages" in case:
                hidden_pages = [slide["page"] for slide in document["slides"] if slide["hidden"]]
                self.assertEqual(hidden_pages, case["hidden_pages"], case["name"])
            for page, expected_notes in case.get("notes_by_page", {}).items():
                slide = next(slide for slide in document["slides"] if slide["page"] == int(page))
                self.assertEqual(slide["notes"], expected_notes, case["name"])

            for page, count in case.get("page_block_counts", {}).items():
                slide = next(slide for slide in document["slides"] if slide["page"] == int(page))
                self.assertEqual(len(slide["blocks"]), count, case["name"])
            if case.get("require_all_bbox"):
                self.assertTrue(all(block["bbox"] is not None for block in all_blocks))

            expected_attachments = case.get("attachments")
            if isinstance(expected_attachments, dict):
                attachments_dir = self.output_dir / Path(case["path"]).stem / "attachments"
                actual_names = {path.name for path in attachments_dir.iterdir() if path.is_file()}
                self.assertEqual(actual_names, set(expected_attachments), case["name"])
                for filename, expected_sha256 in expected_attachments.items():
                    attachment = attachments_dir / filename
                    digest = hashlib.sha256(attachment.read_bytes()).hexdigest()
                    self.assertEqual(digest, expected_sha256, case["name"])

            for page, prefixes in case.get("ordered_content_prefixes_by_page", {}).items():
                slide = next(slide for slide in document["slides"] if slide["page"] == int(page))
                actual = [block["content"] for block in slide["blocks"]]
                self.assertEqual(len(actual), len(prefixes), case["name"])
                for content, prefix in zip(actual, prefixes):
                    self.assertTrue(content.startswith(prefix), case["name"])

            if "page" not in case:
                continue
            slide = next(slide for slide in document["slides"] if slide["page"] == case["page"])
            actual = [block["content"] for block in slide["blocks"]]
            if "ordered_content" in case:
                self.assertEqual(actual, case["ordered_content"], case["name"])
            if "ordered_content_prefixes" in case:
                self.assertEqual(len(actual), len(case["ordered_content_prefixes"]), case["name"])
                for content, prefix in zip(actual, case["ordered_content_prefixes"]):
                    self.assertTrue(content.startswith(prefix), case["name"])
            if case.get("require_bbox"):
                self.assertTrue(all(block["bbox"] is not None for block in slide["blocks"]))

    def test_attachment_names_are_stable_when_output_directory_is_reused(self) -> None:
        case = next(
            case
            for case in self.expected["cases"]
            if case["name"].startswith("embedded PDF OLE contents")
        )
        source = self.corpus_root / case["path"]
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            output = root / "output"
            work = root / "work"
            pptx2markdown.convert(source, output_dir=output, work_dir=work, output_format="json")
            document_path = output / source.stem / f"{source.stem}.json"
            first_document = document_path.read_bytes()
            first_names = sorted(
                path.name for path in (output / source.stem / "attachments").iterdir()
            )

            pptx2markdown.convert(source, output_dir=output, work_dir=work, output_format="json")
            second_document = document_path.read_bytes()
            second_names = sorted(
                path.name for path in (output / source.stem / "attachments").iterdir()
            )

        self.assertEqual(second_document, first_document)
        self.assertEqual(second_names, first_names)


if __name__ == "__main__":
    unittest.main()
