from __future__ import annotations

import contextlib
import hashlib
import io
import json
import tempfile
import unittest
from pathlib import Path

import pptx2markdown


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class RepositoryGoldenTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.fixture_dir = Path(__file__).resolve().parent / "fixtures" / "golden"
        cls.manifest = json.loads((cls.fixture_dir / "manifest.json").read_text(encoding="utf-8"))
        cls.cases = cls.manifest["cases"]
        cls.paths = [cls.fixture_dir / case["file"] for case in cls.cases]
        cls.snapshot_dirs = {
            output_format: cls.fixture_dir / relative_path
            for output_format, relative_path in cls.manifest["snapshot_roots"].items()
        }

        expected_files = {case["file"] for case in cls.cases}
        actual_files = {
            path.name for path in cls.fixture_dir.glob("*.pptx") if not path.name.startswith("~$")
        }
        if actual_files != expected_files:
            raise AssertionError(
                f"golden PPTX set differs from manifest: {actual_files ^ expected_files}"
            )
        for case, path in zip(cls.cases, cls.paths):
            if _sha256(path) != case["sha256"]:
                raise AssertionError(f"golden input hash mismatch: {path.name}")
        expected_stems = {Path(case["file"]).stem for case in cls.cases}
        for output_format, extension in (("json", "json"), ("markdown", "md")):
            actual_stems = {
                path.stem for path in cls.snapshot_dirs[output_format].glob(f"*.{extension}")
            }
            if actual_stems != expected_stems:
                raise AssertionError(
                    f"golden {output_format} snapshots differ from manifest: "
                    f"{actual_stems ^ expected_stems}"
                )

        cls.temp_dir = tempfile.TemporaryDirectory()
        cls.temp_root = Path(cls.temp_dir.name)
        cls.outputs: dict[str, Path] = {}
        for output_format in ("json", "markdown"):
            output_dir = cls.temp_root / f"{output_format}-output"
            work_dir = cls.temp_root / f"{output_format}-work"
            with contextlib.redirect_stdout(io.StringIO()):
                exit_code = pptx2markdown.convert(
                    cls.paths,
                    output_dir=output_dir,
                    work_dir=work_dir,
                    output_format=output_format,
                )
            if exit_code != 0:
                raise AssertionError(
                    f"repository golden {output_format} conversion failed: {exit_code}"
                )
            cls.outputs[output_format] = output_dir

        cls.documents = {
            Path(case["file"]).stem: json.loads(
                (
                    cls.outputs["json"]
                    / Path(case["file"]).stem
                    / f"{Path(case['file']).stem}.json"
                ).read_text(encoding="utf-8")
            )
            for case in cls.cases
        }

    @classmethod
    def tearDownClass(cls) -> None:
        if hasattr(cls, "temp_dir"):
            cls.temp_dir.cleanup()

    def test_complete_json_markdown_and_asset_outputs_are_stable(self) -> None:
        for case in self.cases:
            stem = Path(case["file"]).stem
            expected_assets = case["assets"]
            for output_format, extension in (("json", "json"), ("markdown", "md")):
                package_dir = self.outputs[output_format] / stem
                result = package_dir / f"{stem}.{extension}"
                snapshot = self.snapshot_dirs[output_format] / f"{stem}.{extension}"
                self.assertEqual(
                    result.read_bytes(),
                    snapshot.read_bytes(),
                    f"{case['file']} snapshot ({output_format})",
                )
                self.assertEqual(
                    _sha256(result),
                    case[f"{output_format}_sha256"],
                    f"{case['file']} ({output_format})",
                )
                actual_assets = {
                    path.relative_to(package_dir).as_posix(): _sha256(path)
                    for path in package_dir.rglob("*")
                    if path.is_file() and path != result
                }
                self.assertEqual(
                    actual_assets,
                    expected_assets,
                    f"{case['file']} assets ({output_format})",
                )

    def test_batch_conversion_summary_is_stable_in_both_formats(self) -> None:
        for output_format, output_dir in self.outputs.items():
            conversion_manifest = json.loads(
                (output_dir / "convert_manifest.json").read_text(encoding="utf-8")
            )
            self.assertEqual(
                conversion_manifest["summary"],
                self.manifest["expected_summary"],
                output_format,
            )

    def test_xycut_coordinate_policy_is_explicit(self) -> None:
        document = self.documents["synthetic_reading_order"]
        actual = {
            slide["page"]: [block["content"] for block in slide["blocks"]]
            for slide in document["slides"]
        }
        self.assertEqual(
            actual,
            {
                1: ["Same row: y differs by 1–3 px", "row-y0", "row-y1", "row-y3"],
                2: [
                    "Same column: x differs by 1–3 px",
                    "column-x0",
                    "column-x1",
                    "column-x3",
                ],
                3: [
                    "Columns separated despite a 1 px overlap",
                    "overlap1-left-top",
                    "overlap1-left-bottom",
                    "overlap1-right-top",
                    "overlap1-right-bottom",
                ],
                4: [
                    "A 24 px overlap forces row-major reading",
                    "overlap24-left-top",
                    "overlap24-right-top",
                    "overlap24-left-bottom",
                    "overlap24-right-bottom",
                ],
                5: [
                    "A 0.5 px positive gutter still defines columns",
                    "gutter-left-top",
                    "gutter-left-bottom",
                    "gutter-right-top",
                    "gutter-right-bottom",
                ],
                6: [
                    "Different heights remain on the same visual row",
                    "height-short",
                    "height-tall",
                    "height-medium",
                ],
                7: [
                    "Full-width regions wrap two independent columns",
                    "region-intro",
                    "region-left-1",
                    "region-right-1",
                    "region-left-2",
                    "region-right-2",
                    "region-conclusion",
                ],
                8: [
                    "Empty and whitespace-only shapes do not create blocks",
                    "visible-after-noise",
                ],
            },
        )

    def test_text_object_policy_and_speaker_notes_are_explicit(self) -> None:
        document = self.documents["synthetic_text_objects"]
        slides = {slide["page"]: slide for slide in document["slides"]}

        self.assertEqual(
            [block["content"] for block in slides[2]["blocks"]][1:],
            ["0", "1", "2", "3", "4"],
        )
        self.assertEqual(slides[6]["notes"], "speaker-note-line-one\nspeaker-note-line-two")
        self.assertEqual(
            [block["content"] for block in slides[6]["blocks"]],
            ["Speaker notes remain separate from visible blocks", "visible-slide-body"],
        )
        self.assertIn("[example.com](https://example.com/docs)", slides[1]["blocks"][1]["content"])
        self.assertIn("| merged body |  | tail |", slides[4]["blocks"][1]["content"])
        self.assertIn("**Chart type:** Column Chart", slides[5]["blocks"][1]["content"])


if __name__ == "__main__":
    unittest.main()
