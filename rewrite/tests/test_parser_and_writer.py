from __future__ import annotations

import json
import shutil
import tempfile
import unittest
from pathlib import Path

import pptx2markdown
from pptx2markdown.assets import AssetStore
from pptx2markdown.errors import AssetResolutionError, AssetUnavailableError, SaveConflictError
from pptx2markdown.model import (
    AttachmentBlock,
    HeadingBlock,
    ImageBlock,
    ListBlock,
    PresentationDocument,
    SlideDocument,
    SourceDocument,
    TableBlock,
    TextBlock,
)
from pptx2markdown.presentation import ParsedPresentation

from .pptx_factory import build_test_pptx


class ParserAndWriterTests(unittest.TestCase):
    def test_parse_returns_complete_path_independent_document(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            source = build_test_pptx(Path(temporary) / "deck.pptx")
            with pptx2markdown.parse(source) as presentation:
                document = presentation.document
                self.assertEqual(document.schema_version, "2.0")
                self.assertEqual(document.source.name, "deck.pptx")
                self.assertNotIn(str(Path(temporary)), document.model_dump_json())
                self.assertEqual(len(document.slides), 1)
                slide = document.slides[0]
                self.assertEqual(slide.notes, "Presenter note")
                self.assertTrue(any(isinstance(block, HeadingBlock) for block in slide.blocks))
                self.assertTrue(any(isinstance(block, TextBlock) for block in slide.blocks))
                self.assertTrue(any(isinstance(block, ListBlock) for block in slide.blocks))
                self.assertTrue(any(isinstance(block, ImageBlock) for block in slide.blocks))
                self.assertTrue(any(isinstance(block, AttachmentBlock) for block in slide.blocks))
                self.assertTrue(any(isinstance(block, TableBlock) for block in slide.blocks))
                self.assertEqual(len(document.assets), 2)
                self.assertTrue(all(asset.payload_available for asset in document.assets))

    def test_render_requires_explicit_asset_policy(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            source = build_test_pptx(Path(temporary) / "deck.pptx")
            with pptx2markdown.parse(source) as presentation:
                with self.assertRaises(AssetResolutionError):
                    presentation.render_markdown()
                text_only = presentation.render_markdown(assets="omit")
                self.assertNotIn("![", text_only)
                self.assertNotIn("[attachment:", text_only)
                rendered = presentation.render_markdown(
                    asset_url=lambda asset_id: f"/assets/{asset_id}"
                )
                self.assertIn("](/assets/asset-", rendered)
                self.assertIn("[OpenAI](https://openai.com)", rendered)

    def test_save_writes_movable_relative_bundle(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = build_test_pptx(root / "deck.pptx")
            bundle = root / "output" / "deck"
            with pptx2markdown.parse(source) as presentation:
                result = presentation.save(bundle)
            self.assertEqual(result.bundle_dir, bundle.resolve())
            markdown = result.document_path.read_text(encoding="utf-8")
            self.assertIn("![Sales chart](media/image1.png)", markdown)
            self.assertIn("[attachment: media1.mp3](attachments/media1.mp3)", markdown)
            self.assertTrue((bundle / "media" / "image1.png").is_file())
            self.assertTrue((bundle / "attachments" / "media1.mp3").is_file())
            manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
            self.assertEqual(len(manifest["assets"]), 2)

            moved = root / "moved"
            shutil.move(bundle, moved)
            moved_markdown = (moved / "deck.md").read_text(encoding="utf-8")
            self.assertIn("media/image1.png", moved_markdown)
            self.assertTrue((moved / "media" / "image1.png").is_file())

    def test_save_conflict_and_replace(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = build_test_pptx(root / "deck.pptx")
            bundle = root / "output" / "deck"
            with pptx2markdown.parse(source) as presentation:
                presentation.save(bundle)
                with self.assertRaises(SaveConflictError):
                    presentation.save(bundle)
                marker = bundle / "stale.txt"
                marker.write_text("stale", encoding="utf-8")
                presentation.save(bundle, overwrite="replace")
                self.assertFalse((bundle / "stale.txt").exists())

    def test_missing_payload_does_not_leave_partial_bundle(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = build_test_pptx(root / "deck.pptx")
            bundle = root / "output" / "deck"
            with pptx2markdown.parse(source, include_assets=False) as presentation:
                with self.assertRaises(AssetUnavailableError):
                    presentation.save(bundle, assets="copy")
            self.assertFalse(bundle.exists())
            self.assertFalse(any((root / "output").glob(".deck.tmp-*")))

    def test_convert_and_json_bundle(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = build_test_pptx(root / "deck.pptx")
            result = pptx2markdown.convert(
                source,
                output_dir=root / "converted",
                output_format="json",
            )
            payload = json.loads(result.document_path.read_text(encoding="utf-8"))
            self.assertEqual(payload["schema_version"], "2.0")
            self.assertEqual(len(payload["assets"]), 2)

    def test_identical_named_payloads_share_one_bundle_file(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "deck.pptx"
            source.write_bytes(b"synthetic-source")
            store = AssetStore()
            first = store.register_bytes(
                b"same-image",
                logical_key="ppt/media/image.png",
                kind="image",
                original_name="image.png",
                source_part="ppt/media/image.png",
            )
            second = store.register_bytes(
                b"same-image",
                logical_key="ppt/alternate/image.png",
                kind="image",
                original_name="image.png",
                source_part="ppt/alternate/image.png",
            )
            document = PresentationDocument(
                source=SourceDocument(
                    name="deck.pptx",
                    sha256="0" * 64,
                ),
                assets=tuple(sorted((first, second), key=lambda asset: asset.id)),
                slides=(
                    SlideDocument(
                        page=1,
                        blocks=(
                            ImageBlock(asset_id=first.id),
                            ImageBlock(asset_id=second.id),
                        ),
                    ),
                ),
            )
            with ParsedPresentation(
                document=document,
                assets=store,
                source_path=source,
            ) as presentation:
                result = presentation.save(root / "bundle")
            self.assertEqual(len(list((result.bundle_dir / "media").iterdir())), 1)
            markdown = result.document_path.read_text(encoding="utf-8")
            self.assertEqual(markdown.count("media/image.png"), 2)

    def test_batch_disambiguates_duplicate_source_names(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            first = build_test_pptx(root / "a" / "deck.pptx")
            second = build_test_pptx(root / "b" / "deck.pptx")
            result = pptx2markdown.convert_many(
                (first, second),
                output_dir=root / "output",
            )
            self.assertEqual(result.exit_code, 0)
            self.assertEqual(len(result.succeeded), 2)
            bundle_names = {saved.bundle_dir.name for saved in result.succeeded}
            self.assertEqual(len(bundle_names), 2)


if __name__ == "__main__":
    unittest.main()
