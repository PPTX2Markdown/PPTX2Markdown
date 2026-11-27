from __future__ import annotations

import mimetypes
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from pptx2markdown.assets import AssetLimits, AssetStore, media_type_for, safe_basename
from pptx2markdown.errors import AssetLimitError, AssetUnavailableError


class MediaTypeTests(unittest.TestCase):
    def test_media_types_do_not_depend_on_the_platform_registry(self) -> None:
        names = ("model3d1.glb", "clip.mov", "notes.txt", "deck.ppt", "unknown.xyz")
        expected = [media_type_for(name) for name in names]
        with mock.patch.object(mimetypes, "guess_type", return_value=("x/changed", None)):
            self.assertEqual([media_type_for(name) for name in names], expected)
        self.assertEqual(expected[0], "model/gltf-binary")
        self.assertEqual(expected[-1], "application/octet-stream")


class AssetStoreTests(unittest.TestCase):
    def test_stable_ids_and_payload_deduplication(self) -> None:
        store = AssetStore()
        first = store.register_bytes(
            b"same",
            logical_key="ppt/media/image1.png",
            kind="image",
            original_name="image1.png",
            source_part="ppt/media/image1.png",
        )
        second = store.register_bytes(
            b"same",
            logical_key="ppt/media/image2.png",
            kind="image",
            original_name="image2.png",
            source_part="ppt/media/image2.png",
        )

        self.assertNotEqual(first.id, second.id)
        self.assertEqual(first.sha256, second.sha256)
        self.assertEqual(store.read_bytes(first.id), b"same")
        self.assertEqual(store.read_bytes(second.id), b"same")
        self.assertEqual(
            tuple(asset.id for asset in store.descriptors()), tuple(sorted((first.id, second.id)))
        )
        store.close()

    def test_payload_omission_retains_metadata(self) -> None:
        store = AssetStore()
        descriptor = store.register_bytes(
            b"payload",
            logical_key="ppt/media/image.png",
            kind="image",
            original_name="image.png",
            source_part="ppt/media/image.png",
            preserve_payload=False,
        )
        self.assertFalse(descriptor.payload_available)
        with self.assertRaises(AssetUnavailableError):
            store.read_bytes(descriptor.id)
        store.close()

    def test_limits_are_enforced(self) -> None:
        store = AssetStore(
            AssetLimits(
                max_assets=1,
                max_single_asset_bytes=4,
                max_total_asset_bytes=4,
                spool_threshold_bytes=2,
            )
        )
        with self.assertRaises(AssetLimitError):
            store.register_bytes(
                b"12345",
                logical_key="ppt/media/large.bin",
                kind="attachment",
                original_name="large.bin",
                source_part="ppt/media/large.bin",
            )
        store.close()

    def test_materialization_and_close(self) -> None:
        store = AssetStore()
        descriptor = store.register_bytes(
            b"asset",
            logical_key="ppt/media/image.png",
            kind="image",
            original_name="image.png",
            source_part="ppt/media/image.png",
        )
        with tempfile.TemporaryDirectory() as temporary:
            target = Path(temporary) / "image.png"
            store.materialize(descriptor.id, target)
            self.assertEqual(target.read_bytes(), b"asset")
        store.close()
        with self.assertRaises(AssetUnavailableError):
            store.read_bytes(descriptor.id)

    def test_safe_basename(self) -> None:
        self.assertEqual(safe_basename("../../image.png", fallback="asset.bin"), "image.png")
        self.assertEqual(safe_basename("CON", fallback="asset.bin"), "_CON")


if __name__ == "__main__":
    unittest.main()
