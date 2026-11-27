from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

import pptx2markdown

from .pptx_factory import build_test_pptx


class ChunksBundleTests(unittest.TestCase):
    def test_chunks_bundle_has_one_json_line_per_chunk_with_bundle_links(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            result = pptx2markdown.convert(
                build_test_pptx(root / "deck.pptx"),
                output_dir=root / "out",
                output_format="chunks",
            )
            self.assertEqual(result.document_path.name, "deck.chunks.jsonl")
            lines = result.document_path.read_text(encoding="utf-8").splitlines()
            chunks = [json.loads(line) for line in lines]
            self.assertEqual([chunk["metadata"]["page"] for chunk in chunks], [1])
            self.assertEqual(chunks[0]["metadata"]["source"], "deck.pptx")
            images = [asset for asset in result.assets if asset.path.suffix == ".png"]
            relative = images[0].path.relative_to(result.bundle_dir).as_posix()
            self.assertIn(f"]({relative})", chunks[0]["text"])
            manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
            self.assertEqual(manifest["output_format"], "chunks")


if __name__ == "__main__":
    unittest.main()
