from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pptx2markdown
from pptx2markdown.model import ListBlock, ListItem
from pptx2markdown.render import render_block_markdown

from .pptx_factory import build_test_pptx


def _numbered(text: str, *, level: int = 0, start: int | None = None) -> str:
    start_at = f' startAt="{start}"' if start else ""
    return (
        f'<a:p><a:pPr lvl="{level}"><a:buAutoNum type="arabicPeriod"{start_at}/></a:pPr>'
        f"<a:r><a:t>{text}</a:t></a:r></a:p>"
    )


SLIDE = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Steps"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="2000000"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/>
        {_numbered("Collect", start=3)}{_numbered("Clean", level=1)}{_numbered("Merge", level=1)}
        {_numbered("Publish")}
      </p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>
"""


class OrderedListTests(unittest.TestCase):
    def test_start_number_and_nested_numbering(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = build_test_pptx(Path(directory) / "lists.pptx", SLIDE)
            with pptx2markdown.parse(source) as presentation:
                (block,) = presentation.document.slides[0].blocks
        self.assertIsInstance(block, ListBlock)
        self.assertEqual(block.start, 3)
        self.assertEqual(
            render_block_markdown(block, asset_resolver=None, asset_mode="omit").splitlines(),
            ["3. Collect", "    1. Clean", "    2. Merge", "4. Publish"],
        )

    def test_bullets_keep_two_space_nesting(self) -> None:
        block = ListBlock(items=(ListItem(text="a"), ListItem(text="b", level=1)))
        self.assertEqual(
            render_block_markdown(block, asset_resolver=None, asset_mode="omit"), "- a\n  - b"
        )


BULLET = '<a:pPr lvl="1"><a:buChar char="-"/></a:pPr>'


class MixedListTests(unittest.TestCase):
    def test_bullets_under_numbered_items_stay_one_list(self) -> None:
        body = (
            _numbered("Rationale")
            + f"<a:p>{BULLET}<a:r><a:t>Problem</a:t></a:r></a:p>"
            + _numbered("Learn")
            + f"<a:p>{BULLET}<a:r><a:t>Marketing</a:t></a:r></a:p>"
        )
        slide = SLIDE.split("<p:txBody><a:bodyPr/><a:lstStyle/>")[0] + (
            "<p:txBody><a:bodyPr/><a:lstStyle/>"
            + body
            + "</p:txBody></p:sp></p:spTree></p:cSld></p:sld>"
        )
        with tempfile.TemporaryDirectory() as directory:
            source = build_test_pptx(Path(directory) / "mixed.pptx", slide)
            with pptx2markdown.parse(source) as presentation:
                (block,) = presentation.document.slides[0].blocks
        self.assertEqual([item.ordered for item in block.items], [None, False, None, False])
        self.assertEqual(
            render_block_markdown(block, asset_resolver=None, asset_mode="omit").splitlines(),
            ["1. Rationale", "    - Problem", "2. Learn", "    - Marketing"],
        )

    def test_a_list_starting_deep_is_not_a_code_block(self) -> None:
        block = ListBlock(items=(ListItem(text="a", level=3), ListItem(text="b", level=4)))
        self.assertEqual(
            render_block_markdown(block, asset_resolver=None, asset_mode="omit"), "- a\n  - b"
        )


if __name__ == "__main__":
    unittest.main()
