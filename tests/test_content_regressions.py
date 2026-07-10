from __future__ import annotations

import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path

from pptx2markdown.main_converter.run_pptx_to_markdown import (
    _ooxml_part_from_relationship,
)
from pptx2markdown.pptx_inheritance.resolver import _is_visible_materialized_shape
from pptx2markdown.structure_analyzer.constants import NS
from pptx2markdown.structure_analyzer.pipeline import materialize_tree_by_objects
from pptx2markdown.structure_analyzer.structure import SlideObject
from pptx2markdown.structure_analyzer.text_rules import is_numbered_heading_text


def _slide_object(**overrides: object) -> SlideObject:
    values: dict[str, object] = {
        "shape_id": "master:7",
        "xml_index": 2_000_001,
        "tag": "sp",
        "name": "Master text",
        "ph_type": None,
        "ph_idx": None,
        "x": 100,
        "y": 200,
        "coord_source": "master",
        "text": "Inherited text",
        "normalized": "inherited text",
        "is_footer": False,
        "is_decorative": False,
        "is_heading": False,
        "is_title_placeholder": False,
        "font_pt": 18.0,
        "bbox": (100, 200, 1100, 700),
        "source_part": "master",
        "inheritance_kind": "materialized",
    }
    values.update(overrides)
    return SlideObject(**values)


def _empty_slide_tree() -> ET.ElementTree:
    xml = f"""
    <p:sld xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">
      <p:cSld>
        <p:spTree>
          <p:nvGrpSpPr/>
          <p:grpSpPr/>
        </p:spTree>
      </p:cSld>
    </p:sld>
    """
    return ET.ElementTree(ET.fromstring(xml))


class ContentRegressionTests(unittest.TestCase):
    def test_relationship_target_recovers_part_inside_reordered_path(self) -> None:
        rels_path = Path("work/structure/sample/_rels/slide9.reordered.xml.rels")
        target = "../../target_slides/sample/ppt/charts/chart1.xml"

        self.assertEqual(
            _ooxml_part_from_relationship(rels_path, target),
            "ppt/charts/chart1.xml",
        )

    def test_master_only_text_is_inserted_into_empty_slide(self) -> None:
        master_xml = f"""
        <p:sldMaster xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">
          <p:cSld>
            <p:spTree>
              <p:nvGrpSpPr/>
              <p:grpSpPr/>
              <p:sp>
                <p:nvSpPr><p:cNvPr id="7" name="Master text"/></p:nvSpPr>
                <p:spPr/>
                <p:txBody><a:p><a:r><a:t>Inherited text</a:t></a:r></a:p></p:txBody>
              </p:sp>
            </p:spTree>
          </p:cSld>
        </p:sldMaster>
        """

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source_slide = root / "slide.xml"
            source_master = root / "master.xml"
            source_slide.write_text(ET.tostring(_empty_slide_tree().getroot(), encoding="unicode"))
            source_master.write_text(master_xml)
            tree = _empty_slide_tree()

            materialize_tree_by_objects(
                tree,
                [_slide_object()],
                {"master_xml": str(source_master)},
                root,
                root / "_rels" / "slide.reordered.xml.rels",
                source_slide,
            )

        texts = [node.text for node in tree.findall(".//a:t", NS)]
        self.assertEqual(texts, ["Inherited text"])

    def test_thin_master_picture_is_filtered_as_decoration(self) -> None:
        thin_picture = ET.fromstring(
            f"""
            <p:pic xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">
              <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10000" cy="100"/></a:xfrm></p:spPr>
            </p:pic>
            """
        )
        regular_picture = ET.fromstring(
            f"""
            <p:pic xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">
              <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10000" cy="5000"/></a:xfrm></p:spPr>
            </p:pic>
            """
        )

        self.assertFalse(_is_visible_materialized_shape(thin_picture, "pic", None))
        self.assertTrue(_is_visible_materialized_shape(regular_picture, "pic", None))

    def test_numbered_heading_accepts_space_before_period(self) -> None:
        self.assertTrue(is_numbered_heading_text("3 . Text Box Heading"))
        self.assertTrue(is_numbered_heading_text("2 . 1 Nested Heading"))
        self.assertFalse(is_numbered_heading_text("3 items in a sentence"))


if __name__ == "__main__":
    unittest.main()
