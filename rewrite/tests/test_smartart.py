from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pptx2markdown
from pptx2markdown.model import SmartArtBlock

from .pptx_factory import build_test_pptx

GOLDEN = Path(__file__).resolve().parents[2] / "tests" / "fixtures" / "golden"
DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _point(model_id: str, text: str | None, point_type: str = "") -> str:
    kind = f' type="{point_type}"' if point_type else ""
    body = f"<a:r><a:t>{text}</a:t></a:r>" if text else ""
    return (
        f'<dgm:pt modelId="{model_id}"{kind}><dgm:prSet/><dgm:spPr/>'
        f"<dgm:t><a:bodyPr/><a:p>{body}</a:p></dgm:t></dgm:pt>"
    )


def _data_model(points: str, connections: str) -> str:
    return (
        f'<dgm:dataModel xmlns:dgm="{DGM}" xmlns:a="{A}">'
        f"<dgm:ptLst>{points}</dgm:ptLst><dgm:cxnLst>{connections}</dgm:cxnLst>"
        "</dgm:dataModel>"
    )


def _parent_of(model_id: str, source: str, destination: str, order: int) -> str:
    return (
        f'<dgm:cxn modelId="{model_id}" type="parOf" srcId="{source}" destId="{destination}" '
        f'srcOrd="{order}" destOrd="0"/>'
    )


DOC = '<dgm:pt modelId="0" type="doc"><dgm:prSet loCatId="hierarchy"/><dgm:spPr/></dgm:pt>'

HIERARCHY = _data_model(
    DOC + _point("1", "Strategy") + _point("2", "Hire # people") + _point("3", "Ship v2"),
    _parent_of("10", "0", "1", 0) + _parent_of("11", "1", "2", 0) + _parent_of("12", "1", "3", 1),
)
UNLABELED = _data_model(DOC + _point("1", None) + _point("2", None), _parent_of("10", "0", "1", 0))


def _frame(shape_id: int, y: int, rid: str) -> str:
    return f"""
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="{shape_id}" name="Diagram {shape_id}"/>
          <p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="0" y="{y}"/><a:ext cx="5000000" cy="1000000"/></p:xfrm>
        <a:graphic><a:graphicData uri="{DGM}">
          <dgm:relIds xmlns:dgm="{DGM}" r:dm="{rid}"/>
        </a:graphicData></a:graphic>
      </p:graphicFrame>"""


SLIDE = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="{A}"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      {_frame(2, 0, "rId20")}{_frame(3, 2000000, "rId21")}{_frame(4, 4000000, "rId22")}
    </p:spTree>
  </p:cSld>
</p:sld>
"""

DATA_RELATIONSHIP = (
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"'
)
RELATIONSHIPS = (
    f'<Relationship Id="rId20" {DATA_RELATIONSHIP} Target="../diagrams/data1.xml"/>'
    f'<Relationship Id="rId21" {DATA_RELATIONSHIP} Target="../diagrams/data2.xml"/>'
    f'<Relationship Id="rId22" {DATA_RELATIONSHIP} Target="../diagrams/missing.xml"/>'
)


class SmartArtTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = build_test_pptx(
                Path(directory) / "smartart.pptx",
                SLIDE,
                extra_slide_relationships=RELATIONSHIPS,
                extra_parts={
                    "ppt/diagrams/data1.xml": HIERARCHY,
                    "ppt/diagrams/data2.xml": UNLABELED,
                },
            )
            with pptx2markdown.parse(source) as presentation:
                cls.document = presentation.document
        cls.blocks = {block.shape_id: block for block in cls.document.slides[0].blocks}

    def test_hierarchy_is_an_indented_list(self) -> None:
        block = self.blocks["2"]
        self.assertIsInstance(block, SmartArtBlock)
        lines = block.markdown.splitlines()
        self.assertEqual([line.strip() for line in lines][0], "- Strategy")
        children = [line for line in lines if "Hire" in line or "Ship" in line]
        self.assertEqual(len(children), 2)
        self.assertTrue(all(line.startswith((" ", "\t")) for line in children))

    def test_nodes_without_text_are_counted_not_dropped(self) -> None:
        self.assertEqual(self.blocks["3"].markdown, "[smartart: 2 unlabeled nodes]")

    def test_missing_data_part_is_reported(self) -> None:
        self.assertEqual(self.blocks["4"].kind, "unsupported")
        codes = [warning.code for warning in self.document.warnings]
        self.assertIn("smartart-data-unavailable", codes)


@unittest.skipUnless(GOLDEN.is_dir(), "repository fixtures are not available")
class FixtureSmartArtTests(unittest.TestCase):
    def test_unlabeled_diagram_matches_the_legacy_label(self) -> None:
        with pptx2markdown.parse(GOLDEN / "pptx-automizer_smartart.pptx") as presentation:
            markdown = [
                block.markdown
                for block in presentation.document.slides[2].blocks
                if isinstance(block, SmartArtBlock)
            ]
        self.assertEqual(markdown, ["[smartart: 6 unlabeled nodes]"])


if __name__ == "__main__":
    unittest.main()
