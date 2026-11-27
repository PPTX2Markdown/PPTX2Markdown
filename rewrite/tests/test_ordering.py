from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pptx2markdown
from pptx2markdown.libreoffice import find_soffice
from pptx2markdown.ordering import AS_CONVERTED, OrderRequest, choose_reading_orders_with

from .pptx_factory import build_test_pptx

INCH = 914_400


def _text_box(shape_id: int, text: str, x: float, y: float) -> str:
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{shape_id}" name="T{shape_id}"/><p:cNvSpPr/>'
        f'<p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{int(x * INCH)}" y="{int(y * INCH)}"/>'
        f'<a:ext cx="{3 * INCH}" cy="{int(1.2 * INCH)}"/></a:xfrm></p:spPr><p:txBody>'
        f"<a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"
    )


def build_tracks(path: Path) -> Path:
    """Two tracks of three steps, spaced so rows and columns are equally plausible."""

    shapes = "".join(
        _text_box(10 + 3 * row + step, f"{track} step {step}", 0.5 + 3.5 * step, 1 + 2 * row)
        for row, track in enumerate(("Alpha", "Beta"))
        for step in range(3)
    )
    slide = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        f"<p:grpSpPr/>{shapes}</p:spTree></p:cSld></p:sld>"
    )
    return build_test_pptx(path, slide)


class ChooseReadingOrdersTests(unittest.TestCase):
    def setUp(self) -> None:
        self.directory = tempfile.TemporaryDirectory()
        self.source = build_tracks(Path(self.directory.name) / "tracks.pptx")
        self.requests: list[OrderRequest] = []

    def tearDown(self) -> None:
        self.directory.cleanup()

    def judge(self, choice: str | None):
        def pick(request: OrderRequest) -> str | None:
            self.requests.append(request)
            return choice

        return pick

    def test_the_judge_sees_the_candidates_and_its_pick_becomes_a_rule(self) -> None:
        with pptx2markdown.parse(self.source) as presentation:
            recipe = choose_reading_orders_with(
                presentation, self.judge("rows first"), render=False
            )
            presentation.apply(recipe)
            texts = [block.text for block in presentation.document.slides[0].blocks]
        (request,) = self.requests
        self.assertEqual(request.page, 1)
        self.assertIsNone(request.slide_pdf)
        self.assertEqual(list(request.candidates)[0], AS_CONVERTED)
        self.assertEqual(
            request.candidates["rows first"][:3], ["Alpha step 0", "Alpha step 1", "Alpha step 2"]
        )
        self.assertEqual(recipe.rules[0].target, "slide[1]")
        self.assertEqual(texts[:3], ["Alpha step 0", "Alpha step 1", "Alpha step 2"])

    def test_presentation_review_answers_apply_as_a_recipe(self) -> None:
        with pptx2markdown.parse(self.source) as presentation:
            items = presentation.review()
            presentation.apply(
                pptx2markdown.recipe_from_answers(items, {"reading-order@slide[1]": "rows-first"})
            )
            texts = [b.text for b in presentation.document.slides[0].blocks]
        self.assertEqual(texts[:3], ["Alpha step 0", "Alpha step 1", "Alpha step 2"])

    def test_keeping_or_unknown_choices_add_no_rule(self) -> None:
        with pptx2markdown.parse(self.source) as presentation:
            for choice in (None, AS_CONVERTED, "something else"):
                recipe = choose_reading_orders_with(presentation, self.judge(choice), render=False)
                self.assertEqual(recipe.rules, ())

    @unittest.skipUnless(find_soffice(), "LibreOffice is not installed")
    def test_only_flagged_slides_are_rendered_as_one_page_pdfs(self) -> None:
        with pptx2markdown.parse(self.source) as presentation:
            choose_reading_orders_with(presentation, self.judge(None))
        (request,) = self.requests
        self.assertTrue(request.slide_pdf and request.slide_pdf.startswith(b"%PDF"))


if __name__ == "__main__":
    unittest.main()
