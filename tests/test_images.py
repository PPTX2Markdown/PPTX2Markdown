from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pptx2markdown
from pptx2markdown import ImageDescription, ImageRequest, describe_images_with
from pptx2markdown.model import ImageBlock

from .pptx_factory import SLIDE, build_test_pptx

# The factory's picture with no description and a default name, which review flags.
UNDESCRIBED = SLIDE.replace('name="Picture 1" descr="Sales chart"', 'name="Picture 1"')


class DescribeImagesWithTests(unittest.TestCase):
    def setUp(self) -> None:
        self.directory = tempfile.TemporaryDirectory()
        self.source = build_test_pptx(Path(self.directory.name) / "deck.pptx", UNDESCRIBED)
        self.requests: list[ImageRequest] = []

    def tearDown(self) -> None:
        self.directory.cleanup()

    def _image_id(self, presentation: pptx2markdown.ParsedPresentation) -> str:
        (block,) = [b for b in presentation.document.slides[0].blocks if isinstance(b, ImageBlock)]
        assert block.id is not None
        return block.id

    def test_the_describer_sees_the_picture_and_its_slide(self) -> None:
        def describe(request: ImageRequest) -> ImageDescription:
            self.requests.append(request)
            return ImageDescription(description="Bar chart of revenue by region")

        with pptx2markdown.parse(self.source) as presentation:
            recipe = describe_images_with(presentation, describe)
            block_id = self._image_id(presentation)
            presentation.apply(recipe)
            (image,) = [
                b for b in presentation.document.slides[0].blocks if isinstance(b, ImageBlock)
            ]
        (request,) = self.requests
        self.assertEqual((request.block_id, request.page), (block_id, 1))
        self.assertEqual(request.media_type, "image/png")
        self.assertTrue(request.data.startswith(b"\x89PNG"))
        self.assertIn("Quarterly Report", request.slide_markdown)
        self.assertEqual(image.description, "Bar chart of revenue by region")

    def test_decorative_pictures_are_excluded_and_unanswered_ones_left_alone(self) -> None:
        with pptx2markdown.parse(self.source) as presentation:
            block_id = self._image_id(presentation)
            decorative = describe_images_with(
                presentation, lambda request: ImageDescription(decorative=True)
            )
            unanswered = describe_images_with(presentation, lambda request: None)
        self.assertEqual([(r.target, r.exclude) for r in decorative.rules], [(block_id, True)])
        self.assertEqual(unanswered.rules, ())

    def test_media_type_filter_skips_without_calling_the_describer(self) -> None:
        with pptx2markdown.parse(self.source) as presentation:
            recipe = describe_images_with(
                presentation,
                lambda request: self.fail("must not be called"),
                media_types={"image/jpeg"},
            )
        self.assertEqual(recipe.rules, ())


if __name__ == "__main__":
    unittest.main()
