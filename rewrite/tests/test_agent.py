from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace

import pptx2markdown
from pptx2markdown.agent import (
    DEFAULT_MODEL,
    AgentError,
    _Answer,
    _Answers,
    _ImageDescription,
    _RecipeDraft,
    _RuleDraft,
    answer_review,
    describe_images,
    document_view,
    draft_recipe,
)
from pptx2markdown.recipe import Recipe, Rule, merge_recipes

from .pptx_factory import build_test_pptx
from .test_review import _document


class FakeClient:
    """Stands in for anthropic.Anthropic: returns queued parsed outputs."""

    def __init__(self, *outputs, stop_reason: str = "end_turn") -> None:
        self.calls: list[dict] = []
        self._outputs = list(outputs)
        self._stop_reason = stop_reason
        self.beta = SimpleNamespace(messages=SimpleNamespace(parse=self._parse))

    def _parse(self, **kwargs):
        self.calls.append(kwargs)
        output = self._outputs.pop(0)
        return SimpleNamespace(
            parsed_output=output,
            stop_reason=self._stop_reason,
            content=[SimpleNamespace(type="text", text=output.model_dump_json())],
        )


def _rule(target: str, **fields) -> _RuleDraft:
    values = dict.fromkeys(_RuleDraft.model_fields, None)
    values.update(target=target, reason="test", **fields)
    return _RuleDraft(**values)


def _draft(*rules: _RuleDraft) -> _RecipeDraft:
    return _RecipeDraft(
        summary="Drops the image.",
        merged_cells=None,
        include_notes=True,
        include_hidden=True,
        rules=list(rules),
    )


class DraftRecipeTests(unittest.TestCase):
    def test_request_becomes_a_validated_recipe(self) -> None:
        client = FakeClient(_draft(_rule("s1.b9", exclude=True)))
        draft = draft_recipe(_document(), "drop the picture", client=client)
        self.assertEqual(draft.recipe.rules, (Rule(target="s1.b9", exclude=True, reason="test"),))
        self.assertEqual(draft.matches[0].block_ids, ("s1.b9",))
        self.assertEqual(draft.summary, "Drops the image.")

        call = client.calls[0]
        self.assertEqual(call["model"], DEFAULT_MODEL)
        self.assertEqual(call["output_format"], _RecipeDraft)
        self.assertEqual(call["fallbacks"], "default")
        context, request = call["messages"][0]["content"]
        self.assertEqual(context["cache_control"], {"type": "ephemeral"})
        self.assertIn("<!-- s1.b6 table -->", context["text"])
        self.assertIn('"topic": "table-header-rows"', context["text"])
        self.assertIn("drop the picture", request["text"])

    def test_a_rule_selecting_nothing_is_sent_back_once(self) -> None:
        client = FakeClient(
            _draft(_rule("s9.b1", exclude=True)), _draft(_rule("slide[1]/image", exclude=True))
        )
        draft = draft_recipe(_document(), "drop the picture", client=client)
        self.assertEqual(draft.unmatched, ())
        feedback = client.calls[1]["messages"]
        self.assertEqual([message["role"] for message in feedback], ["user", "assistant", "user"])
        self.assertIn("'s9.b1' selects nothing", feedback[2]["content"])

    def test_invalid_selector_is_reported_back_and_last_draft_returned(self) -> None:
        client = FakeClient(_draft(_rule("tables[x]")), _draft(_rule("s9.b1", exclude=True)))
        draft = draft_recipe(_document(), "x", client=client)
        self.assertIn("target", client.calls[1]["messages"][2]["content"])
        self.assertEqual(draft.unmatched, ("s9.b1",))

    def test_refusal_is_an_error(self) -> None:
        with self.assertRaises(AgentError):
            draft_recipe(_document(), "x", client=FakeClient(_draft(), stop_reason="refusal"))

    def test_view_annotates_every_block(self) -> None:
        view = document_view(_document())
        self.assertIn("<!-- s1.b9 image -->", view)
        self.assertIn("<!-- s1.b2 text -->", view)


class AnswerReviewTests(unittest.TestCase):
    def test_picked_options_become_rules(self) -> None:
        client = FakeClient(
            _Answers(
                answers=[
                    _Answer(item=0, option=1),
                    _Answer(item=1, option=0),
                    _Answer(item=9, option=0),
                ]
            )
        )
        recipe = answer_review(_document(), client=client)
        self.assertEqual(recipe.rules[0].target, "s1.b2")
        self.assertEqual(recipe.rules[0].heading_level, 2)
        self.assertEqual(len(recipe.rules), 1)


class DescribeImagesTests(unittest.TestCase):
    def test_images_are_sent_with_their_slide_and_described(self) -> None:
        slide = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:pic>
      <p:nvPicPr><p:cNvPr id="4" name="Picture 3"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rIdImage"/></p:blipFill>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>
    </p:pic>
  </p:spTree></p:cSld>
</p:sld>"""
        client = FakeClient(_ImageDescription(decorative=False, description="Sales up 10%"))
        with tempfile.TemporaryDirectory() as directory:
            source = build_test_pptx(Path(directory) / "deck.pptx", slide)
            with pptx2markdown.parse(source) as presentation:
                recipe = describe_images(presentation, client=client, language="Korean")
                image_id = presentation.document.slides[0].blocks[0].id
        self.assertEqual(
            recipe.rules,
            (
                Rule(
                    target=image_id, description="Sales up 10%", reason="described from the image"
                ),
            ),
        )
        image, text = client.calls[0]["messages"][0]["content"]
        self.assertEqual(image["source"]["media_type"], "image/png")
        self.assertIn("Korean", text["text"])


class MergeRecipesTests(unittest.TestCase):
    def test_rules_concatenate_and_settings_combine(self) -> None:
        merged = merge_recipes(
            Recipe(merged_cells="blank", rules=(Rule(target="table", header_rows=2),)),
            Recipe(include_notes=False, rules=(Rule(target="s1.b1", exclude=True),)),
        )
        self.assertEqual(merged.merged_cells, "blank")
        self.assertFalse(merged.include_notes)
        self.assertEqual([rule.target for rule in merged.rules], ["table", "s1.b1"])


if __name__ == "__main__":
    unittest.main()
