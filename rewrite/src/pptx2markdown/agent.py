"""Optional language-model layer: requests, review answers and images become recipes.

The model never touches the parser or the document. It reads an annotated view of
the converted deck and writes a recipe, which is validated, checked against the
document (every rule must select something) and applied deterministically like a
hand-written one. Install with ``pip install "pptx2markdown[agent]"``; credentials
come from the environment as for any Anthropic client.
"""

from __future__ import annotations

import base64
import json
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Any, Literal

from pydantic import BaseModel, ValidationError

from .errors import PPTX2MarkdownError
from .images import ImageDescription, ImageRequest, describe_images_with
from .model import PresentationDocument
from .ordering import AS_CONVERTED, OrderRequest, choose_reading_orders_with
from .recipe import SELECTOR_HELP, Recipe, Rule, RuleMatch, explain_recipe
from .render import annotated_markdown
from .review import ReviewItem, recipe_from_answers, review, review_data

if TYPE_CHECKING:
    from .presentation import ParsedPresentation

DEFAULT_MODEL = "claude-opus-5-5"
# Refusals are re-run on a fallback model chosen by the API instead of failing.
_FALLBACK_BETA = "server-side-fallback-2026-07-01"
_MAX_VIEW_CHARS = 600_000
_VISION_MEDIA_TYPES = {"image/png", "image/jpeg", "image/gif", "image/webp"}
_MAX_IMAGE_BYTES = 5 * 1024 * 1024

Effort = Literal["low", "medium", "high", "xhigh", "max"]


class AgentError(PPTX2MarkdownError):
    """The model could not produce a usable recipe."""


class _RuleDraft(BaseModel):
    target: str
    exclude: bool | None
    merged_cells: Literal["fill", "header", "blank"] | None
    header_rows: int | None
    caption_rows: int | None
    heading_level: int | None
    as_text: bool | None
    description: str | None
    notes: bool | None
    reason: str


class _RecipeDraft(BaseModel):
    summary: str
    merged_cells: Literal["fill", "header", "blank"] | None
    include_notes: bool
    include_hidden: bool
    rules: list[_RuleDraft]


class _Answer(BaseModel):
    item: str
    option: str


class _Answers(BaseModel):
    answers: list[_Answer]


class _OrderChoice(BaseModel):
    reason: str
    choice: str


class _ImageDescription(BaseModel):
    decorative: bool
    description: str | None


@dataclass(frozen=True)
class RecipeDraft:
    """A model-written recipe, what it selects, and the model's summary of it."""

    recipe: Recipe
    summary: str
    matches: tuple[RuleMatch, ...]

    @property
    def unmatched(self) -> tuple[str, ...]:
        return tuple(
            match.target for match in self.matches if not (match.block_ids or match.pages)
        )


_SYSTEM = f"""\
You adjust how a PowerPoint deck was converted to Markdown for retrieval and reading.
You never edit the text yourself. You write a recipe: document-wide settings plus
rules, each selecting blocks or slides and saying what to do with them. A
deterministic program applies it.

The deck is shown as Markdown in which every block is preceded by a comment with its
id and kind, e.g. `<!-- s3.b2 table -->`. Slides start with `[Page_N]`.

Selector syntax: {SELECTOR_HELP}

Rule fields (leave a field null to keep the block as it is):
- exclude: drop the selected blocks, or whole slides for a slide selector.
- merged_cells: tables; "fill" repeats a merged value in every cell it covers,
  "header" only in the header row, "blank" never.
- header_rows: tables; how many rows form the header (rows are combined per column).
- caption_rows: tables; how many leading rows are a title above the table.
- heading_level: 1-6; sets a heading's level or turns a text block into a heading.
- as_text: true turns a heading or list into plain text lines.
- description: images; a description used as alt text.
- notes: slide selectors; false drops the speaker notes, true keeps them.
- reason: one short sentence on why, for the person reviewing the recipe.

Settings: merged_cells (null keeps each table's policy), include_notes,
include_hidden (false drops hidden slides).

Write the fewest rules that do what was asked. Point at single blocks by id, and use
kind or slide selectors only when the request is about all of them. Base every rule
on what the deck shows; do not guess at content you cannot see. If the request
cannot be expressed as a recipe, return no rules and say why in the summary. The
summary tells the user in one or two sentences what the recipe changes.
"""


def document_view(document: PresentationDocument) -> str:
    """The deck as annotated Markdown, the way the model sees it."""

    view = annotated_markdown(document)
    if len(view) > _MAX_VIEW_CHARS:
        raise AgentError(
            f"the deck's Markdown has {len(view)} characters, more than the "
            f"{_MAX_VIEW_CHARS} the agent sends; apply a recipe that drops slides first"
        )
    return view


def _client(client: Any) -> Any:
    if client is not None:
        return client
    try:
        import anthropic
    except ImportError as exc:  # pragma: no cover - depends on the environment
        raise AgentError(
            'the agent needs the anthropic package: pip install "pptx2markdown[agent]"'
        ) from exc
    return anthropic.Anthropic()


def _parse(
    client: Any,
    *,
    model: str,
    effort: Effort,
    system: str,
    messages: list[dict[str, Any]],
    output_format: type[BaseModel],
    max_tokens: int = 16000,
):
    response = client.beta.messages.parse(
        model=model,
        max_tokens=max_tokens,
        system=system,
        messages=messages,
        output_format=output_format,
        output_config={"effort": effort},
        betas=[_FALLBACK_BETA],
        fallbacks="default",
    )
    if response.stop_reason == "refusal":
        raise AgentError("the model declined the request")
    if response.stop_reason == "max_tokens" or response.parsed_output is None:
        raise AgentError(f"the model returned no usable output ({response.stop_reason})")
    return response


def _to_recipe(draft: _RecipeDraft) -> Recipe:
    return Recipe(
        merged_cells=draft.merged_cells,
        include_notes=draft.include_notes,
        include_hidden=draft.include_hidden,
        rules=tuple(
            Rule.model_validate(rule.model_dump(exclude_none=True)) for rule in draft.rules
        ),
    )


def draft_recipe(
    document: PresentationDocument,
    request: str,
    *,
    review_items: Sequence[ReviewItem] | None = None,
    client: Any = None,
    model: str = DEFAULT_MODEL,
    effort: Effort = "medium",
    attempts: int = 2,
) -> RecipeDraft:
    """Turn a request ("drop the appendix, tables keep merged cells blank") into a recipe.

    The deck view and the converter's open questions (``review_items``, computed
    when not given) go first and are cached, so several requests about one deck
    reuse them. A recipe that fails validation or has a rule selecting nothing is
    sent back once with the problem; ``attempts`` bounds the round trips.
    """

    if not request.strip():
        raise ValueError("request must not be empty")
    client = _client(client)
    items = review(document) if review_items is None else tuple(review_items)
    context = (
        f"<deck name={json.dumps(document.source.name)}>\n{document_view(document)}\n</deck>\n\n"
        f"<open_questions>\n{json.dumps(review_data(tuple(items)), ensure_ascii=False)}\n"
        "</open_questions>"
    )
    messages: list[dict[str, Any]] = [
        {
            "role": "user",
            "content": [
                {"type": "text", "text": context, "cache_control": {"type": "ephemeral"}},
                {"type": "text", "text": f"<request>\n{request}\n</request>"},
            ],
        }
    ]
    last: RecipeDraft | None = None
    for _ in range(max(1, attempts)):
        response = _parse(
            client,
            model=model,
            effort=effort,
            system=_SYSTEM,
            messages=messages,
            output_format=_RecipeDraft,
        )
        draft: _RecipeDraft = response.parsed_output
        try:
            recipe = _to_recipe(draft)
        except ValidationError as exc:
            problems = [f"{'.'.join(map(str, e['loc']))}: {e['msg']}" for e in exc.errors()]
        else:
            last = RecipeDraft(recipe, draft.summary, explain_recipe(document, recipe))
            if not last.unmatched:
                return last
            problems = [f"rule target {target!r} selects nothing" for target in last.unmatched]
        messages.append({"role": "assistant", "content": response.content})
        messages.append(
            {
                "role": "user",
                "content": "The recipe cannot be used as is:\n- "
                + "\n- ".join(problems)
                + "\nReturn the corrected recipe.",
            }
        )
    if last is not None:
        return last
    raise AgentError("the model did not produce a valid recipe")


def answer_review(
    document: PresentationDocument,
    items: Sequence[ReviewItem] | None = None,
    *,
    guidance: str = "",
    client: Any = None,
    model: str = DEFAULT_MODEL,
    effort: Effort = "low",
) -> Recipe:
    """Let the model pick an option for each open question; the picks form a recipe.

    Image descriptions need the pictures themselves, so ``describe_images`` handles
    them; here the model may only keep or drop an image.
    """

    items = review(document) if items is None else tuple(items)
    if not items:
        return Recipe()
    client = _client(client)
    response = _parse(
        client,
        model=model,
        effort=effort,
        system=(
            "You check decisions a converter made while turning a PowerPoint deck into "
            "Markdown. For each question, answer with its id and the id of the option "
            "that is right for the deck shown, using the evidence and the surrounding "
            "blocks. When the deck does not settle a question, pick the suggested option."
        ),
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": f"<deck>\n{document_view(document)}\n</deck>",
                        "cache_control": {"type": "ephemeral"},
                    },
                    {
                        "type": "text",
                        "text": "<questions>\n"
                        f"{json.dumps(review_data(items), ensure_ascii=False)}\n"
                        f"</questions>\n<guidance>\n{guidance}\n</guidance>",
                    },
                ],
            }
        ],
        output_format=_Answers,
    )
    # Answers naming no known item or option are dropped: the converter's decision stands.
    by_id = {item.id: item for item in items}
    answers = {
        answer.item: answer.option
        for answer in response.parsed_output.answers
        if answer.item in by_id and by_id[answer.item].option(answer.option) is not None
    }
    return recipe_from_answers(items, answers)


def describe_images(
    presentation: "ParsedPresentation",
    *,
    block_ids: Sequence[str] | None = None,
    language: str | None = None,
    client: Any = None,
    model: str = DEFAULT_MODEL,
    effort: Effort = "low",
) -> Recipe:
    """Describe images that have none, so retrieval can find what they show.

    Each image is sent with its slide's Markdown for context. Decorative images
    become ``exclude`` rules, the rest ``description`` rules. Only PNG, JPEG, GIF
    and WebP payloads up to 5 MB are sent; others are skipped. To use another
    model, local or not, call ``pptx2markdown.images.describe_images_with``.
    """

    client = _client(client)
    language_line = f" Write it in {language}." if language else ""

    def describe(request: ImageRequest) -> ImageDescription:
        response = _parse(
            client,
            model=model,
            effort=effort,
            system=(
                "You describe pictures from slides so that a search index can find "
                "what they show. A description is one or two sentences of what the "
                "picture conveys: chart type and the trend or values it shows, what a "
                "diagram or screenshot shows, or who and what a photo shows. Mark "
                "logos, backgrounds, icons and other decoration as decorative."
            ),
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": request.media_type,
                                "data": base64.standard_b64encode(request.data).decode("ascii"),
                            },
                        },
                        {
                            "type": "text",
                            "text": f"<slide>\n{request.slide_markdown}\n</slide>\nDescribe "
                            f"the picture on this slide.{language_line}",
                        },
                    ],
                }
            ],
            output_format=_ImageDescription,
            max_tokens=4000,
        )
        result: _ImageDescription = response.parsed_output
        return ImageDescription(description=result.description, decorative=result.decorative)

    return describe_images_with(
        presentation,
        describe,
        block_ids=block_ids,
        media_types=_VISION_MEDIA_TYPES,
        max_bytes=_MAX_IMAGE_BYTES,
    )


def choose_reading_orders(
    presentation: "ParsedPresentation",
    *,
    source: str | Path | None = None,
    client: Any = None,
    model: str = DEFAULT_MODEL,
    effort: Effort = "low",
) -> Recipe:
    """Pick the reading order of each flagged slide by looking at it.

    Only slides the converter flagged are sent: the slide rendered as a one-page
    PDF (LibreOffice, locally) with its candidate orders. Without LibreOffice or a
    source file the model sees the candidates as text only. See
    ``pptx2markdown.ordering.choose_reading_orders_with`` for other judges.
    """

    client = _client(client)

    def judge(request: OrderRequest) -> str | None:
        listing = "\n\n".join(
            f"<candidate name={json.dumps(label)}>\n"
            + "\n".join(f"{index}. {words}" for index, words in enumerate(blocks, 1))
            + "\n</candidate>"
            for label, blocks in request.candidates.items()
        )
        content: list[dict[str, Any]] = []
        if request.slide_pdf is not None:
            content.append(
                {
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": base64.standard_b64encode(request.slide_pdf).decode("ascii"),
                    },
                }
            )
        content.append(
            {
                "type": "text",
                "text": f"Slide {request.page} was flagged ({', '.join(request.reasons)}).\n"
                f"{listing}\n\nWhich candidate reads the slide the way its author meant? "
                "Answer with the candidate's name exactly.",
            }
        )
        response = _parse(
            client,
            model=model,
            effort=effort,
            system=(
                "You decide the reading order of a presentation slide for a document "
                "converter. Each candidate lists the slide's text blocks in one order. Pick "
                "the order a person reading the slide would follow: a card, column or "
                "diagram step is read through before the next, titles before what they "
                f"introduce. If the candidates read equally well, choose {AS_CONVERTED!r}."
            ),
            messages=[{"role": "user", "content": content}],
            output_format=_OrderChoice,
            max_tokens=2000,
        )
        choice = response.parsed_output.choice.strip()
        return choice if choice in request.candidates else None

    return choose_reading_orders_with(presentation, judge, source=source)
