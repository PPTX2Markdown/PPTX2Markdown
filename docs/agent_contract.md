# The converter–agent contract

[한국어](agent_contract.ko.md)

Status: in effect for the next-generation engine in `rewrite/` (review schema
version 1).

## Roles

pptx2markdown is a deterministic PPTX converter that says where it is unsure. It
never calls a model while parsing. Anything that judges (a person, a UI, a local
model, Claude, or a multi-format agent built elsewhere) sits on top and talks to
it through three plain-data interfaces:

| Interface | Direction | What it is |
| --- | --- | --- |
| Document | converter → agent | Markdown with block ids (`<!-- s3.b2 table -->`) and canonical JSON |
| Decision points (`review()`) | converter → agent | Where the converter chose without being sure, with ready-made alternatives |
| Recipe | agent → converter | The answers as rules; replays deterministically, never edits the source |

`audit()` is a fourth, read-only signal: text typed on the slides that the
document lacks. A recipe cannot bring it back, so an agent reports it instead of
answering it.

The split keeps each side replaceable. The converter can grow new decision topics
without any agent changing code; an agent can serve several formats if every
converter reports decisions in this shape.

## A decision point

```json
{
  "id": "table-header-rows@s3.b2",
  "topic": "table-header-rows",
  "target": "s3.b2",
  "page": 3,
  "question": "The header has merged groups; is the next row part of the header?",
  "reasons": ["grouped-header"],
  "look": "text",
  "impact": "high",
  "options": [
    {"id": "keep", "label": "One header row", "rule": null},
    {"id": "two-rows", "label": "Two header rows",
     "rule": {"target": "s3.b2", "header_rows": 2, "reason": "grouped two-row header"}}
  ],
  "suggested": "two-rows",
  "free_text": null,
  "asset_id": null,
  "evidence": {"header": ["Item", "2024", "2024"], "next_row": ["Item", "H1", "H2"], "header_spans_rows": true}
}
```

| Field | Meaning |
| --- | --- |
| `id` | `<topic>@<target>`; stable for one deck and converter version |
| `topic` | What kind of decision (table below) |
| `target` | Recipe selector the options' rules apply to |
| `page` | The slide to look at |
| `reasons` | Machine-readable codes for why it is asked |
| `look` | What a judge needs: `text` (the slide's Markdown is enough), `slide` (its rendering), `image` (the picture, `asset_id`) |
| `impact` | `high`: a wrong answer corrupts content for retrieval; `low`: formatting |
| `options` | Answers with stable ids; the first is always `keep` (rule `null`, the converter's choice) |
| `suggested` | The option to take when nothing settles it |
| `free_text` | When set, an answer may be free text that fills this rule field |
| `evidence` | Topic-specific facts the converter used |

`look` and `impact` let an agent budget: answer `text` items from the Markdown
in one batch, render only the pages of `slide` items, fetch only the pictures of
`image` items, and skip `low` items when time or money is short.

### Topics

| Topic | Target | Reasons | Look | Impact | Options |
| --- | --- | --- | --- | --- | --- |
| `reading-order` | `slide[N]` | `two-way-cut`, `no-cut`, `overlapping-boxes`, `fragile-gap`, `container-across-content` | slide | high | `keep`, `rows-first`, `columns-allowing-overlap`, `as-stored-in-the-file` (those that differ) |
| `table-header-rows` | block | `grouped-header`, `numeric-first-row` | text | high | `keep`, `two-rows` / `no-header` |
| `image-description` | block | `no-alt-text`, `default-picture-name` | image | high | `keep`, `decorative`; free text → `description` |
| `table-caption` | block | `merged-first-row` | text | low | `keep`, `header` |
| `heading` | block | `font-size-only` | text | low | `keep`, `text` |
| `subheading` | block | `bold-line-introduces-block` | text | low | `keep`, `heading` |

Reading-order evidence lists, per option id, the first words of each block in
that order.

## Answering

`recipe_from_answers(items, {item id: option id or free text})` builds the
recipe; the MCP tool `answers_to_recipe` does the same. Items left out keep the
converter's decision, or take `suggested` with `unanswered="suggested"`. Unknown
item ids and options raise one error listing every problem.

Model-free judge loops build the requests and collect the answers, for any
callable:

- `describe_images_with(presentation, describe)`: picture bytes, alt text and
  the slide's Markdown in; a description or "decorative" out.
- `choose_reading_orders_with(presentation, judge)`: each flagged slide as a
  one-page PDF (LibreOffice), its reasons and candidates in; a candidate name out.

`pptx2markdown.snapshot.slide_pdf` / `slide_png` render one slide for any other
judge, and the MCP tool `presentation_slides` shows the slides of `look: "slide"`
items to the host model, so it can answer them without an API key of its own.

`pptx2markdown.agent` (the optional `[agent]` extra) is these loops with Claude.
It is the only module that imports a model SDK; the core and the MCP server
never load it (a test enforces this).

## Versioning

`schema_version` in `--review` output and in the MCP tool's result is bumped when
a field changes meaning or goes away. New topics, reasons, options and fields
keep the version, so an agent must ignore what it does not know and may always
fall back to `suggested`. Recipes are versioned by their own JSON Schema
(`--recipe-schema`).
