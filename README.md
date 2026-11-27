# pptx2markdown

[![Quality](https://github.com/PPTX2Markdown/PPTX2Markdown/actions/workflows/quality.yml/badge.svg)](https://github.com/PPTX2Markdown/PPTX2Markdown/actions/workflows/quality.yml)
[![PyPI](https://img.shields.io/pypi/v/pptx2markdown)](https://pypi.org/project/pptx2markdown/)
[![Python](https://img.shields.io/pypi/pyversions/pptx2markdown)](https://pypi.org/project/pptx2markdown/)
[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/LICENSE)

English | [한국어](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/README.ko.md)

`pptx2markdown` converts PowerPoint presentations into deterministic Markdown,
JSON or retrieval chunks. It reads the native OOXML structure and shape geometry
directly, never calls a model, and says where it is unsure: the decisions worth a
second look come with ready-made alternatives, so a person, a UI or a model can
settle them with a recipe that replays deterministically.

## Why pptx2markdown?

- **Reading order that follows the slide** — cards, labelled boxes and grouped
  shapes are read as units and backgrounds are ignored; slides whose layout can be
  read more than one way are flagged with the candidate orders.
- **Tables that keep their meaning** — the full cell grid with merge spans, merged
  values repeated so every row stands alone, grouped two-row headers on request.
- **Decision points instead of silent guesses** — `review()` lists uncertain
  choices with stable ids, reasons, what to look at and how much it matters;
  answers become a recipe.
- **Nothing lost quietly** — `audit()` re-reads the slide XML and reports typed
  text missing from the output.
- **Retrieval-ready** — slide-aligned chunks with citation metadata; parsing from
  bytes or a stream; an MCP server for agents.
- **Rich native content** — text, headings, lists, tables, charts, SmartArt,
  equations, images, speaker notes, sections and embedded attachments.

What each converter keeps of a deck, from the
[feature benchmark](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/benchmarks/README.md):

| | pptx2markdown | anydoc 0.2.5 | MarkItDown 0.1.9 |
| --- | :---: | :---: | :---: |
| Features kept (of 20) | **20** | 13 | 8 |

Upgrading from 0.1.x? The output, CLI flags and Python API changed; see the
[migration guide](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/MIGRATION.md).

## Installation

### Let an AI agent install it

A local coding agent can inspect the environment, choose an isolated install,
optionally configure LibreOffice, and verify the CLI. Give it this prompt and
review every permission request:

```text
Read https://raw.githubusercontent.com/PPTX2Markdown/PPTX2Markdown/main/.github/agent-install.md and install pptx2markdown for this machine. You may run the commands needed for installation after explaining them. Ask me before any administrator, sudo, password, system package-manager, shell-profile, PATH, or LibreOffice change.
```

See the complete [agent installation procedure](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/.github/agent-install.md).

### Manual installation

```bash
uv tool install pptx2markdown
# or
pipx install pptx2markdown
# or, inside an active virtual environment
python -m pip install pptx2markdown
```

Extras: `pptx2markdown[mcp]` (the MCP server) and `pptx2markdown[images]`
(trims EMF/WMF pictures drawn as PNG).

LibreOffice is optional. It is used only when asked for: `.ppt` input, EMF/WMF
pictures as PNG, and rendering the slides whose reading order is in question.

```bash
# macOS
brew install --cask libreoffice

# Ubuntu / Debian
sudo apt-get install libreoffice-impress
```

## Quick start

```bash
# One file -> ./output/deck/deck.md
pptx2markdown deck.pptx

# Several files, a custom output root, one worker per CPU
pptx2markdown decks/*.pptx --output-dir converted --jobs 0

# Canonical JSON, or one JSON line per slide chunk
pptx2markdown deck.pptx --output-format json
pptx2markdown deck.pptx --output-format chunks

# What the converter was unsure about, and text it could not keep
pptx2markdown deck.pptx --review
pptx2markdown deck.pptx --audit
```

## Requirements

- Python 3.12 or later
- Pydantic 2.10 or later
- omml2latex 0.1.0 or later
- chart2md 0.1.0 or later
- smartart2md 0.1.1 or later
- olefile 0.47 or later
- local `.pptx` input; legacy `.ppt` only when asked for, through LibreOffice

There is no OCR, VLM, API key, cloud inference or PowerPoint automation in the
parsing path; models only ever see what you hand them. A `.ppt` file is converted
only with `ppt_converter="libreoffice"` (`--ppt-converter libreoffice`), which runs
LibreOffice headless in temporary directories that are removed afterwards
(`SOFFICE_PATH` selects the program). The same goes for `vector_images="png"`
(`--vector-images png`): EMF/WMF pictures, which browsers, Markdown viewers and
vision models cannot read, are drawn as PNG in one LibreOffice run and take the
original's place under the same asset id; with the `images` extra
(`pip install "pptx2markdown[images]"`) the page margin LibreOffice adds is trimmed.

## Python API

```python
import pptx2markdown

with pptx2markdown.parse("deck.pptx") as presentation:
    for slide in presentation.slides:
        print(slide.page, slide.blocks)

    result = presentation.save("output/deck")
    print(result.document_path)
```

`parse()` also takes the file's bytes or a binary stream, so an upload can be
converted without touching the disk (`name` is recorded as the source name):

```python
with pptx2markdown.parse(request.files["deck"].stream, name="deck.pptx") as presentation:
    markdown = presentation.render_markdown(assets="omit")
```

In-memory input is limited to `PackageLimits.max_package_bytes` (1 GiB).

`parse()` does not leave persistent output. Embedded payloads are kept in a
managed spooled store and no temporary path appears in the canonical model.

Markdown rendered without saving must receive an explicit asset policy:

```python
with pptx2markdown.parse("deck.pptx") as presentation:
    text_only = presentation.render_markdown(assets="omit")

    web_markdown = presentation.render_markdown(
        asset_url=lambda asset_id: f"/assets/{asset_id}",
    )
```

### Chunks for retrieval

`slide_chunks()` returns one chunk per slide with citation metadata (source,
page, slide and deck title, block ids, block kinds). A slide longer than
`max_chars` is split at block boundaries, never inside a table. Each chunk starts
with a comment naming the deck and slide, so a retrieved chunk keeps its context.

```python
with pptx2markdown.parse("deck.pptx") as presentation:
    for chunk in pptx2markdown.slide_chunks(presentation.document, max_chars=2000):
        index.add(chunk.text, metadata=chunk.metadata)
```

`--output-format chunks` writes the same chunks as JSON lines
(`<deck>.chunks.jsonl`, `{"text": ..., "metadata": ...}` per line) with image
links into the bundle; `--max-chunk-chars` sets the split size.

### Recipes

A recipe adjusts a parsed document without touching the parser: document-wide
settings plus rules that select blocks or slides and say what to do with them.
It is plain data, so it can be stored next to a deck, reviewed, shared across a
team, or written by a language model from a request such as "drop the appendix
and treat the first two rows of the sales table as its header". Applying it is
deterministic.

```python
recipe = pptx2markdown.Recipe(
    merged_cells="blank",
    rules=[
        {"target": "slide[12-]", "exclude": True, "reason": "appendix"},
        {"target": "slide[3]/table[0]", "header_rows": 2},
        {"target": "s5.b2", "description": "Revenue by quarter, 2021-2024"},
        {"target": "s1.b3", "heading_level": 2},
    ],
)
with pptx2markdown.parse("deck.pptx", recipe=recipe) as presentation:
    ...
```

Every block has an id `s<page>.b<n>`; `render_markdown(block_ids=True)` shows
them as `<!-- s3.b2 table -->` comments so a reviewer or a model can point at
exact blocks, and `explain_recipe()` lists what each rule selects before it is
applied. Selectors:

| Selector | Selects |
| --- | --- |
| `s3.b2` | one block by id |
| `slide[3]`, `slide[2-5]`, `slide[12-]`, `slide[*]`, `slide[hidden]` | whole slides |
| `table`, `image`, `heading`, ... or `*` | every block of a kind |
| `table[0]`, `table[-1]` | the first or last block of a kind on each slide |
| `slide[3]/table[0]` | the first table on slide 3 |
| `section[Appendix]`, `section[Appendix]/image` | the slides (or blocks) of a PowerPoint section |

Settings: `merged_cells`, `include_notes`, `include_hidden`, and `headings`
(`placeholder` keeps only headings from title placeholders, like the legacy
converter's `--headings strict`).

Rule actions: `exclude`, `merged_cells`, `header_rows`, `caption_rows` (tables),
`heading_level` (set a heading's level or turn text into a heading), `as_text`
(headings and lists), `description` (image alt text), `notes` and `order`
(slide selectors: keep or drop notes; read the blocks in a given order of block
ids), and `reason`, which is kept for review only. Later rules win, and
every selector sees the document as parsed. `recipe_json_schema()` (or
`pptx2markdown --recipe-schema`) gives the JSON Schema for forms and structured
model output.

### Review

The converter always decides; `review()` points at the decisions most worth a
second look and offers each alternative as a ready-made recipe rule: a header
with merged groups whose next row looks like labels (two header rows?), a first
row that is all numbers (no header?), a merged first row taken as a table title,
an image with no description, a heading decided by font size alone (a
heading?), a short bold line that introduces a list or table (a
subheading?), or a slide whose layout can be read in more than one order. A
person, a UI or a language model only picks options, and the picks form a
recipe.

Reading order follows the slide's structure: groups and shapes that enclose
other shapes (a card's frame, a labelled box, a picture with callouts) are read
as units, and backgrounds covering the slide are ignored. Where the geometry
cannot settle the order (a grid that reads as well by rows as by columns, side
by side text with no gap between), the slide gets a `reading-order` review item
whose options are the candidate orders as `order` rules, so only those slides
need a person or a model to look at them. `choose_reading_orders_with(presentation,
judge)` renders just those slides as one-page PDFs (LibreOffice, locally) and asks
your `judge` function to name a candidate. See [the design](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/reading_order_design.md).

Every item has the same shape whatever it is about, so a caller can plan
before it looks: a stable `id` (`table-header-rows@s3.b2`), machine-readable
`reasons` (`grouped-header`, `two-way-cut`, ...), the `page` and what to `look` at
to decide (`text`: the slide's Markdown is enough; `slide`: its rendering;
`image`: the picture, by `asset_id`), an `impact` (`high` when a wrong answer
corrupts content, `low` for formatting), options with stable ids (the first is
always `keep`) and the `suggested` one. `recipe_from_answers` turns
`{item id: option id}` into a recipe; an image description is answered with free
text. See [the contract](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/agent_contract.md).

```python
with pptx2markdown.parse("deck.pptx") as presentation:
    items = presentation.review()
    for item in items:
        print(item.id, item.look, item.question, [option.id for option in item.options])
    recipe = pptx2markdown.recipe_from_answers(
        items, {"table-header-rows@s3.b2": "two-rows"}, unanswered="suggested"
    )
```

`pptx2markdown deck.pptx --review` prints the same as JSON (with
`schema_version`) without writing bundles.

### Audit

A rule-based converter fails quietly: a shape it does not understand is simply
missing from the output. `audit()` reads the slide XML again, without the
parser's rules, and reports every piece of typed text that the document does not
contain, with its slide and shape id, plus slides missing from the document.
Text is compared as letters and digits, so escaping, emphasis, list markers and
line breaks never count as loss; footer, date and slide-number placeholders and
cells hidden by a merge are left out. Charts, SmartArt and equations are not
audited.

```python
with pptx2markdown.parse("deck.pptx") as presentation:
    report = presentation.audit()
    print(report.ok, report.coverage, report.losses)
```

`pptx2markdown deck.pptx --audit` prints the report as JSON without writing
bundles and exits with 1 when text was lost. Content removed by a recipe would
be reported too, so the CLI and MCP tool audit the conversion without one.

### Judging with a model

Nothing in the package calls a model. Where judgement helps, it hands a function
you supply exactly what that judgement needs and turns the answers into recipe
rules: `answer` review items with `recipe_from_answers`, describe pictures with
`describe_images_with`, pick reading orders with `choose_reading_orders_with`.
The function can be a person, a local model or a hosted one. Through the MCP
server below, the host model (Claude Code, for example) does the same with no API
key of its own. A multi-format agent can sit on top through
[the contract](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/agent_contract.md).

Images can be described by any model, local or not, with nothing extra to install:
`describe_images_with` hands each picture that needs a description to a function
you supply, with the slide's Markdown for context, and returns a recipe. Nothing
leaves the machine unless that function sends it.

```python
from pptx2markdown import ImageDescription, describe_images_with

def describe(request):  # request.data, .media_type, .alt_text, .slide_markdown
    text = my_local_vision_model(request.data, context=request.slide_markdown)
    return ImageDescription(description=text)

with pptx2markdown.parse("deck.pptx") as presentation:
    presentation.apply(describe_images_with(presentation, describe))
```

### MCP server

`pip install "pptx2markdown[mcp]"` adds `pptx2markdown-mcp`, a Model Context
Protocol server (stdio) so an agent such as Claude can run the recipe workflow
itself: `convert_presentation` (Markdown with block ids, optionally with a
recipe applied, or only some `pages`), `review_presentation` (filterable by
topic, look, impact and page), `answers_to_recipe` (review answers
to a recipe), `audit_presentation` (text the
conversion lost), `presentation_images` (pictures without a description, as
images, so the host model can describe them in a recipe), `presentation_slides`
(the slides whose reading order is in question, rendered locally by LibreOffice,
so the host model can pick the order), `check_recipe` (what
each rule selects), `presentation_chunks` and `recipe_schema`. Two prompts
walk a host model through the workflow: `prepare_for_retrieval` (audit, answer the
review, describe pictures, check the recipe, chunk) and `apply_request` (turn a
plain-language request into a checked recipe). It reads local `.pptx` files and
writes nothing.

```bash
claude mcp add pptx2markdown -- pptx2markdown-mcp
```

or, in an MCP client configuration:

```json
{"mcpServers": {"pptx2markdown": {"command": "pptx2markdown-mcp"}}}
```

Use `convert()` for parse-and-save:

```python
result = pptx2markdown.convert(
    "deck.pptx",
    output_dir="output",
)
```

## CLI

```bash
pptx2markdown deck.pptx
pptx2markdown first.pptx second.pptx --output-dir converted
pptx2markdown decks/*.pptx --jobs 0   # one worker process per CPU
pptx2markdown deck.pptx --output-format json
pptx2markdown deck.pptx --output-format chunks --max-chunk-chars 2000
pptx2markdown deck.pptx --merged-cells blank
pptx2markdown deck.pptx --recipe recipe.json
pptx2markdown --recipe-schema
pptx2markdown deck.pptx --review
pptx2markdown deck.pptx --audit
pptx2markdown old-deck.ppt --ppt-converter libreoffice --vector-images png
```

PowerPoint's lock files (`~$deck.pptx`, written while a deck is open) are skipped
with a warning, so a glob such as `decks/*.pptx` does not fail on them.

Default bundle:

```text
output/
└── deck/
    ├── deck.md
    ├── media/
    ├── attachments/
    └── bundle-manifest.json
```

Markdown links are relative to `deck.md`, so the whole `deck/` directory can be
moved without breaking embedded assets.

## Supported static content

- slide order and hidden-slide metadata
- geometry-only reading order by recursive XY-cut: columns separated by a vertical
  gap are read one after another, a full-width title or divider is read before
  the columns below it, a stack of slightly offset boxes stays top to bottom, an
  overlap of up to one pixel still counts as a gap, and background frames and
  footers cannot seal a column gap
- headings from title/subtitle placeholders, section numbering and font size;
  each heading records its `source` (`placeholder`, `position`, `numbering`,
  `text` or `font`) so the weakest decisions can be reviewed or dropped
- placeholder type, position, font size and bullets inherited from the slide
  layout and master; footer placeholders are dropped
- text shapes and pictures drawn from the slide layout and master (blocks keep
  `source_part`; thin lines and bands are decoration; the recipe setting
  `template_pictures: false` leaves repeated logos out)
- text, hyperlinks and lists; consecutive plain paragraphs of a shape form one
  text block (one line each) and consecutive bullets one list, while text, lists
  and equations in the same shape stay separate typed blocks; a typed `▶` starts
  a list item, and a paragraph indented with typed spaces right after an item
  continues it
- bold and italic runs as `**bold**`/`*italic*` when they emphasise part of the
  text; bold over a whole table cell or a whole multi-line text box is styling
  and is left out, and emphasis a renderer could not show is dropped; text, heading and list `text`
  is escaped inline Markdown, so typed `# x`, `- x`, `<tag>` or `[x](y)` render
  literally instead of becoming a heading, list, HTML or link
- equations (OMML) as LaTeX via omml2latex
- embedded images, including shapes filled with a picture
- embedded audio/video payloads
- OLE, 3D model, audio and video attachments, labelled by file name; an OLE
  compound file is unpacked to the file inside it (the Packager's original file
  or the server's document stream) and its ProgID is kept as `program`
- native tables: JSON keeps the full grid (`rows`, with each merge's text,
  `row_span`/`col_span`, and `merged_from` on the cells it covers) next to the
  rendered `markdown`; the first row is the header and an empty header becomes
  `col_N`; `merged_cells` decides what a covered cell shows (`fill`, the
  default, repeats the merged value so every row stands alone for retrieval;
  `header` repeats it only in the header row; `blank`
  never); a first row merged across the whole table is a title written once
  above it (`caption_rows`); bullets, numbering and links inside cells
  are kept, and pictures lying on a cell or filling its background are written
  into that cell
- charts via chart2md: classic charts from their cached values, Office 2016
  chartEx charts (funnel, waterfall, ...) from their embedded workbook, which is
  checked like the package itself and falls back to cached values when rejected
- SmartArt via smartart2md: hierarchies as indented lists, process and cycle
  layouts as numbered steps, matrices as tables and pictures as images; a
  diagram whose nodes have no text is counted as `[smartart: N unlabeled nodes]`
- speaker notes
- grouped shapes and Markup Compatibility `AlternateContent`

Unsupported or unavailable content produces a structured warning or an
`UnsupportedBlock`; it is not silently treated as successfully converted.

OCR and tables drawn from the layout or master are not supported;
image descriptions come from a describer you supply.

## Development

```bash
python -m pip install -e ".[mcp]"
python -m unittest discover -v
python scripts/update_goldens.py --check
ruff check src tests scripts benchmarks
ruff format --check src tests scripts benchmarks
```

When conversion output changes on purpose, review the diff of the golden
snapshots before accepting them with `python scripts/update_goldens.py --accept`.

Coming from 0.1.x? See [the migration guide](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/MIGRATION.md).

See [ARCHITECTURE.md](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/ARCHITECTURE.md) for ownership, lifecycle and security
details.

## License

Copyright 2026 Moonyoung Lee, Inseong Lee, Seung-Woo Jeong

The project is distributed under the [Apache License 2.0](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/LICENSE).
Third-party golden fixture attribution is recorded in
[`THIRD_PARTY_NOTICES.md`](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/tests/fixtures/golden/THIRD_PARTY_NOTICES.md).
