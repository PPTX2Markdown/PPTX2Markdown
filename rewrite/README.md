# PPTX2Markdown rewrite

[한국어](README.ko.md)

This directory is a standalone rewrite. It does not import or modify the legacy
converter in the repository root.

The rewrite separates four responsibilities:

1. `parse()` reads OOXML into a path-independent `PresentationDocument`.
2. `AssetStore` owns image and attachment payloads for the lifetime of the
   returned `ParsedPresentation`.
3. renderers turn the canonical document into Markdown or JSON without writing
   files.
4. `save()` writes a verified, portable bundle through a staging directory.

## Requirements

- Python 3.12 or later
- Pydantic 2.10 or later
- omml2latex 0.1.0 or later
- chart2md 0.1.0 or later
- smartart2md 0.1.1 or later
- olefile 0.47 or later
- local `.pptx` input

There is no OCR, VLM, API key, cloud inference, PowerPoint automation or
LibreOffice dependency in the parsing path; the optional agent is a separate layer.

## Python API

```python
import pptx2markdown

with pptx2markdown.parse("deck.pptx") as presentation:
    for slide in presentation.slides:
        print(slide.page, slide.blocks)

    result = presentation.save("output/deck")
    print(result.document_path)
```

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

Settings: `merged_cells`, `include_notes`, `include_hidden`, and `headings`
(`placeholder` keeps only headings from title placeholders, like the legacy
converter's `--headings strict`).

Rule actions: `exclude`, `merged_cells`, `header_rows`, `caption_rows` (tables),
`heading_level` (set a heading's level or turn text into a heading), `as_text`
(headings and lists), `description` (image alt text), `notes` (slide
selectors), and `reason`, which is kept for review only. Later rules win, and
every selector sees the document as parsed. `recipe_json_schema()` (or
`pptx2markdown --recipe-schema`) gives the JSON Schema for forms and structured
model output.

### Review

The converter always decides; `review()` points at the decisions most worth a
second look and offers each alternative as a ready-made recipe rule: a header
with merged groups whose next row looks like labels (two header rows?), a first
row that is all numbers (no header?), a merged first row taken as a table title,
an image with no description, a heading decided by font size alone (a
heading?), or a short bold line that introduces a list or table (a
subheading?). A person, a UI or a language model only picks options, and
the picks form a recipe.

```python
with pptx2markdown.parse("deck.pptx") as presentation:
    for item in pptx2markdown.review(presentation.document):
        print(item.target, item.question, [option.label for option in item.options])
```

`pptx2markdown deck.pptx --review` prints the same as JSON without writing
bundles.

### Agent (optional)

`pip install "pptx2markdown[agent]"` adds a language-model layer on top of
recipes and review. The model never edits the parser or the text: it reads the
deck with block ids and writes a recipe, which is validated, checked so that every
rule selects something (a failing recipe is sent back once with the problem), and
applied like a hand-written one. It uses Claude (`claude-opus-5-5` by default,
with the API's refusal fallback on) through the Anthropic SDK and the usual
`ANTHROPIC_API_KEY` or `ant auth login` credentials. Nothing in the parsing path
calls a model.

```python
from pptx2markdown import agent, merge_recipes

with pptx2markdown.parse("deck.pptx") as presentation:
    draft = agent.draft_recipe(
        presentation.document,
        "This goes into a search index: drop the appendix after slide 20 and "
        "treat grouped table headers as two rows.",
    )
    print(draft.summary, draft.unmatched)
    answers = agent.answer_review(presentation.document)  # picks review options
    images = agent.describe_images(presentation, language="Korean")  # vision
    presentation.apply(merge_recipes(draft.recipe, answers, images))
```

`pptx2markdown deck.pptx --ask "drop the appendix"` prints the drafted recipe as
JSON for review; apply it with `--recipe`.

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
pptx2markdown deck.pptx --output-format json
pptx2markdown deck.pptx --output-format chunks --max-chunk-chars 2000
pptx2markdown deck.pptx --merged-cells blank
pptx2markdown deck.pptx --recipe recipe.json
pptx2markdown --recipe-schema
pptx2markdown deck.pptx --review
pptx2markdown deck.pptx --ask "drop the appendix" > recipe.json
```

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
- text shapes drawn from the slide layout and master
- text, hyperlinks and lists; consecutive plain paragraphs of a shape form one
  text block (one line each) and consecutive bullets one list, while text, lists
  and equations in the same shape stay separate typed blocks; a typed `▶` starts
  a list item, and a paragraph indented with typed spaces right after an item
  continues it
- bold and italic runs as `**bold**`/`*italic*`; text, heading and list `text`
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

Legacy `.ppt`, pictures and tables drawn from the layout or master, OCR, image
descriptions and vector image rasterization are not part of this first
standalone rewrite.

## Development

From the repository root:

```bash
PYTHONPATH=rewrite/src python -m unittest discover -s rewrite/tests -t rewrite -v
ruff check rewrite
ruff format --check rewrite
```

See [ARCHITECTURE.md](ARCHITECTURE.md) for ownership, lifecycle and security
details.

