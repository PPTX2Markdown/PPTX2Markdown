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
- local `.pptx` input

There is no OCR, VLM, API key, cloud inference, PowerPoint automation or
LibreOffice dependency in the parsing path.

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
pptx2markdown deck.pptx --merged-cells blank
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
- geometry-only reading order
- headings from title/subtitle placeholders, section numbering and font size
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
- embedded images
- embedded audio/video payloads
- OLE, 3D model, audio and video attachments, labelled by file name; an OLE
  compound file is unpacked to the file inside it (the Packager's original file
  or the server's document stream) and its ProgID is kept as `program`
- native tables: JSON keeps the full grid (`rows`, with each merge's text,
  `row_span`/`col_span`, and `merged_from` on the cells it covers) next to the
  rendered `markdown`; the first row is the header and an empty header becomes
  `col_N`; `merged_cells` decides what a covered cell shows (`fill`, the
  default, repeats the merged value so every row stands alone for retrieval;
  `header` repeats it only in the header, as the legacy converter does; `blank`
  never); bullets, numbering and links inside cells are kept, and pictures lying
  on a cell or filling its background are written into that cell
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

