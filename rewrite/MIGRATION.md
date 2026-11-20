# Moving from the released converter to the rewrite

[한국어](MIGRATION.ko.md)

The rewrite is published under the same package name, `pptx2markdown`, so
upgrading replaces the released converter. This page lists what changes for a
user of 0.1.x.

## Command line

| 0.1.x | Rewrite |
| --- | --- |
| `--output-dir`, `-v` | same |
| `--output-format markdown\|json` | same, plus `chunks` (JSON Lines for retrieval) |
| `--headings auto\|strict` | same; `strict` is the recipe setting `headings: "placeholder"` |
| `--merged-cells fill\|blank` | same, plus `header` (repeat only in the header row) |
| `--work-dir` | not needed: nothing is written outside the bundle |
| `--placeholder-inheritance` | always on (`style`): placeholders inherit position, size and bullets |
| `--inherited-shapes` | always `visible`: text and pictures drawn on the layout or master are kept, inside groups too; a template picture covering the slide is background art and left out (recipe `template_pictures: false` drops all template pictures) |
| `.ppt` input, `--ppt-converter` | not supported: save the deck as `.pptx` first |
| `--convert-vector-images` | not supported: EMF/WMF images are kept as they are |
| — | `--no-assets`, `--overwrite`, `--tolerant`, `--recipe`, `--review`, `--audit`, `--ask`, `--max-chunk-chars`, `--recipe-schema` |

Inputs are named explicitly; running without inputs no longer converts every
presentation in the current directory.

## Python API

| 0.1.x | Rewrite |
| --- | --- |
| `convert(inputs, output_dir=...) -> int` (exit code) | `convert(path, output_dir=...) -> SaveResult`; `convert_many(paths) -> BatchConversionResult` with `.exit_code` |
| — | `parse(path \| bytes \| stream) -> ParsedPresentation`: the document in memory, nothing written |
| — | `slide_chunks()`, `Recipe`, `review()`, `pptx2markdown.agent`, `pptx2markdown-mcp` |

`convert()` no longer configures logging.

## Output

The bundle layout is the same: `<output>/<deck>/<deck>.md` (or `.json`), with
`media/` and `attachments/`. Each bundle has its own `bundle-manifest.json` with
asset hashes instead of one `convert_manifest.json` for the whole run.

The JSON is `PresentationDocument` 2.0. Blocks are typed instead of carrying a
Markdown `content` string:

| Kind | 1.0 field | 2.0 fields |
| --- | --- | --- |
| text, heading | `content` (+ `heading_level`) | `text` (+ `level`, `source`) |
| list | `content` (Markdown) | `items[]` with `text` and `level`, `ordered`, `start` |
| table | `content` (Markdown) | `markdown` and the cell grid `rows` (spans, merges), `header_rows`, `caption_rows` |
| math | `content` | `latex` |
| image | `content` (Markdown link) | `asset_id`, `alt_text`, `description` |
| attachment | `content` (Markdown link) | `asset_id`, `label` (file name), `program` |
| chart, smartart | `content` | `markdown` (+ `asset_ids`) |

Every block also has an `id` (`s3.b2`). The document lists its `assets` with
hashes and media types, records the source's SHA-256 and carries structured
`warnings`.

## Behaviour

Differences you will see in the Markdown, all deliberate:

- Source characters that Markdown would interpret are escaped (`\# x`,
  `\<tag>`), and bold or italic that emphasises part of a text is kept.
- A list of several numbered items is a list, not one long heading.
- A table's first row merged across the whole table is a title above it.
- Text, lists and equations in the same shape are separate blocks.
- A paragraph indented with `lvl` but without a bullet is not a list item.
- OLE attachments are unpacked to the file inside them and named after it.
