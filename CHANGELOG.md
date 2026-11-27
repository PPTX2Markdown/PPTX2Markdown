# Changelog

[한국어](CHANGELOG.ko.md)

All notable changes to the `pptx2markdown` package on PyPI.

## 0.2.0 — 2026-11-27

A new engine replaces the 0.1 converter. It was developed next to it and checked
against it on every change (a parity harness over the fixtures and a 966-deck
corpus), then went further. See the
[migration guide](docs/MIGRATION.md) for the changed CLI flags, API and output.

### Added

- **Python API**: `parse()` from a path, bytes or a stream returns the document
  in memory; `convert()` / `convert_many(jobs=N)` write portable bundles through
  a staging directory; assets are owned by the parsed presentation.
- **Reading order that follows the slide**: groups and enclosing frames are read
  as units, backgrounds are ignored, grids read row by row; slides whose layout
  can be read more than one way are flagged with candidate orders.
- **Decision points**: `review()` lists uncertain decisions (reading order,
  table header rows and captions, headings, subheadings, image descriptions)
  with stable ids, reason codes, what to look at, impact, options and a
  suggestion (review schema version 1); `recipe_from_answers()` turns answers
  into a recipe. See [the converter–agent contract](docs/agent_contract.md).
- **Recipes**: block ids (`s3.b2`) and JSON rules with selectors (`slide[12-]`,
  `section[Appendix]/image`, `table[0]`) that exclude content, set table header
  and caption rows, heading levels, image descriptions, notes and reading order;
  `--recipe`, `--recipe-schema`.
- **Audit**: `audit()` / `--audit` re-reads the slide XML and reports typed text
  missing from the output.
- **Retrieval chunks**: `slide_chunks()` and `--output-format chunks` (JSON Lines
  with citation metadata, `--max-chunk-chars`).
- **MCP server** (`pptx2markdown[mcp]`, `pptx2markdown-mcp`): convert, review,
  answer, audit, chunk, check recipes, and show pictures and flagged slides to
  the host model.
- Model-neutral judge loops: `describe_images_with()`,
  `choose_reading_orders_with()`, and `snapshot.slide_pdf()` / `slide_png()`.
- Tables keep the full cell grid with merge spans in JSON; a body row merged
  across the whole table is written once; grouped headers can span two rows.
- Inline bold and italic, escaped Markdown, PowerPoint sections, ISO Strict
  OOXML, chartEx charts from their workbook, OLE attachments unpacked with their
  ProgID, vertical text joined, layout and master pictures.
- `--jobs` converts decks in worker processes.

### Changed

- The CLI flags, Python API and JSON schema (`schema_version` 2.0) are new; see
  the migration guide.
- `.ppt` input and EMF/WMF-to-PNG run only when asked for
  (`--ppt-converter libreoffice`, `--vector-images png`), through LibreOffice.
- No work directory: nothing is written outside the output bundle.

### Removed

- The 0.1 converter and its options `--work-dir`, `--placeholder-inheritance`,
  `--inherited-shapes` and `--convert-vector-images`, and PowerPoint COM
  conversion of `.ppt`. The 0.1 code stays on the `legacy-0.1` branch.

## 0.1 line (unreleased fixes, `legacy-0.1` branch)

These fixes to the 0.1 converter were made after 0.1.2 and are not published;
the 0.2.0 engine has all of them.

### Changed

- Merged table cells repeat the merged value in every cell they cover, in the
  header and the body, so each row reads on its own (`--merged-cells fill`,
  the default). `--merged-cells blank` (API: `merged_cells="blank"`) keeps the
  value only in the cell the merge starts from. Previously only header merges
  repeated.
- Slides follow the presentation's slide list (`p:sldIdLst`) and are numbered
  by position. A deck reordered in PowerPoint used to come out in file-name
  order, and slide parts not in the presentation were converted too.

### Fixed

- A chart declaring an absurd point count (`ptCount="999999999999"`) made the
  chart converter exhaust memory; such a chart is now reported as failed.
- Embedded files named in a non-Latin Windows code page (`보고서.hwp`,
  `新建文本文档.txt`) got garbled names; the Unicode name Packager stores is used.
- A cell inside a 2D merged block (both `hMerge` and `vMerge`) pointed at its
  left neighbour and stayed empty.
- A table without rows failed its whole slide; it is now skipped with a warning
  and counted in `table_skipped_blocks`.
- Raw HTML in slide text (for example `<script>`) was written into the Markdown
  as live HTML; it is now escaped.
- An input named `...pptx` could delete the whole `--work-dir`; same-named
  decks overwrote each other; speaker-notes relationships could read files
  outside the package; zip entries differing only in case or Unicode
  normalization, and attachment names differing only in case, collided; stale
  output from an earlier run was left in place.
- A named Office lock file (`~$deck.pptx`) is skipped with a warning instead of
  silently.

### Performance

- The source package is opened once per deck and layout, master and
  relationship parts are parsed once: a 1000-slide deck converts in about 5.5 s
  instead of 12.3 s.

## 0.1.2 — 2026-09-16

- Credit the project authors and maintainers; NOTICE copyright update.

## 0.1.1 — 2026-07-29

- Render images correctly and fix the license metadata.
- Preserve original vector (EMF/WMF) image assets; optional conversion to PNG
  with `--convert-vector-images`, batched per deck.

## 0.1.0 — 2026-07-13

- First release on PyPI: static OOXML parsing, XY-cut reading order, Markdown
  and JSON from one `PresentationDocument`, golden regression suite.
