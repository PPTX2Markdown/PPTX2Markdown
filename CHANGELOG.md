# Changelog

[한국어](CHANGELOG.ko.md)

All notable changes to the released converter (`src/`, the `pptx2markdown`
package on PyPI). The next-generation engine in `rewrite/` is unreleased and
documented in [rewrite/README.md](rewrite/README.md).

## Unreleased

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
