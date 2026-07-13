# PresentationDocument 1.0

English | [한국어](output_schema.ko.md)

`PresentationDocument` is the canonical output contract. Markdown and JSON are
two renderings of the same ordered document; parsing does not diverge by output
format.

The machine-readable schema ships with the Python package as
[`presentation_document.schema.json`](../src/pptx2markdown/presentation_document.schema.json).

## Top level

- `schema_version`: always `"1.0"`.
- `source.name`: source basename only. Absolute or workspace-dependent paths are
  intentionally excluded.
- `source.format`: `"pptx"` or `"ppt"`.
- `slides`: slides in strictly increasing page order. Page numbers are positive
  and unique.
- `slides[].hidden`: whether the source slide is hidden from the normal
  slideshow. Hidden slides remain in the document; Markdown renders
  `<!-- hidden: true -->` immediately after their page marker.
- `slides[].notes`: optional speaker notes extracted from the notes body
  placeholder. Notes remain separate from visual `blocks`; Markdown renders
  them under a `[Speaker_Notes]` marker after the slide content. Notes-page
  headers, footers, dates, and slide numbers are excluded.

PowerPoint review comments are intentionally excluded. They are collaboration
metadata, are not displayed in the slideshow, and must not be mixed with slide
content or speaker notes.

## Content blocks

The order of `blocks` is the reading order. `kind` is one of `text`, `heading`,
`list`, `math`, `image`, `chart`, `smartart`, `table`, `attachment`, or
`unsupported`.

`attachment` preserves an embedded package or media file under the output
`attachments/` directory. Its content is a Markdown link such as
`[attachment: report.pdf](attachments/report.pdf)`. OLE Packager, Package, and
CONTENTS streams are unpacked when possible; an unreadable container is retained
as `.bin` instead of being discarded.

- `content`: non-empty Markdown-compatible content.
- `shape_id`: source shape identifier when one exists.
- `heading_level`: required for `heading`, forbidden for every other kind, and
  restricted to 1 through 6.
- `bbox`: optional shape geometry with `x`, `y`, `width`, `height`, and the fixed
  unit `"emu"`. Width and height are positive; x and y may be negative.
- `source_part`: `slide`, `layout`, or `master`.

Unknown fields are rejected at every level. A future incompatible change must
use a new `schema_version`; fields in 1.0 are not silently reinterpreted.

## Markdown rendering contract

Markdown is a deterministic rendering of the same `PresentationDocument`, not a
second parser output.

- A document with no slides renders as a zero-byte file.
- Each slide starts with `[Page_<page>]`.
- A hidden slide adds `<!-- hidden: true -->` after its page marker.
- Blocks follow in `slides[].blocks` order. Only `heading` adds Markdown syntax
  automatically (`#` through `######`); every other block already contains its
  complete Markdown-compatible representation.
- Speaker notes, when present, follow visual blocks under `[Speaker_Notes]`.
- Media and attachment links are relative to the per-document output directory.

Whitespace between these sections is part of the golden contract and is covered
by byte-for-byte JSON and Markdown snapshots.

## Validation and versioning policy

The shipped JSON Schema validates the serialized structure. Runtime Pydantic
validators additionally enforce semantic invariants that JSON Schema cannot
express concisely: slide pages are unique and strictly increasing, headings
require a level, and non-heading blocks cannot carry a heading level.

Version `1.0` is the current frozen development contract. The project has not
been released, so no compatibility adapter is maintained for earlier internal
shapes. An intentional incompatible change must update `schema_version`, this
document, the checked-in JSON Schema, and the repository golden snapshots in the
same change.

## Determinism

The document excludes timestamps, work directories, cache paths, and absolute
source paths. For identical input and parser options, JSON output is expected to
be byte-for-byte stable.
