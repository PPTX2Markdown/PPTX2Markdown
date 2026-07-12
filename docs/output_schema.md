# PresentationDocument 1.0

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

## Content blocks

The order of `blocks` is the reading order. `kind` is one of `text`, `heading`,
`list`, `math`, `image`, `chart`, `smartart`, `table`, or `unsupported`.

- `content`: non-empty Markdown-compatible content.
- `shape_id`: source shape identifier when one exists.
- `heading_level`: required for `heading`, forbidden for every other kind, and
  restricted to 1 through 6.
- `bbox`: optional shape geometry with `x`, `y`, `width`, `height`, and the fixed
  unit `"emu"`. Width and height are positive; x and y may be negative.
- `source_part`: `slide`, `layout`, or `master`.

Unknown fields are rejected at every level. A future incompatible change must
use a new `schema_version`; fields in 1.0 are not silently reinterpreted.

## Determinism

The document excludes timestamps, work directories, cache paths, and absolute
source paths. For identical input and parser options, JSON output is expected to
be byte-for-byte stable.
