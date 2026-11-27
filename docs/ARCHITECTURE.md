# Architecture

[한국어](ARCHITECTURE.ko.md)

## Dependency direction

```text
api / cli / mcp_server
    ↓
parser → shapes (geometry, footers, decoration) / paragraphs (runs, lists)
       → ooxml / inheritance / reading_order / headings / inline
       → omml / charts / smartart / tables / attachments
    ↓
model + AssetStore
    ↓
ParsedPresentation ←── recipe (apply_recipe: document → document)
    ↓                  review (document → questions with recipe rules)
renderer / chunks
    ↓
transactional bundle writer
```

The canonical document never owns a local path or binary payload. Runtime
ownership lives in `ParsedPresentation`, which closes its `AssetStore`.

## Decisions as data

The parser always decides: reading order, headings, list structure, table
headers and titles. Everything after parsing works on the canonical document
only:

- `recipe.apply_recipe` is a pure function from a document and a recipe to a
  document. Selectors address blocks by their id (`s<page>.b<n>`, assigned in
  reading order) or by slide and kind; selectors always see the document as
  parsed, so a rule's targets never depend on earlier rules.
- `review.review` reads the document and returns the decisions most worth a
  second look, each with alternatives that are ready-made recipe rules.
- `audit.audit` re-reads the slide XML without the parser's rules and reports
  typed text the document does not contain: a check on the parser that shares
  none of its code paths for text, so it catches what the parser silently drops.
- `images.describe_images_with` is the model-neutral image loop: pictures and
  their slide go to any describer function, rules come back. A hosted model, a local
  model or an OCR engine plugs in the same way; the package imports no model SDK.

Because tables keep their full grid (`rows`, merge spans, `merged_from`) and
headings keep the signal that made them headings (`source`), a recipe can
re-render or reclassify blocks without parsing the package again.

## Inline Markdown

Text, heading, list and table-cell text is escaped inline Markdown. While the
parser runs its text rules (headings, typed bullets), generated markup (links,
inline math, emphasis) is wrapped in private-use sentinels; `inline.finalize`
escapes only the source characters and removes the sentinels when blocks are
built. Speaker notes stay plain in the model and are escaped when rendered.

## Reading order

`reading_order` is a recursive XY-cut: a vertical gap through the region splits
columns read one after another; otherwise the top band is read first and the
rest is cut again. A stack of single closely spaced boxes stays top to bottom,
an overlap of up to one pixel still counts as a gap, and footers and decorative
shapes do not take part in finding gaps.

## Asset identity

Each logical OOXML part receives an ID derived from its normalized package part
name. SHA-256 of the bytes is recorded separately. This keeps logical identity
stable while permitting payload deduplication.

Small payloads remain in memory. Payloads above the spool threshold move to an
OS-managed temporary file whose path is never returned. Closing the parsed
presentation closes every payload.

## Rendering and writing

Image and attachment blocks contain an `asset_id`, not Markdown with a path.
The Markdown renderer requires either a URL resolver or explicit omission.

The bundle writer materializes assets first, verifies size and SHA-256, renders
relative links, writes a deterministic manifest, verifies containment, and only
then renames the staging directory into place. Existing destinations are not
modified unless `overwrite="replace"` is explicit.

## Determinism

The canonical document excludes timestamps, absolute input paths, work paths and
temporary names. Assets and warnings are sorted deterministically. Bundle file
collisions use content hash suffixes.

## Security boundaries

Before any part is read, the package validates entry count, encrypted entries,
symlinks, traversal, duplicate normalized names, uncompressed sizes and
compression ratios. Relationship targets are normalized relative to their
source part and cannot leave the ZIP. XML with a DTD or entity declaration is
rejected.

chart2md reads chart parts only through an adapter over the validated package,
never through its own `zipfile`. An embedded chart workbook is a ZIP of its own,
so it gets the same entry, size and ratio checks plus the DTD check on its XML
parts before chart2md opens it; a rejected workbook leaves a warning and the
chart falls back to its cached values.

An embedded OLE compound file is unpacked with olefile from an in-memory
buffer whose size is already bounded by the asset limits; it is never passed
as raw bytes, which olefile would read as a file path when short.

