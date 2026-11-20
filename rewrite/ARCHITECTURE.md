# Rewrite architecture

[한국어](ARCHITECTURE.ko.md)

## Dependency direction

```text
api / cli
    ↓
parser → ooxml / inheritance / reading_order / headings / omml / charts / tables
    ↓
model + AssetStore
    ↓
ParsedPresentation
    ↓
renderer
    ↓
transactional bundle writer
```

The canonical document never owns a local path or binary payload. Runtime
ownership lives in `ParsedPresentation`, which closes its `AssetStore`.

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

