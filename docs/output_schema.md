# PresentationDocument 2.0

English | [한국어](output_schema.ko.md)

`--output-format json` and `ParsedPresentation.render_json()` write one
`PresentationDocument`, the same model the Markdown is rendered from. Its JSON
Schema comes from the package:

```python
from pptx2markdown import PresentationDocument

schema = PresentationDocument.model_json_schema()
```

## Top level

| Field | Meaning |
| --- | --- |
| `schema_version` | `"2.0"`; bumped when a field changes meaning or goes away |
| `source` | the deck's base name, format and SHA-256 (never a full path) |
| `assets` | images and attachments: id, kind, media type, original name, size, SHA-256, source part |
| `slides` | slides in presentation order |
| `warnings` | structured warnings about content that could not be converted |

## Slides

| Field | Meaning |
| --- | --- |
| `page` | 1-based position in the presentation's slide list |
| `hidden` | whether the slide is hidden in PowerPoint |
| `section` | the PowerPoint section the slide belongs to, if any |
| `blocks` | content blocks in reading order |
| `notes` | speaker notes |
| `reading_order` | when the layout can be read more than one way: the reasons and candidate orders |

Every block has an `id` (`s<page>.b<n>`), a `kind` and, when the shape has one,
its `bbox` in EMU and `shape_id`. Kinds: `heading` (with `level` and `source`),
`text`, `list` (items with level and ordering), `table` (the full `rows` grid with
merge spans next to the rendered `markdown`), `image` and `attachment` (an
`asset_id`), `chart`, `smartart`, `math` (LaTeX) and `unsupported` (what could not
be converted, never silently dropped).

The bundle written next to the document has a `bundle-manifest.json` with the
source hash and the path, size and SHA-256 of every saved asset.
