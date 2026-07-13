# Main converter

English | [한국어](main_converter.ko.md)

The main converter turns one or more PowerPoint files into the common
`PresentationDocument` representation and then renders Markdown or JSON. The
installed entry point is `pptx2markdown`; the implementation lives in
`src/pptx2markdown/main_converter/`.

## Pipeline

1. Resolve `.pptx` and legacy `.ppt` inputs.
2. Convert `.ppt` to `.pptx` with the selected external converter.
3. Extract and validate the OOXML package below the work directory.
4. Resolve slide, layout, and master inheritance.
5. Order native shapes with recursive geometric XY-cut.
6. Parse shapes into ordered `ContentBlock` objects.
7. Build and validate one `PresentationDocument`.
8. Render either deterministic Markdown or canonical JSON.
9. Write a batch conversion manifest.

Markdown and JSON do not use separate parsers.

## Input policy

With explicit arguments, each `.pptx` or `.ppt` selection is processed:

```bash
pptx2markdown first.pptx second.ppt
```

Without arguments, presentations in the current directory and the work
directory's `target_pptx/` staging directory are discovered:

```bash
pptx2markdown
```

Temporary Office lock files matching `~$*.pptx` are skipped silently. A normal
missing input remains an error.

## Output policy

The default final root is `./output` and the default work root is
`./.pptx2markdown`.

```text
output/
├── convert_manifest.json
└── <deck>/
    ├── <deck>.md           # or <deck>.json
    ├── media/
    └── attachments/

.pptx2markdown/
├── target_pptx/
├── target_slides/
├── structure_analysis/
├── table_pipeline/
└── .cache/
```

Final documents and copied assets are written only under the output root.
Extracted packages, analysis sidecars, and caches are written only under the
work root.

## Reading order

Every slide uses recursive XY-cut over native shape bounding boxes. There is no
public reading-order option. Text, numbering, headings, placeholders, and
decorative classifications do not change ordering. XML index is used only for
deterministic ties and objects without usable geometry.

## Public CLI options

```text
--output-dir <path>
--work-dir <path>
--output-format markdown|json
--headings auto|strict
--placeholder-inheritance none|geometry|style
--inherited-shapes none|visible|all
--ppt-converter auto|powerpoint|libreoffice
--verbose
```

Use `pptx2markdown --help` as the authoritative option reference.

## Programmatic API

```python
import pptx2markdown

exit_code = pptx2markdown.convert(
    ["first.pptx", "second.pptx"],
    output_dir="converted",
    work_dir="work",
    output_format="json",
)
```

The API returns `0` when every selected slide converts successfully and `1` if
any slide fails.

## Related documentation

- [Output schema](output_schema.md)
- [Structure analyzer](structure_analyzer.md)
- [Repository golden suite](../tests/fixtures/golden/README.md)
