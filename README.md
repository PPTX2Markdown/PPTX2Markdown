# pptx2markdown

[![PyPI](https://img.shields.io/pypi/v/pptx2markdown)](https://pypi.org/project/pptx2markdown/)
[![Python](https://img.shields.io/pypi/pyversions/pptx2markdown)](https://pypi.org/project/pptx2markdown/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/LICENSE)

Convert PowerPoint (`.pptx` / `.ppt`) presentations into clean, structured Markdown — built for RAG pipelines and document processing.

[한국어 README](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/README.ko.md)

## Features

- **Text & headings** — heading levels inferred from slide structure, font size, and placeholder inheritance (slide → layout → master)
- **Tables** — native PPTX tables rendered as Markdown tables
- **Charts & SmartArt** — converted to Markdown via [chart2md](https://pypi.org/project/chart2md/) and [smartart2md](https://pypi.org/project/smartart2md/)
- **Formulas** — OMML equations converted to LaTeX via [omml2latex](https://pypi.org/project/omml2latex/)
- **Images** — copied from the PPTX package as local assets and linked deterministically
- **Reading order** — deterministic recursive XY-cut on native shape geometry

## Installation

```bash
pip install pptx2markdown
```

**LibreOffice** (optional) is used to convert legacy `.ppt` inputs and EMF/WMF images:

```bash
# macOS
brew install libreoffice
# Ubuntu / Debian
sudo apt-get install -y libreoffice libreoffice-impress
```

On Windows, PowerPoint COM automation is used for `.ppt` conversion when available; otherwise set `SOFFICE_PATH` to your LibreOffice executable.

## Quick start

```bash
# Convert one file → ./output/deck/deck.md
pptx2markdown deck.pptx

# Convert every .pptx/.ppt in the current directory
pptx2markdown

# Choose the output directory
pptx2markdown deck.pptx -o converted/

# Write the canonical intermediate representation instead of Markdown
pptx2markdown deck.pptx --output-format json
```

Or from Python:

```python
import pptx2markdown

pptx2markdown.convert("deck.pptx", output_dir="converted")
```

## Options

### Reading order

Reading order is always determined by recursive XY-cut on shape bounding boxes.
The algorithm uses top-left ordering when no further geometric cut is possible
and the original XML index only as a deterministic tie-breaker or missing-bbox fallback.

### Other flags

| Flag | Default | Description |
| --- | --- | --- |
| `-o, --output-dir` | `./output` | Where converted output is written |
| `--work-dir` | `./.pptx2markdown` | Intermediate files (extraction, caches) |
| `--output-format` | `markdown` | Final output (`markdown`/`json`) |
| `--headings` | `auto` | Heading detection (`auto`/`strict`) |
| `--placeholder-inheritance` | `style` | How much layout/master style to inherit (`none`/`geometry`/`style`) |
| `--inherited-shapes` | `visible` | Materialize layout/master shapes (`none`/`visible`/`all`) |
| `--ppt-converter` | `auto` | `.ppt` conversion backend (`powerpoint`/`libreoffice`) |
| `--verbose` | off | Debug logging |

Run `pptx2markdown --help` for the full list.

## Output

```
output/
├── convert_manifest.json
└── <deck-name>/
    ├── <deck-name>.md     # or <deck-name>.json
    └── media/             # copied image assets
```

`convert_manifest.json` records per-slide status, warnings, and block statistics.
JSON output uses the same `PresentationDocument` intermediate representation that
the Markdown renderer consumes. Slides contain ordered blocks with `kind`,
`content`, `shape_id`, and optional `heading_level` fields.

## Documentation

- [Converter internals](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/main_converter.md)
- [Structure analyzer](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/structure_analyzer.md)

## License

[MIT](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/LICENSE)
