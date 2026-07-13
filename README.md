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
- **Embedded attachments** — preserves PDF, audio, Office, ZIP, and OLE Packager payloads as linked files
- **Formulas** — OMML equations converted to LaTeX via [omml2latex](https://pypi.org/project/omml2latex/)
- **Images** — copied from the PPTX package as local assets and linked deterministically
- **Reading order** — deterministic recursive XY-cut on native shape geometry

## Installation

```bash
pip install pptx2markdown
```

### Let an AI agent install it

You can delegate environment detection, isolated installation, optional
LibreOffice setup, and verification to a local coding agent. Copy this prompt
to the agent and approve commands only after reviewing its explanation:

```text
Read https://raw.githubusercontent.com/PPTX2Markdown/PPTX2Markdown/main/.github/agent-install.md and install pptx2markdown for this machine. You may run the commands needed for installation after explaining them. Ask me before any administrator, sudo, password, system package-manager, shell-profile, PATH, or LibreOffice change.
```

For Codex, you may start with `/plan` so it can ask about the installation scope
and optional legacy `.ppt` support before making changes. The full agent procedure
is in [`.github/agent-install.md`](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/.github/agent-install.md).

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

Generated paths use one shared layout. Final documents are written only below
`output/`; extracted packages, structure-analysis files, table-pipeline files,
and caches are written only below `.pptx2markdown/`. Supplying `--output-dir`
or `--work-dir` moves the corresponding root without changing this layout.

```
output/
├── convert_manifest.json
└── <deck-name>/
    ├── <deck-name>.md     # or <deck-name>.json
    └── media/             # copied image assets

.pptx2markdown/
├── target_slides/         # extracted PPTX packages
├── structure_analysis/    # reordered XML and analysis sidecars
├── table_pipeline/        # standalone table-pipeline artifacts
└── .cache/                # reusable intermediate caches
```

`convert_manifest.json` records per-slide status, warnings, and block statistics.
JSON output uses the same `PresentationDocument` intermediate representation that
the Markdown renderer consumes. Slides contain ordered blocks with `kind`,
`content`, `shape_id`, optional `heading_level`, EMU `bbox`, and `source_part`
fields. The source is recorded by basename and format without an absolute path.

## Documentation

- [Converter internals](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/main_converter.md)
- [Output schema](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/output_schema.md)
- [Structure analyzer](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/structure_analyzer.md)

## License

[MIT](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/LICENSE)
