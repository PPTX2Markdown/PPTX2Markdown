# pptx2markdown

[![Quality](https://github.com/PPTX2Markdown/PPTX2Markdown/actions/workflows/quality.yml/badge.svg)](https://github.com/PPTX2Markdown/PPTX2Markdown/actions/workflows/quality.yml)
[![PyPI](https://img.shields.io/pypi/v/pptx2markdown)](https://pypi.org/project/pptx2markdown?cacheBust=1/)
[![Python](https://img.shields.io/pypi/pyversions/pptx2markdown)](https://pypi.org/project/pptx2markdown?cacheBust=1/)
[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/LICENSE)

English | [한국어](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/README.ko.md)

`pptx2markdown` converts PowerPoint presentations into deterministic Markdown or
JSON for search, RAG, and document-processing pipelines. It parses native OOXML
structure and shape geometry directly: no OCR, vision model, API key, or cloud
service is used.

## Why pptx2markdown?

- **One static pipeline** — `.pptx` content is parsed from XML, relationships,
  and embedded package parts.
- **Deterministic reading order** — recursive geometric XY-cut orders native
  shapes without semantic or model-based overrides.
- **One output contract** — Markdown and JSON are rendered from the same
  versioned `PresentationDocument`.
- **Rich native content** — text, headings, lists, tables, charts, SmartArt,
  formulas, images, speaker notes, and embedded attachments are retained.
- **Reviewable regressions** — real PPTX fixtures are checked against
  byte-for-byte JSON and Markdown golden outputs on Linux, macOS, and Windows.

## Supported content

| Content | Behavior |
| --- | --- |
| Text and headings | Preserves text and infers headings from native placeholder and style evidence |
| Lists | Preserves ordered and unordered list structure |
| Tables | Renders native PowerPoint tables as Markdown tables |
| Charts | Converts native charts through [chart2md](https://pypi.org/project/chart2md/) |
| SmartArt | Converts diagram content through [smartart2md](https://pypi.org/project/smartart2md/) |
| Formulas | Converts OMML equations to LaTeX through [omml2latex](https://pypi.org/project/omml2latex/) |
| Images | Copies original image assets and writes standard Markdown image links |
| Speaker notes | Keeps notes separate from visible slide content |
| Attachments | Preserves recoverable PDF, audio, video, Office, ZIP, 3D, and OLE payloads as linked files |

The parser does **not** OCR text inside images, transcribe audio or video,
reconstruct animations, interpret connector semantics, or reproduce slides
pixel-for-pixel. PowerPoint review comments are intentionally excluded.

## Requirements

- Python 3.12 or newer
- `.pptx`: no Microsoft Office or LibreOffice required
- legacy `.ppt`: Microsoft PowerPoint on Windows or LibreOffice
- optional EMF/WMF-to-PNG conversion: LibreOffice

## Installation

### Let an AI agent install it

A local coding agent can inspect the environment, choose an isolated install,
optionally configure LibreOffice, and verify the CLI. Give it one of these
prompts and review every permission request.

Most agents:

```text
Read https://raw.githubusercontent.com/PPTX2Markdown/PPTX2Markdown/main/.github/agent-install.md and install pptx2markdown for this machine. You may run the commands needed for installation after explaining them. Ask me before any administrator, sudo, password, system package-manager, shell-profile, PATH, or LibreOffice change.
```

Codex interactive planning:

```text
/plan
Read https://raw.githubusercontent.com/PPTX2Markdown/PPTX2Markdown/main/.github/agent-install.md and install pptx2markdown for this machine. You may run the commands needed for installation after explaining them. Ask me before any administrator, sudo, password, system package-manager, shell-profile, PATH, or LibreOffice change.
```

See the complete [agent installation procedure](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/.github/agent-install.md).

### Manual installation

For an isolated CLI, use `uv`:

```bash
uv tool install pptx2markdown
```

Alternatively, use `pipx` or an active virtual environment:

```bash
pipx install pptx2markdown
# or, inside an active virtual environment
python -m pip install pptx2markdown
```

LibreOffice is optional:

```bash
# macOS
brew install --cask libreoffice

# Ubuntu / Debian
sudo apt-get install libreoffice libreoffice-impress
```

On Windows, `.ppt` conversion uses PowerPoint COM automation when available and
otherwise tries LibreOffice. `SOFFICE_PATH` may point to a specific LibreOffice
executable.

## Quick start

```bash
# One file -> ./output/deck/deck.md
pptx2markdown deck.pptx

# Every .pptx/.ppt in the current directory
pptx2markdown

# Multiple files and a custom output root
pptx2markdown first.pptx second.pptx -o converted/

# Canonical JSON instead of Markdown
pptx2markdown deck.pptx --output-format json
```

Temporary Office lock files matching `~$*.pptx` are silently skipped during
batch conversion.

### Optional EMF/WMF conversion

By default, embedded EMF and WMF files are copied unchanged and linked with
standard Markdown, for example `![image](media/diagram.wmf)`. This preserves
the original presentation asset and does not start LibreOffice. Many browsers
and Markdown previewers cannot display these Windows vector formats directly.

Use the opt-in flag when the generated Markdown must display those images in a
typical browser or previewer:

```bash
pptx2markdown deck.pptx --convert-vector-images
```

With `--convert-vector-images` enabled:

- only embedded `.emf` and `.wmf` image assets are converted;
- LibreOffice runs headlessly and writes a `.png` file under `media/`;
- the Markdown/JSON image block points to the generated PNG instead of the
  original vector file;
- PNG, JPEG, GIF, SVG, WebP, and other image formats are copied unchanged;
- conversion is deferred until all slides are processed, then all unique
  EMF/WMF assets are sent to one headless LibreOffice process in a batch;
- repeated references to the same vector asset reuse one converted file; and
- if LibreOffice is unavailable or conversion fails, a warning is logged and
  each unconverted EMF/WMF file and link is preserved instead of failing the
  entire presentation or discarding successful conversions.

LibreOffice is discovered through `SOFFICE_PATH`, standard installation
locations, or the executable search path. Successful conversion does not copy
the original vector file into the output package. Set `SOFFICE_PATH` when a
specific LibreOffice installation must be used.

From Python:

```python
import pptx2markdown

exit_code = pptx2markdown.convert(
    "deck.pptx",
    output_dir="converted",
    output_format="markdown",
    convert_vector_images=True,
)
if exit_code != 0:
    raise RuntimeError("conversion failed")
```

Example Markdown:

```markdown
[Page_1]

# Quarterly results

- Revenue increased 18%
- Operating margin reached 24%

| Region | Revenue |
| --- | ---: |
| APAC | $12.4M |

![image](media/image1.png)
```

## CLI options

| Flag | Default | Description |
| --- | --- | --- |
| `-o, --output-dir` | `./output` | Final Markdown/JSON and copied assets |
| `--work-dir` | `./.pptx2markdown` | Extracted packages, analysis files, and caches |
| `--output-format` | `markdown` | `markdown` or `json` |
| `--headings` | `auto` | `auto` or placeholder-only `strict` heading detection |
| `--placeholder-inheritance` | `style` | `none`, `geometry`, or `style` inheritance |
| `--inherited-shapes` | `visible` | `none`, `visible`, or `all` layout/master shapes |
| `--ppt-converter` | `auto` | `auto`, `powerpoint`, or `libreoffice` for `.ppt` |
| `--convert-vector-images` | off | Convert embedded EMF/WMF images to PNG with LibreOffice |
| `--verbose` | off | Debug logging |

Run `pptx2markdown --help` for the authoritative CLI reference.

## Output layout

```text
output/
├── convert_manifest.json
└── <deck-name>/
    ├── <deck-name>.md       # or <deck-name>.json
    ├── media/
    └── attachments/

.pptx2markdown/
├── target_slides/
├── structure_analysis/
├── table_pipeline/
└── .cache/
```

`convert_manifest.json` records file and slide status, warnings, failures, and
block statistics. JSON output conforms to the packaged
`PresentationDocument 1.0` schema. Source paths are reduced to a basename and
generated links use portable `/` separators.

## Reading order and determinism

Reading order always uses recursive XY-cut over native shape bounding boxes. It
splits geometric regions into columns and rows, then uses top-left order when no
further cut is possible. Original XML order is only a deterministic tie-breaker
or a fallback for objects without usable geometry.

For identical input and options, JSON and Markdown outputs are expected to be
byte-for-byte stable across supported operating systems. The repository keeps
17 reviewable PPTX fixtures with both expected formats and emitted-asset hashes.

## Security model

Normal `.pptx` conversion is local and makes no network or model calls. OOXML
extraction rejects path traversal, links, encrypted entries, relationship
escapes, and packages that exceed configured archive limits. Extracted
attachments are data from the input presentation; inspect them before opening.
Legacy `.ppt` conversion invokes the selected external Office converter.
`--convert-vector-images` also invokes LibreOffice for embedded EMF/WMF assets.

## Documentation

- [Output schema](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/output_schema.md)
- [Main converter](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/main_converter.md)
- [Structure analyzer](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/structure_analyzer.md)
- [Golden PPTX suite](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/tests/fixtures/golden/README.md)

## Development

```bash
python -m pip install -e .
python -m unittest discover -v
python scripts/update_goldens.py --check
ruff check src tests scripts
ruff format --check src tests scripts
```

When parser behavior changes intentionally, review the generated Markdown and
JSON diffs before accepting new golden snapshots.

## License

Copyright 2026 HANKOOK TIRE & TECHNOLOGY CO., LTD.

The project is distributed under the [Apache License 2.0](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/LICENSE).
Third-party golden fixture attribution is recorded in
[`THIRD_PARTY_NOTICES.md`](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/tests/fixtures/golden/THIRD_PARTY_NOTICES.md).
