# pptx2markdown

[![PyPI](https://img.shields.io/pypi/v/pptx2markdown)](https://pypi.org/project/pptx2markdown/)
[![Python](https://img.shields.io/pypi/pyversions/pptx2markdown)](https://pypi.org/project/pptx2markdown/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Convert PowerPoint (`.pptx` / `.ppt`) presentations into clean, structured Markdown — built for RAG pipelines and document processing.

[한국어 README](README.ko.md)

## Features

- **Text & headings** — heading levels inferred from slide structure, font size, and placeholder inheritance (slide → layout → master)
- **Tables** — native PPTX tables rendered as Markdown tables
- **Charts & SmartArt** — converted to Markdown via [chart2md](https://pypi.org/project/chart2md/) and [smartart2md](https://pypi.org/project/smartart2md/)
- **Formulas** — OMML equations converted to LaTeX via [omml2latex](https://pypi.org/project/omml2latex/)
- **Images** — copied as assets, or described in Markdown by a vision-language model (Gemini / OpenAI / OpenRouter / local Qwen2.5-VL)
- **Reading order** — XML order by default, with optional XY-cut or [Surya](https://github.com/VikParuchuri/surya) layout-based reordering

## Installation

```bash
pip install pptx2markdown
```

Optional extras (each pulls in PyTorch — multi-GB install):

```bash
pip install "pptx2markdown[surya]"      # --reading-order surya
pip install "pptx2markdown[local-vlm]"  # --image-vlm-provider local (Qwen2.5-VL)
pip install "pptx2markdown[all]"        # everything
```

**LibreOffice** (optional) is used to convert legacy `.ppt` inputs and EMF/WMF images, and to render PDFs for Surya mode:

```bash
# macOS
brew install libreoffice
# Ubuntu / Debian
sudo apt-get install -y libreoffice libreoffice-impress
```

On Windows, PowerPoint COM automation is used for `.ppt` conversion when available; otherwise set `SOFFICE_PATH` to your LibreOffice executable.

## Quick start

```bash
# Convert one file → ./output/xml/deck/result.md
pptx2markdown deck.pptx

# Convert every .pptx/.ppt in the current directory
pptx2markdown

# Choose the output directory
pptx2markdown deck.pptx -o converted/
```

Or from Python:

```python
import pptx2markdown

pptx2markdown.convert("deck.pptx", output_dir="converted")
```

## Options

### Reading order

```bash
pptx2markdown deck.pptx --reading-order xml    # default: slide XML order
pptx2markdown deck.pptx --reading-order xycut  # recursive XY-cut on shape geometry
pptx2markdown deck.pptx --reading-order surya  # Surya layout model (requires [surya] extra)
```

### Image description with a VLM

By default images are linked as assets. Pass `--image-vlm-provider` to describe document-like images (tables, diagrams, screenshots) as Markdown instead:

```bash
# Gemini — put GEMINI_API_KEY in .env or the environment
pptx2markdown deck.pptx --image-vlm-provider gemini

# OpenAI (default model: gpt-4.1-mini) — needs OPENAI_API_KEY
pptx2markdown deck.pptx --image-vlm-provider openai

# OpenRouter (default model: google/gemini-2.5-flash) — needs OPENROUTER_API_KEY
pptx2markdown deck.pptx --image-vlm-provider openrouter

# Local Qwen2.5-VL — requires [local-vlm] extra, GPU recommended
pptx2markdown deck.pptx --image-vlm-provider local --image-vlm-model 3b
```

VLM results are cached in `~/.cache/pptx2markdown/` (override with `PPTX2MARKDOWN_CACHE_DIR`); pass `--ignore-image-vlm-cache` to recompute.

### Other flags

| Flag | Default | Description |
| --- | --- | --- |
| `-o, --output-dir` | `./output` | Where converted Markdown is written |
| `--work-dir` | `./.pptx2markdown` | Intermediate files (extraction, caches) |
| `--not-strict` | off | Relaxed heading detection |
| `--placeholder-inheritance` | `style` | How much layout/master style to inherit (`none`/`geometry`/`style`) |
| `--inherited-shapes` | `visible` | Materialize layout/master shapes (`none`/`visible`/`all`) |
| `--ppt-converter` | `auto` | `.ppt` conversion backend (`powerpoint`/`libreoffice`) |
| `--verbose` | off | Debug logging |

Run `pptx2markdown --help` for the full list.

## Output

```
output/
└── xml/                    # one folder per reading-order mode
    ├── convert_manifest.json
    └── <deck-name>/
        ├── result.md
        └── media/          # copied image assets
```

`convert_manifest.json` records per-slide status, warnings, and block statistics.

## Documentation

- [Converter internals](docs/main_converter.md)
- [Image VLM pipeline](docs/image_pipeline.md)
- [Structure analyzer](docs/structure_analyzer.md)
- [Surya pipeline](docs/surya_pipeline.md)

## License

[MIT](LICENSE)
