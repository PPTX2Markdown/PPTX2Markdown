# Development Status and Handoff

Last updated: 2026-07-13

This document records the current development state so a new work context can
continue without reconstructing the history from chat logs. Read this file and
`git status --short` before changing the worktree.

## Repository State

- Active branch: `dev`
- Baseline commit before the geometry-only XYCut and JSON IR work: `4919d68`.
- Check `git status --short` before editing and preserve any listed changes.
- `pptx_samples/`, `output/`, and `.pptx2markdown/` are local QA inputs and
  generated artifacts. They are not release source files.
- The version immediately before the PyPI restructure is commit `f6fc8c4`.
  Preserve it as branch `archive-pre-pypi-restructure` and annotated tag
  `pre-pypi-restructure`. A local source archive is also stored under the
  ignored `archives/` directory.

## Completed Work

### PyPI packaging

- Reorganized the package under `src/pptx2markdown` with Hatchling.
- Verified the `pptx2markdown` console entry point from a built wheel.
- Restricted sdist contents to package source, license, READMEs, and
  `pyproject.toml`.
- Removed package placeholder files from runtime target directories.
- Updated README output paths and made project links usable on PyPI.
- Built wheel and sdist offline and passed `twine check`.

### Pipeline readability refactor

- Moved presentation input and extraction concerns from the main converter to
  `main_converter/package_inputs.py`.
- Kept the table pipeline around one executable entry point and small
  parse/render APIs.
- Removed the image VLM provider layer; images are copied and linked from
  native PPTX media parts without model or API calls.
- Simplified structure-analysis reporting and inheritance resolver formatting.
- Preserved existing pipeline behavior with sample smoke tests.

### Geometry-only XYCut

Git history showed the following sequence:

- `c11839e`: introduced a Y-first XYCut implementation.
- `519d330`: removed XYCut while integrating slide/layout/master inheritance.
- `36b4e3a`: exposed the removed, pre-inheritance implementation again during
  the PyPI package restructure.

Numbering and comparison-heading overrides added during regression recovery
were not part of the original XYCut algorithm and have been removed. XYCut now:

- uses only shape bounding boxes;
- calculates X and Y projection chunks for every recursive region;
- preserves a tightly spaced single-object Y stream when its median gap is at
  most half the median object height;
- otherwise prioritizes a full-region X cut and recursively reads each column;
- uses the first Y gap to expose lower-region columns when no X cut exists;
- treats bbox overlap within one screen pixel as a shared visual boundary to
  account for PPTX coordinate rounding;
- does not use text, numbering, heading, placeholder, footer, or decorative
  classifications to change the order;
- does not synthesize missing width or height from text content;
- appends objects without usable coordinates by XML index because no geometric
  cut can be calculated for them.

XYCut is now the only reading-order strategy. The public `--reading-order`
option, legacy semantic row clustering, and Surya pipeline have been removed.
Raw XML indexes remain available for diagnostics, deterministic tie-breaking,
and missing-bbox fallback.

Regression coverage is in `tests/test_xycut_reading_order.py` and the real PPTX
fixture `tests/fixtures/xycut_tolerance_cases.pptx`. The fixture covers small
X/Y alignment jitter, one- and two-pixel boundary overlap, a quarter-pixel
positive gap, and unequal textbox heights.

### Static image handling

- Removed `image_pipeline`, all local/remote VLM providers, caches, prompts,
  API-key handling, and public `--image-vlm-*` options.
- Removed the `local-vlm`/`all` extras and their Torch, Transformers, Pillow,
  and dotenv dependency paths.
- Ordinary pictures, table overlays, and table-cell fill images remain
  supported through deterministic media extraction and asset links.

### JSON intermediate representation

- Every slide is parsed into ordered `ContentBlock` objects inside a
  `PresentationDocument` before final output is selected.
- Markdown is rendered only from that document; slide handlers no longer build
  the final Markdown file directly.
- `--output-format json` and `pptx2markdown.convert(..., output_format="json")`
  write the same document as `<deck-name>.json` instead of Markdown.
- The JSON schema records `schema_version`, source, pages, block kinds,
  content, shape IDs, and heading levels.
- Tests validate JSON round trips and confirm that rendering a JSON output
  reproduces the Markdown-mode output.

### Content regression fixes

- Chart and SmartArt relationship targets are normalized back to their OOXML
  package part beginning at `ppt/`. This supports both original and reordered
  relationship files.
- Master/layout-only objects are materialized even when the source slide has no
  direct shapes. This restores master-only text on otherwise empty slides.
- Extremely thin inherited pictures are treated as decorative rules and are
  not repeated as content images on every slide.
- `sample1.pptx` now emits its slide 9 scatter chart and slide 11 master text
  without the repeated thin stripe image.

### Heading modes

- `--headings auto` is the default and uses placeholder, numbering, font-size,
  and position evidence.
- `--headings strict` emits headings only for title/subtitle placeholders.
- The old `--not-strict` compatibility alias has been removed; the project has
  no released compatibility surface to preserve yet.
- Numbered headings tolerate spaces around periods, such as `3 . Heading`.

### Regression and quality gates

- `tests/test_content_regressions.py` covers relationship normalization,
  master-only materialization, inherited stripe filtering, and numbered
  heading syntax.
- `tests/test_sample_regressions.py` converts nine local sample decks and
  checks reading order, chart, math, table, SmartArt, inheritance, and heading
  contracts. It skips explicitly when ignored local sample files are absent.
- `tests/test_package_inputs.py` verifies that Office `~$` lock files are
  silently skipped without hiding genuinely missing inputs.
- `tests/test_xycut_pptx_fixture.py` extracts and analyzes a checked-in PPTX to
  validate both XYCut order and exact EMU threshold boundaries.
- Ruff enforces import ordering, unused-code checks, selected PEP 8 errors, a
  99-character line length, and formatting for `src` and `tests`.
- `.github/workflows/quality.yml` runs lint, format, tests, build, and Twine
  package checks on pushes and pull requests.

### Dead-code and module consolidation

- Removed the complete image-to-Markdown model pipeline after the project
  adopted static parsing only.
- Removed the duplicate table extraction module. `table_pipeline/run.py` is the
  only table pipeline executable and owns extraction.
- Removed dormant file/batch CLIs and unused HTML/CSV rendering from the table
  parse/render modules. Only the Markdown behavior used by the package remains.
- Removed the table spinner/durable-write utility after its shared callers were
  deleted; the remaining sequential writes now use `Path.write_bytes` and
  `Path.write_text` directly.
- Removed unused usage helpers, a pass-through write wrapper, and an ignored
  presentation-collection parameter.
- Reduced the source tree by four Python modules and about 730 lines without
  changing the supported `pptx2markdown` or table-pipeline entry points.

## Latest Verification Baseline

The following checks passed after the geometry-only XYCut, static image, and
JSON intermediate representation changes:

- `python -m compileall -q src/pptx2markdown tests`
- `PYTHONPATH=src python -m unittest discover -v`: 27 tests passed
- `ruff check src tests`
- `ruff format --check src tests`
- all 12 local sample decks: 128 slides converted, 0 slide failures
- 50 math blocks, 35 charts, 15 SmartArt objects, and 39 tables converted
- 25 meaningful images resolved and 0 images unresolved after inherited
  decoration filtering
- CLI help exposes `--output-format {markdown,json}` with no reading-order,
  VLM, or legacy compatibility options.
- JSON output for `reading_order_test.pptx` converted 4 slides with 0 failures.
- Re-rendering that JSON produced the same content as Markdown mode.

Wheel/sdist and Twine checks passed for commit `4919d68`. They could not be
rerun after this work item because the active virtual environment does not have
`build`, Hatchling, or Twine installed, and network access is unavailable.

Pure geometric XYCut now restores the expected order without inspecting text:

- `reading_order_test.pptx`: slide 2 keeps its tight vertical stream; slides 3
  and 4 read the left column before the right column.
- `문제점 목록 발표.pptx`: slide 15 keeps each comparison heading with the body
  below it by splitting the two geometric columns.
- `xy_cut.pptx`: touching column boundaries are separated within the coordinate
  tolerance and sections remain ordered from 1 through 7.

## Known Issues

1. Diagram connectors are not represented semantically.
2. Sample-based tests depend on ignored local `pptx_samples/` files and skip in
   a clean checkout. The always-on unit tests still cover the fixed algorithms.

## Completed Work Sequence

1. Fixed chart relationships, inherited master text, and repeated decoration.
2. Clarified and improved automatic/strict heading behavior.
3. Added unit and sample-based golden regression coverage.
4. Added Ruff and CI quality gates and completed full package verification.

The next review can focus on one pipeline at a time, starting with the main
converter. Preserve behavior with the regression suite while splitting long
or mixed-responsibility functions only where the result is easier to read.

## Important Files

- `src/pptx2markdown/main_converter/run_pptx_to_markdown.py`: conversion
  orchestration, relationship resolution, chart and SmartArt adapters.
- `src/pptx2markdown/main_converter/slide_converter.py`: slide XML parsing into
  the common intermediate document.
- `src/pptx2markdown/main_converter/converter_models.py`: intermediate models
  and Markdown renderer.
- `src/pptx2markdown/main_converter/structure_analysis_pipeline.py`: invokes
  the deterministic structure-analysis stage.
- `src/pptx2markdown/structure_analyzer/structure.py`: reading order and heading
  depth rules.
- `src/pptx2markdown/structure_analyzer/pipeline.py`: structure sidecars and
  reordered slide XML generation.
- `src/pptx2markdown/pptx_inheritance/resolver.py`: effective
  slide/layout/master object resolution.
- `tests/test_xycut_reading_order.py`: current regression tests.
- `tests/test_xycut_pptx_fixture.py`: actual PPTX coordinate-tolerance tests.
- `tests/test_content_regressions.py`: focused content-loss regression tests.
- `tests/test_intermediate_document.py`: JSON round-trip and renderer contract.
- `tests/test_sample_regressions.py`: local end-to-end sample contracts.

## New Context Startup

1. Read this document.
2. Run `git status --short` and preserve all listed changes.
3. Run `PYTHONPATH=src .venv/bin/python -m unittest discover -v`.
4. Use fresh temporary output/work directories for sample conversion so prior
   artifacts do not hide regressions.
5. Run `ruff check src tests` and `ruff format --check src tests` before and
   after refactoring a pipeline.
6. Update this document after each completed work item with changed behavior,
   tests, and remaining risks.
