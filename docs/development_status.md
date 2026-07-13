# Development status and handoff

English | [한국어](development_status.ko.md)

Last verified: 2026-07-13

This document records the current engineering state. Read it together with
`git status --short --branch` before changing the worktree.

## Repository state

- Active development branch: `dev`
- Release target: `0.1.0`
- Release pull request: `dev` -> `main`, PR #17
- Public package layout: `src/pptx2markdown/`
- Packaging backend: Hatchling
- Supported Python: 3.12 and newer
- Local sample, output, work, and archive directories are not release sources

The source state before the PyPI restructure is preserved by branch
`archive-pre-pypi-restructure` and annotated tag `pre-pypi-restructure`.

## Product contract

- Conversion is static: no Surya, OCR, VLM, model provider, API key, or cloud
  inference path remains.
- XY-cut over native shape geometry is the only reading-order algorithm.
- Markdown and JSON are rendered from one `PresentationDocument 1.0`.
- No pre-release compatibility layer is maintained.
- Normal `.pptx` conversion is local. Legacy `.ppt` conversion may invoke
  PowerPoint on Windows or LibreOffice.
- Final output and intermediate work use separate configurable roots.
- Office lock files matching `~$*.pptx` are skipped silently.

## Output contract

The packaged JSON Schema is
`src/pptx2markdown/presentation_document.schema.json`. It defines:

- basename-only source identity
- ordered positive slide pages
- hidden slides and speaker notes
- closed content-block kinds
- heading-level invariants
- optional EMU geometry
- slide/layout/master provenance
- deterministic relative asset links

See [Output schema](output_schema.md).

## Regression assets

The always-on repository suite contains 17 PPTX files:

- two project-owned synthetic presentations
- 15 pinned MIT-licensed upstream presentations
- complete reviewed JSON snapshots
- complete reviewed Markdown snapshots
- input, output, and emitted-asset SHA-256 hashes
- pinned provenance and third-party notices

An additional ignored corpus is used for broad local qualification but is not a
CI dependency. See the [golden suite](../tests/fixtures/golden/README.md).

## Latest verification baseline

The following completed successfully on 2026-07-13:

- Ruff lint and format checks
- 87 unit and regression tests
- 17 repository golden PPTX files in JSON and Markdown modes
- 51 selected real-world PPTX files containing 633 slides
- wheel and sdist build plus `twine check`
- isolated wheel installation and installed-CLI conversion
- GitHub Actions on Ubuntu, macOS, and Windows with Python 3.12 and 3.13

The PR-specific workflow run for PR #17 also passed.

## Security boundaries

Shared OOXML extraction rejects traversal, symlinks, encrypted entries,
relationship escape, excessive entry counts, excessive uncompressed size, and
suspicious compression ratios. Extracted attachments remain untrusted input
data and should not be opened automatically.

## Known limitations

1. Text inside images is not OCRed.
2. Audio and video are not transcribed.
3. Connector semantics, animations, and transitions are not reconstructed.
4. Output preserves document structure, not pixel-perfect slide appearance.
5. PowerPoint review comments are intentionally excluded.
6. Broad local-corpus tests skip when ignored source files are unavailable.

## Release work remaining

1. Keep PR #17 green and complete documentation review.
2. Configure the PyPI Trusted Publisher for `publish.yml` and environment
   `pypi`.
3. Mark the pull request ready and merge it into `main`.
4. Publish the `v0.1.0` GitHub release.
5. Verify the published PyPI package in a clean environment.

## Important files

- `src/pptx2markdown/api.py`: public Python API
- `src/pptx2markdown/main_converter/run_pptx_to_markdown.py`: orchestration and CLI
- `src/pptx2markdown/main_converter/slide_converter.py`: slide-to-document parsing
- `src/pptx2markdown/main_converter/converter_models.py`: document models and renderer
- `src/pptx2markdown/structure_analyzer/structure.py`: XY-cut and heading analysis
- `src/pptx2markdown/ooxml_security.py`: shared package security policy
- `scripts/update_goldens.py`: golden verification and controlled refresh
- `tests/test_repository_goldens.py`: repository golden regression suite

## New-context checklist

```bash
git status --short --branch
python -m unittest discover -v
python scripts/update_goldens.py --check
ruff check src tests scripts
ruff format --check src tests scripts
```

Use fresh temporary output and work directories for manual conversion checks.
Preserve unrelated local changes. When a parser or schema change alters output,
review both JSON and Markdown diffs before accepting snapshots.
