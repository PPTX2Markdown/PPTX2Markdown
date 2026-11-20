# Development status and handoff

English | [한국어](development_status.ko.md)

Last verified: 2026-11-21

This document records the current engineering state. Read it together with
`git status --short --branch` before changing the worktree.

## Repository state

- Active development branch: `dev` (pull requests merge into `dev`; releases
  go `dev` -> `main`)
- Released: `0.1.2` on PyPI; unreleased changes on `dev` are listed in
  [CHANGELOG.md](../CHANGELOG.md)
- Next-generation engine: `rewrite/`, unreleased, see below
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
- Office lock files matching `~$*.pptx` are skipped; a named one is reported
  with a warning.

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

## Next-generation engine (`rewrite/`)

`rewrite/` is a separate package (also named `pptx2markdown`, version
`0.2.0.dev0`) built to replace `src/` once it covers the same content. It keeps
the static pipeline and adds escaped inline Markdown with emphasis, structured
tables, block ids with declarative recipes, a review of uncertain decisions, an
optional agent and MCP server, slide chunks, and in-memory input. See
[rewrite/README.md](../rewrite/README.md) and
[rewrite/ARCHITECTURE.md](../rewrite/ARCHITECTURE.md).

`scripts/parity.py` converts the same decks with both engines and compares them
shape by shape (legacy is the reference). CI fails when the rewrite does worse
than `scripts/parity_baseline.json`; deliberate differences are set back to
legacy behaviour through `scripts/parity_legacy_recipe.json`. On the repository
fixtures 60 of 62 slides match; the two remaining differences are deliberate.

## Latest verification baseline

The following completed successfully on 2026-11-21:

- Ruff lint and format checks for `src/`, `tests/`, `scripts/` and `rewrite/`
- 127 legacy unit and regression tests, 17 golden PPTX files in both modes
- about 150 rewrite tests on Python 3.12, 3.13 and 3.14
- legacy/rewrite parity gate (60/62 fixture slides)
- GitHub Actions on Ubuntu, macOS and Windows

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

## Next steps

1. Review `dev` and release the unreleased changes in CHANGELOG.md as `0.1.3`
   (merge `dev` into `main`, tag, publish).
2. Decide when `rewrite/` replaces `src/` (for example: parity on the local
   corpus, `.ppt` input policy, migration notes for JSON schema 2.0).
3. Make omml2latex run on Python 3.10/3.11 (one f-string with a backslash in
   its expression) to widen supported Python versions.

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
