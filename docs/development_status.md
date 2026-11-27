# Development status and handoff

English | [한국어](development_status.ko.md)

Last verified: 2026-11-27

This document records the current engineering state. Read it together with
`git status --short --branch` before changing the worktree.

## Repository state

- Active development branch: `dev` (pull requests merge into `dev`; releases
  go `dev` -> `main`, then a GitHub release publishes to PyPI)
- Released: `0.2.0`, the engine that replaced the 0.1 converter; changes are in
  [CHANGELOG.md](../CHANGELOG.md)
- The 0.1 converter is kept on branch `legacy-0.1` (bug fixes only, no release
  planned); the removed Claude agent layer is kept on `archive/agent` for a
  separate multi-format agent project
- Package layout: `src/pptx2markdown/`, Hatchling, Python 3.12 and newer

## Product contract

- Conversion is static and model-free: the package imports no model SDK (a test
  enforces it). Judgement plugs in through recipes, `review()` answers and the
  judge loops; see [the converter–agent contract](agent_contract.md).
- The same input and options give the same bytes on every supported OS.
- `.pptx` conversion is local. LibreOffice runs only when asked for: `.ppt`
  input, EMF/WMF as PNG, rendering flagged slides.
- Nothing is written outside the output bundle.

See [ARCHITECTURE.md](ARCHITECTURE.md) for ownership, lifecycle and security,
and [the reading-order design](reading_order_design.md).

## Regression assets

- `tests/fixtures/golden/`: 17 PPTX files (two project-owned synthetic decks,
  15 pinned MIT-licensed upstream decks) with reviewed Markdown and JSON
  snapshots and asset hashes; `scripts/update_goldens.py --check` and
  `tests/test_goldens.py` compare them byte for byte. See the
  [golden suite](../tests/fixtures/golden/README.md).
- `tests/fixtures/xycut_*`: expected reading orders.
- `benchmarks/`: feature checks against MarkItDown and anydoc (20 of 20 kept,
  anydoc 13, MarkItDown 8) and synthetic reading-order cases (9 of 11, the
  other two flagged with the right candidate).
- An ignored local corpus (`.pptx2markdown/corpus/`, 966 decks) for broad
  qualification; not a CI dependency.

Before 0.2.0 the engine was checked shape by shape against the 0.1 converter
(`scripts/parity.py`, now on `legacy-0.1`): 99.8% of legacy's words kept on
the corpus, and no lost text found by `audit()` in any corpus deck.

## Known limitations

1. Text inside images is not OCRed; audio and video are not transcribed.
2. Connector semantics, animations and transitions are not reconstructed.
3. Tables drawn from the slide layout or master are not read.
4. PowerPoint review comments are intentionally excluded.

## Next steps

1. Build the multi-format agent as its own project on the decision-point
   contract (start from `archive/agent`).
2. Validate reading order and review items on Korean business decks; the local
   corpus has almost none.
3. Make omml2latex run on Python 3.10/3.11 to widen supported versions.

## New-context checklist

```bash
git status --short --branch
python -m pip install -e ".[mcp]"
python -m unittest discover -v
python scripts/update_goldens.py --check
ruff check src tests scripts benchmarks
ruff format --check src tests scripts benchmarks
```
