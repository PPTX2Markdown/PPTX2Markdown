# PPTX converter benchmark

[한국어](README.ko.md)

What each converter keeps of a PowerPoint deck, checked feature by feature.

`decks.py` builds one small deck per feature with python-pptx (nothing binary is
stored, so every deck can be rebuilt and opened in PowerPoint), and pairs it with
a check on the Markdown a converter returns. `run.py` runs every installed
converter on every deck and writes `results/results.md` and `results/results.json`.

```bash
python3.12 -m venv .bench && . .bench/bin/activate
pip install -r benchmarks/requirements.txt
python benchmarks/run.py --large 300 --keep /tmp/bench   # keep decks and outputs
```

`reading_order.py` checks reading order on slides whose objects overlap, nest or
sit on a background (shapes are stored in a shuffled order, as in real decks):

```bash
python benchmarks/reading_order.py --keep /tmp/reading-order --verbose
```

See [the reading-order design](../docs/reading_order_design.md) for the cases and
what they found.

## Results

See [results/results.md](results/results.md) for the latest run (versions,
Python and platform are recorded in `results.json`).

## Reading the results

- The features were chosen for what presentation-aware retrieval needs: slide
  boundaries for citations, headings, reading order, tables and merged cells,
  equations, notes, hidden slides, image descriptions. They are the features
  pptx2markdown is built around, so the selection favours it; a general-purpose
  converter makes different trade-offs.
- anydoc is a Rust library covering 14 formats and is about 20 times faster.
  If speed across many formats matters more than slide structure, that is a real
  advantage.
- Each check is a single property of a small synthetic deck. Real decks mix
  everything, so this measures coverage of a feature, not output quality on a
  corpus.
- A check accepts any reasonable spelling of the right result (`**x**` or `__x__`,
  `mc^{2}` or `m{c}^{2}`, any slide marker). If a converter keeps the information
  in a form a check misses, that is a bug in the check: please report it.
