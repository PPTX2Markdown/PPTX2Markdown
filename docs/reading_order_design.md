# Reading order: containers, ambiguity and delegated judgment

[한국어](reading_order_design.ko.md)

Status: agreed design, implemented in stages (see [Plan](#plan)). Applies to the
next-generation engine in `rewrite/`; the released 0.1.x converter keeps its
current behaviour.

## Problem

The rewrite orders a slide's shapes with a recursive XY-cut over flattened
shapes (`rewrite/src/pptx2markdown/reading_order.py`). Groups are flattened into
absolute boxes, frames without text are dropped as decoration, and a region that
no gap can cut falls back to sorting by position. Three failure modes follow:

1. **A background covering the slide** (a picture, or a text box holding a
   watermark) overlaps every region, so no gap exists and the whole slide is
   sorted by position: the lines of side-by-side cards interleave.
2. **Columns are always cut before rows.** In a diagram whose tracks are
   separated by wide horizontal gaps, a narrow vertical gap left by removed
   arrows wins, and the right half of every track is read after the footer.
3. **Groups and frames carry no weight.** The frame that makes a card a card is
   discarded, and a group's members mix with shapes outside it; boxes that
   overlap by more than the one-pixel tolerance seal the gap between columns.

## Evidence

Real decks (rendered with LibreOffice and compared with the converter's order):

| Deck / slide | Layout | Rewrite today |
| --- | --- | --- |
| `brutalist_ai_newspaper_2026.pptx` 2 | Newspaper page, one text box per line, a number and its unit overlap | Cards interleave line by line; story III after the footer |
| `kubernetes_blueprint_2026.pptx` 5 | Three horizontal tracks joined by arrows | Right half of every track after the footer |
| `glassmorphism_demo.pptx` 4 | 4 × 3 cards on a full-slide picture | Whole slide sorted by position |
| `swiss_grid_systems.pptx` 7 | Three clean columns | Correct (control) |

In the corpus, 499 slides in 94 decks have overlapping or contained text
blocks, and 195 slides in 43 decks have groups holding two or more text shapes
(45 with nested groups).

`benchmarks/reading_order.py` builds eleven synthetic cases with stated
expectations ("these phrases appear in this order, with no other group's phrase
between them") and stores shapes in a shuffled order, because a real deck's XML
order follows its editing history. Baseline before this work:

| Case | rewrite | MarkItDown | anydoc |
| --- | :---: | :---: | :---: |
| Card grid on a picture that covers the slide | ❌ | ❌ | ❌ |
| Two columns on a full-slide text box holding a watermark | ❌ | ❌ | ❌ |
| Two diagram tracks, a narrow gap between columns | ❌ | ❌ | ❌ |
| Grouped cards in staggered rows | ❌ | ✅ | ❌ |
| Columns whose boxes overlap by 0.3 inch | ❌ | ❌ | ❌ |
| KPI cards with a unit box overlapping the number | ✅ | ❌ | ❌ |
| Events alternating above and below a time axis | ✅ | ❌ | ❌ |
| Callouts on a picture, caption below | ✅ | ✅ | ❌ |
| A sidebar group holding a group, next to a main group | ✅ | ✅ | ✅ |
| Labelled shapes holding text boxes | ✅ | ❌ | ❌ |
| Control: three cleanly separated columns | ✅ | ❌ | ❌ |
| **Passed** | **6/11** | **3/11** | **1/11** |

anydoc reads shapes in XML order (10/11 before the order was shuffled, 1/11
after); MarkItDown passes the grouped cards because it reads a group as a unit.

## Decisions

### Containers are the default, with a recipe switch

Reading order is a matter of correctness, not taste: lines of one card mixed
into another corrupt every chunk built from them, and an option nobody knows
about protects nobody. The rewrite is unreleased, so its default can still
change without breaking anyone.

| Part | Confidence | How it ships |
| --- | --- | --- |
| Leave backgrounds (≥ 90% of the slide) out of gap finding | High: a background is never a reading unit | Default, no switch |
| Read a frame or group that fully contains text as one unit | High: it is the visible structure | Default; a recipe can turn it off |
| Read regular grids row by row; cut along much wider gaps first | Medium: some grids are meant column by column | Default; recipe can change it; flagged for review |

The switch is a recipe setting (for example `"reading_order": "containers"` or
`"xycut"`), not a CLI flag: recipes can target single slides, sit next to the
other conversion settings, and are what review answers and the agent produce.

A part becomes the default only if corpus slides whose order changes are,
checked against their rendering, better and practically never worse; a part that
makes some layout worse ships behind its switch instead.

### No geometric rule is always right

Reading order is the author's intent, and a PPTX stores only positions, sizes,
groups and an editing-history order; the same geometry can mean different
orders. The container approach still fails on, for example, a highlight box that
spans two columns, an oversized text box that happens to enclose others, a
before/after comparison meant to be read across rows, or a half-slide photo with
unrelated text on it. The goal is therefore:

1. a default that is right most often, chosen on corpus evidence;
2. detecting where it may be wrong;
3. delegating only those places to a person or a model, whose choice is stored
   as a recipe and replays deterministically.

## Detecting ambiguity from the XML

Every signal is computed from shape boxes, without rendering:

| Signal | Meaning | Typical case |
| --- | --- | --- |
| Two-way cut | A region can be cut both vertically and horizontally, the gaps are of similar width (0.5–2×), and the two orders differ | Card grids, diagram tracks |
| No cut | A region with three or more text shapes has no gap and falls back to position sort | Newspaper pages, unremoved backgrounds |
| Partial overlap | Text boxes overlap without containment, especially across a chosen cut | Overlapping columns, number + unit |
| Container across a cut | A frame or group crosses the column boundary its parent region was cut at | A highlight box over two columns |
| Fragile gap | The separating gap is within a few pixels of the tolerance | Nearly touching boxes |

The first two fall out of the XY-cut itself: it records where it had a choice.
XML order is not a trigger (it is mostly editing history) but breaks ties when one
candidate order equals it.

A flagged slide becomes a `review()` item carrying its reasons and two or three
candidate orders (rows, columns, XML order) as ready-made recipe rules. Only
flagged slides are rendered (LibreOffice, locally) and shown to a vision model
with the candidates; the choice is a recipe rule such as
`{"target": "slide[4]", "reading_order": "rows"}` or an explicit
`{"target": "slide[4]", "order": ["s4.b3", "s4.b1", ...]}`.

Not detectable: geometrically clean layouts whose intended order differs from
the obvious one. Catching those needs every slide rendered and judged, which can
be offered as an explicit, costly option.

## Verification

- `benchmarks/reading_order.py` must reach all cases the design addresses
  (overlapping columns need a separate tolerance change).
- Corpus slides whose order changes are extracted and compared with their
  rendering before a part becomes the default.
- Ambiguity detection: every failing benchmark case and every known failing real
  slide must be flagged (recall); the flagged share of corpus slides should stay
  in single-digit percent (cost); a sample of flagged slides is checked by eye
  (precision).
- The legacy parity order score is expected to move; the baseline is updated
  only after the changed slides have been reviewed.

## Results (stages 1–3)

Benchmark: rewrite 6/11 → **9/11** (MarkItDown 3/11, anydoc 1/11). Of the two
still failing, the diagram tracks are flagged and their "rows first" candidate is
correct; overlapping columns need the partial-overlap signal and a tolerance
change.

Corpus (2,786 slides): 171 slides change text order versus the XY-cut alone.
Checked against their renderings, nearly all improve: card grids, KPI tiles,
newspaper stories and diagram panels read as units, and the Kubernetes tracks
read track by track. The remaining regressions found (a header date read last,
numbers in ovals) are both flagged for review. Five slides change a font-size
heading because a different block now comes first.

Ambiguity: 229 slides (8.2%; about 6% without one templated 369-slide deck) are
flagged, 152 for "no gap to cut" and 88 for "two-way cut". Implemented signals:
two-way cut (gaps within 2×, with each candidate following its own preference all
the way down) and no cut (only when two text boxes sit side by side; pictures do
not count). Partial overlap, containers across a cut and fragile gaps are not
yet implemented.

What did not work: cutting along the widest gap first (instead of columns first)
broke the timeline and still failed the diagram tracks (7/11): gap width alone
cannot tell columns from tracks, which is why that choice is delegated.

Lessons from the corpus: decorative groups (a panel made of two rectangles) and
brackets inside a frame must not take part in gap finding at any depth, and a
tiny box just outside a placeholder must not be treated as enclosed (the
containment tolerance is relative to the box's own size).

### All five signals

The remaining signals were added without changing any default order (0 of 2,786
slides change): *overlapping boxes* (side-by-side text overlapping by up to 15% of
the narrower box, with a "columns allowing overlap" candidate), *fragile gap*
(two real columns, two boxes or more each, parted by under two pixels) and
*container across content* (a bare frame such as a highlight box cutting through
text; text boxes and pictures overlapping text were mostly ordinary layout and
are not counted). Every failing benchmark case is now flagged with a correct
candidate (diagram tracks: rows first; overlapping columns: columns allowing
overlap). Flagged share: 9.9% (about 8.8% without the templated deck).

### Stage 4: judging flagged slides

`pptx2markdown.ordering.choose_reading_orders_with(presentation, judge)` renders
only flagged slides, each as a one-page PDF through LibreOffice's PDF export with
a page range (about 1.3 s per slide, no new dependency), and hands the judge the
reasons and candidates (first words of each block, in each order). The pick
becomes an `order` rule; a pick that names no candidate adds none. The MCP tool
`presentation_slides` shows the same slides to the host model as pictures. Without LibreOffice or a source file the judge sees text only.

## Plan

1. Leave slide-covering backgrounds out of gap finding.
2. Container tree: frames and groups that fully contain text become units,
   ordered recursively; regular grids read row by row; much wider gaps cut first.
3. Ambiguity signals as `review()` items with candidate orders; measure the
   flagged share on the corpus.
4. Reading-order recipe rules and an agent step that picks among candidates from
   the rendered slide.
