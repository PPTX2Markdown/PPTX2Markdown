# Structure analyzer

English | [한국어](structure_analyzer.ko.md)

The structure analyzer reads slide XML, resolves optional placeholder
inheritance, calculates deterministic reading order, and writes diagnostic JSON
plus reordered XML. The main converter invokes this stage automatically.

## Module layout

- `constants.py`: OOXML namespaces, tags, and placeholder constants
- `xml_primitives.py`: shared XML and bounding-box helpers
- `text_rules.py`: text normalization and heading syntax rules
- `extractor.py`: XML-to-`SlideObject` extraction
- `structure.py`: `SlideObject`, XY-cut ordering, and heading analysis
- `pipeline.py`: analysis reports, reordered XML, manifests, and file I/O
- `extract_structure_analysis.py`: diagnostic CLI
- `check_native_table_support.py`: native-table inspection utility

## Reading-order policy

Reading order always uses recursive geometric XY-cut. It does not inspect text,
numbering, headings, placeholder roles, or decorative labels to change order.

For each recursive region, the algorithm examines X and Y projections. It
preserves a tightly spaced vertical stream when appropriate, otherwise exposes
column cuts before recursively processing lower regions. When geometry cannot
split a region further, objects use top-left order with XML index as a stable
tie-breaker.

Coordinate tolerance accounts for small OOXML rounding differences. The
checked-in PPTX regressions cover 1–3 px alignment jitter, small positive
gutters, one-pixel boundary overlap, overlap beyond the tolerance, and mixed
text-box heights. A separate semantic row-clustering pass is not used.

## Placeholder inheritance

Objects with incomplete direct geometry or style may inherit matching
placeholder information through:

```text
slide -> slideLayout -> slideMaster
```

The default main-converter policy is `style` inheritance with `visible`
layout/master-only shapes.

## Diagnostic CLI

Process one or more slide XML files:

```bash
python -m pptx2markdown.structure_analyzer.extract_structure_analysis \
  slide1.xml slide2.xml
```

With no input, the command scans `<work-dir>/target_slides`:

```bash
python -m pptx2markdown.structure_analyzer.extract_structure_analysis
```

Select a work root or output directory:

```bash
python -m pptx2markdown.structure_analyzer.extract_structure_analysis \
  slide1.xml --work-dir .pptx2markdown --output-dir analysis
```

## Diagnostic outputs

- `<slide>.structure_analysis.json`
- `<slide>.reordered.xml`
- `structure_analysis_manifest.json`

The JSON report includes the schema version, geometric structure order, raw XML
order, ordered XML indexes, object counts, confidence information, and detected
native tables and images.

## Options

```text
--output-dir <path>
--work-dir <path>
--strict
--placeholder-inheritance none|geometry|style
--inherited-shapes none|visible|all
```

`--strict` changes heading detection and candidate thresholds; it does not
select another reading-order algorithm.

## Regression coverage

- `tests/test_xycut_reading_order.py`: focused coordinate cases
- `tests/test_xycut_pptx_fixture.py`: real PPTX coordinate-tolerance fixture
- `tests/test_xycut_layout_pptx_fixture.py`: columns, dividers, rotations,
  tables, ties, schema, and byte determinism
- `tests/fixtures/golden/synthetic_reading_order.pptx`: reviewed repository
  golden output

See the [golden suite](../tests/fixtures/golden/README.md) for acceptance policy.
