"""Reading-order cases: slides whose objects overlap, nest or sit on a background.

Each case builds a deck with python-pptx and states what a reader would expect,
as constraints on where marker phrases appear in a converter's Markdown:

- ``together``: the phrases of one card, column or event appear in this order with
  no phrase of another group between them;
- ``order``: phrases that must appear in this order (title before body, track A
  before track B, 2019 before 2020).

Shapes are stored in a shuffled order (backgrounds stay at the back), because
the order of a real deck's XML follows its editing history, not its reading
order; a converter that only follows the XML would otherwise pass for free.

What has no single right answer (a card grid read row by row or column by
column) is left unconstrained. The cases come from layouts that broke the
rewrite's XY-cut on real decks: a picture covering the slide, a narrow vertical
gap between diagram columns that beats the wide gaps between tracks, one text box
per line with small overlaps, plus timelines, callouts on a picture, nested groups
and container shapes.

    python benchmarks/reading_order.py                 # every installed converter
    python benchmarks/reading_order.py --keep out/     # keep decks and Markdown
"""

from __future__ import annotations

import argparse
import io
import random
import sys
import tempfile
import traceback
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.util import Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))

from decks import _png  # noqa: E402
from run import CONVERTERS  # noqa: E402

WIDE = Inches(13.333)
HIGH = Inches(7.5)


@dataclass(frozen=True)
class Case:
    name: str
    description: str
    build: Callable[[Path], None]
    together: tuple[tuple[str, ...], ...] = ()
    order: tuple[str, ...] = ()
    phrases: tuple[str, ...] = field(default=())

    def tracked(self) -> list[str]:
        seen: dict[str, None] = {}
        for phrase in (*self.order, *(p for group in self.together for p in group)):
            seen.setdefault(phrase, None)
        return list(seen)


def _deck() -> Presentation:
    prs = Presentation()
    prs.slide_width, prs.slide_height = WIDE, HIGH
    return prs


def _slide(prs: Presentation):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _frame(shapes, shape, x: float, y: float, w: float, h: float):
    """An outlined shape with no fill, so it never hides text drawn under it."""

    frame = shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    frame.fill.background()
    return frame


def _background(shape) -> None:
    shape.name = f"Background {shape.shape_id}"


def _shuffle(path: Path, seed: str) -> None:
    """Store each slide's shapes in a scrambled order, backgrounds first."""

    prs = Presentation(path)
    rng = random.Random(seed)
    for slide in prs.slides:
        tree = slide.shapes._spTree
        shapes = [child for child in tree if child.tag.rsplit("}", 1)[-1] != "nvGrpSpPr"]
        shapes = [child for child in shapes if child.tag.rsplit("}", 1)[-1] != "grpSpPr"]
        backgrounds = [s for s in shapes if 'name="Background' in _xml(s)]
        others = [s for s in shapes if s not in backgrounds]
        rng.shuffle(others)
        for child in shapes:
            tree.remove(child)
        for child in backgrounds + others:
            tree.append(child)
    prs.save(path)


def _xml(element) -> str:
    from lxml import etree

    return etree.tostring(element, encoding="unicode")[:400]


def _text(shapes, text: str, x: float, y: float, w: float = 3, h: float = 0.4, size: int = 14):
    box = shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    box.text_frame.text = text
    for paragraph in box.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(size)
    return box


def _picture(shapes, x: float, y: float, w: float, h: float):
    return shapes.add_picture(io.BytesIO(_png()), Inches(x), Inches(y), Inches(w), Inches(h))


def _card(shapes, prefix: str, x: float, y: float, w: float = 3.6) -> None:
    """A card drawn the way exported decks do it: a frame and one text box per line."""

    _frame(shapes, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, 1.5)
    _text(shapes, f"{prefix} heading", x + 0.2, y + 0.15, w - 0.4, size=16)
    _text(shapes, f"{prefix} first line", x + 0.2, y + 0.6, w - 0.4, size=11)
    _text(shapes, f"{prefix} second line", x + 0.2, y + 0.9, w - 0.4, size=11)


def _card_phrases(prefix: str) -> tuple[str, ...]:
    return (f"{prefix} heading", f"{prefix} first line", f"{prefix} second line")


# --- cases -------------------------------------------------------------------


def build_background_picture(path: Path) -> None:
    prs = _deck()
    slide = _slide(prs)
    _background(_picture(slide.shapes, 0, 0, 13.333, 7.5))  # full-slide art, behind all
    _text(slide.shapes, "Agent layers overview", 0.6, 0.4, 8, 0.6, size=28)
    for row, y in enumerate((1.5, 3.4)):
        for column, x in enumerate((0.6, 4.6, 8.6)):
            _card(slide.shapes, f"Card {row + 1}{'ABC'[column]}", x, y)
    prs.save(path)


GRID = [f"Card {row}{column}" for row in (1, 2) for column in "ABC"]


def build_background_textbox(path: Path) -> None:
    prs = _deck()
    slide = _slide(prs)
    frame = slide.shapes.add_textbox(0, 0, WIDE, HIGH)  # a full-slide box holding a watermark
    frame.text_frame.text = "DRAFT watermark"
    _background(frame)
    _text(slide.shapes, "Quarterly plan", 0.6, 0.4, 8, 0.6, size=28)
    for index, words in enumerate(("Left column", "Right column")):
        x = 0.6 + index * 6.4
        for line in range(3):
            _text(slide.shapes, f"{words} line {line + 1}", x, 1.6 + line * 0.45, 5.5)
    prs.save(path)


def build_diagram_tracks(path: Path) -> None:
    prs = _deck()
    slide = _slide(prs)
    _text(slide.shapes, "Pod lifecycle", 0.5, 0.3, 8, 0.6, size=28)
    for track, y in (("Phase track", 1.5), ("Container track", 4.2)):
        _text(slide.shapes, track, 0.5, y, 3, 0.4, size=12)
        for step, x in enumerate((0.8, 5.0, 9.2)):
            name = f"{track.split()[0]} step {step + 1}"
            _frame(slide.shapes, MSO_SHAPE.RECTANGLE, x, y + 0.6, 3.2, 1.4)
            _text(slide.shapes, name, x + 0.2, y + 0.75, 2.8, size=16)
            _text(slide.shapes, f"{name} detail", x + 0.2, y + 1.3, 2.8, size=11)
            if step < 2:
                slide.shapes.add_connector(
                    MSO_CONNECTOR.STRAIGHT,
                    Inches(x + 3.2),
                    Inches(y + 1.3),
                    Inches(x + 4.2),
                    Inches(y + 1.3),
                )
    prs.save(path)


def _step(track: str, step: int) -> tuple[str, str]:
    name = f"{track} step {step}"
    return name, f"{name} detail"


def build_grouped_cards(path: Path) -> None:
    """Cards as groups, rows staggered so no vertical gap runs through both rows."""

    prs = _deck()
    slide = _slide(prs)
    _text(slide.shapes, "Grouped cards", 0.6, 0.4, 8, 0.6, size=28)
    for prefix, x, y in (
        ("Group A", 0.6, 1.5),
        ("Group B", 5.2, 1.5),
        ("Group C", 2.8, 3.6),
        ("Group D", 7.4, 3.6),
    ):
        group = slide.shapes.add_group_shape()
        _card(group.shapes, prefix, x, y, w=4.2)
    prs.save(path)


def build_overlapping_columns(path: Path) -> None:
    """Two columns whose boxes overlap by 0.3 inch although the text does not."""

    prs = _deck()
    slide = _slide(prs)
    _text(slide.shapes, "Two overlapping columns", 0.6, 0.4, 8, 0.6, size=28)
    for line in range(4):
        _text(slide.shapes, f"West text {line + 1}", 0.6, 1.6 + line * 0.45, 6.3)
        _text(slide.shapes, f"East text {line + 1}", 6.6, 1.7 + line * 0.45, 6.0)
    prs.save(path)


def build_big_numbers(path: Path) -> None:
    """KPI cards: a big number with its unit box overlapping it, newspaper style."""

    prs = _deck()
    slide = _slide(prs)
    _text(slide.shapes, "Issue at a glance", 0.6, 0.4, 8, 0.6, size=28)
    for index, (number, unit) in enumerate((("2.02", "T"), ("3.8", "B"), ("20.7", "B"))):
        x = 0.6 + index * 4.2
        _text(slide.shapes, f"Metric {index + 1} label", x, 1.5, 3.5, size=11)
        _text(slide.shapes, f"Metric {index + 1} value {number}", x, 1.9, 2.6, 1.0, size=40)
        _text(slide.shapes, f"Metric {index + 1} unit {unit}", x + 2.3, 2.3, 1.2, 0.6, size=20)
        _text(slide.shapes, f"Metric {index + 1} note one", x, 3.1, 3.5, size=11)
        _text(slide.shapes, f"Metric {index + 1} note two", x, 3.4, 3.5, size=11)
    prs.save(path)


def _metric(index: int) -> tuple[str, ...]:
    number = ("2.02", "3.8", "20.7")[index - 1]
    unit = ("T", "B", "B")[index - 1]
    return (
        f"Metric {index} label",
        f"Metric {index} value {number}",
        f"Metric {index} unit {unit}",
        f"Metric {index} note one",
        f"Metric {index} note two",
    )


def build_timeline(path: Path) -> None:
    """Events alternate above and below the axis; time runs left to right."""

    prs = _deck()
    slide = _slide(prs)
    _text(slide.shapes, "Company timeline", 0.6, 0.4, 8, 0.6, size=28)
    slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(0.6), Inches(3.75), Inches(12.7), Inches(3.75)
    )
    for index, year in enumerate((2019, 2020, 2021, 2022)):
        x = 0.8 + index * 3.0
        y = 2.2 if index % 2 == 0 else 4.2
        _text(slide.shapes, f"Year {year}", x, y, 2.6, size=18)
        _text(slide.shapes, f"Event of {year}", x, y + 0.5, 2.6, size=11)
    prs.save(path)


def build_callouts(path: Path) -> None:
    """Callouts placed on top of a diagram picture, with a caption below it."""

    prs = _deck()
    slide = _slide(prs)
    _text(slide.shapes, "Architecture diagram", 0.6, 0.3, 8, 0.6, size=28)
    _background(_picture(slide.shapes, 2.0, 1.2, 9.3, 4.8))
    _text(slide.shapes, "Callout gateway", 2.4, 1.5, 2.5)
    _text(slide.shapes, "Callout workers", 7.5, 3.3, 2.5)
    _text(slide.shapes, "Callout storage", 3.5, 5.2, 2.5)
    _text(slide.shapes, "Figure caption below", 2.0, 6.3, 9.3)
    prs.save(path)


def build_nested_groups(path: Path) -> None:
    prs = _deck()
    slide = _slide(prs)
    _text(slide.shapes, "Nested groups", 0.6, 0.3, 8, 0.6, size=28)
    sidebar = slide.shapes.add_group_shape()
    heading = sidebar.shapes.add_group_shape()
    _text(heading.shapes, "Sidebar title", 0.6, 1.4, 3.4, size=18)
    _text(heading.shapes, "Sidebar subtitle", 0.6, 1.9, 3.4, size=11)
    for line in range(3):
        _text(sidebar.shapes, f"Sidebar item {line + 1}", 0.6, 2.6 + line * 0.45, 3.4)
    main = slide.shapes.add_group_shape()
    for line in range(4):
        _text(main.shapes, f"Main paragraph {line + 1}", 4.5, 1.4 + line * 0.9, 8.2, 0.8)
    prs.save(path)


def build_containers(path: Path) -> None:
    """Shapes that carry their own label and hold text boxes placed inside them."""

    prs = _deck()
    slide = _slide(prs)
    _text(slide.shapes, "Two-step process", 0.6, 0.3, 8, 0.6, size=28)
    for index, x in enumerate((0.6, 6.9)):
        frame = _frame(slide.shapes, MSO_SHAPE.ROUNDED_RECTANGLE, x, 1.4, 5.8, 4.5)
        frame.text_frame.text = f"Step {index + 1} label"
        frame.text_frame.vertical_anchor = 1  # top
        for line in range(2):
            name = f"Step {index + 1} inner {line + 1}"
            _text(slide.shapes, name, x + 0.4, 2.4 + line * 1.2, 5.0)
    prs.save(path)


def build_swiss_columns(path: Path) -> None:
    """Control case: three columns of one-line boxes, cleanly separated."""

    prs = _deck()
    slide = _slide(prs)
    _text(slide.shapes, "Three typefaces", 0.6, 0.4, 8, 0.6, size=28)
    for index, name in enumerate(("Akzidenz", "Helvetica", "Univers")):
        x = 0.6 + index * 4.2
        for line in range(4):
            _text(slide.shapes, f"{name} line {line + 1}", x, 1.6 + line * 0.45, 3.6)
    prs.save(path)


CASES: tuple[Case, ...] = (
    Case(
        "background-picture",
        "Card grid on a picture that covers the slide",
        build_background_picture,
        together=tuple(_card_phrases(prefix) for prefix in GRID),
        order=("Agent layers overview", "Card 1A heading"),
    ),
    Case(
        "background-textbox",
        "Two columns on a full-slide text box holding a watermark",
        build_background_textbox,
        together=(
            tuple(f"Left column line {n}" for n in (1, 2, 3)),
            tuple(f"Right column line {n}" for n in (1, 2, 3)),
        ),
        order=("Quarterly plan", "Left column line 1", "Right column line 1"),
    ),
    Case(
        "diagram-tracks",
        "Two diagram tracks read left to right, a narrow gap between columns",
        build_diagram_tracks,
        together=tuple(
            _step(track, step) for track in ("Phase", "Container") for step in (1, 2, 3)
        ),
        order=(
            "Phase track",
            "Phase step 1",
            "Phase step 2",
            "Phase step 3",
            "Container track",
            "Container step 1",
            "Container step 2",
            "Container step 3",
        ),
    ),
    Case(
        "grouped-cards",
        "Grouped cards in staggered rows",
        build_grouped_cards,
        together=tuple(_card_phrases(f"Group {letter}") for letter in "ABCD"),
        order=("Grouped cards", *(f"Group {letter} heading" for letter in "ABCD")),
    ),
    Case(
        "overlapping-columns",
        "Columns whose boxes overlap by 0.3 inch",
        build_overlapping_columns,
        together=(
            tuple(f"West text {n}" for n in range(1, 5)),
            tuple(f"East text {n}" for n in range(1, 5)),
        ),
        order=("West text 1", "East text 1"),
    ),
    Case(
        "big-numbers",
        "KPI cards with a unit box overlapping the number",
        build_big_numbers,
        together=tuple(_metric(index) for index in (1, 2, 3)),
        order=("Metric 1 label", "Metric 2 label", "Metric 3 label"),
    ),
    Case(
        "timeline",
        "Events alternating above and below a time axis",
        build_timeline,
        together=tuple((f"Year {year}", f"Event of {year}") for year in (2019, 2020, 2021, 2022)),
        order=("Year 2019", "Year 2020", "Year 2021", "Year 2022"),
    ),
    Case(
        "callouts",
        "Callouts on a picture, caption below",
        build_callouts,
        order=(
            "Architecture diagram",
            "Callout gateway",
            "Callout workers",
            "Callout storage",
            "Figure caption below",
        ),
    ),
    Case(
        "nested-groups",
        "A sidebar group holding a group, next to a main group",
        build_nested_groups,
        together=(
            ("Sidebar title", "Sidebar subtitle", *(f"Sidebar item {n}" for n in (1, 2, 3))),
            tuple(f"Main paragraph {n}" for n in range(1, 5)),
        ),
        order=("Sidebar title", "Main paragraph 1"),
    ),
    Case(
        "containers",
        "Labelled shapes holding text boxes",
        build_containers,
        together=(
            ("Step 1 label", "Step 1 inner 1", "Step 1 inner 2"),
            ("Step 2 label", "Step 2 inner 1", "Step 2 inner 2"),
        ),
        order=("Two-step process", "Step 1 label", "Step 2 label"),
    ),
    Case(
        "swiss-columns",
        "Control: three cleanly separated columns",
        build_swiss_columns,
        together=tuple(
            tuple(f"{name} line {n}" for n in range(1, 5))
            for name in ("Akzidenz", "Helvetica", "Univers")
        ),
        order=("Akzidenz line 1", "Helvetica line 1", "Univers line 1"),
    ),
)


# --- scoring -----------------------------------------------------------------


def check(case: Case, markdown: str) -> list[str]:
    """What the Markdown gets wrong about the case; empty when it reads as expected."""

    flat = " ".join(markdown.replace("\\", "").replace("*", "").split())
    position = {phrase: flat.find(phrase) for phrase in case.tracked()}
    problems = [f"missing {phrase!r}" for phrase, index in position.items() if index < 0]
    if problems:
        return problems
    sequence = sorted(position, key=position.__getitem__)
    for earlier, later in zip(case.order, case.order[1:]):
        if position[earlier] > position[later]:
            problems.append(f"{later!r} before {earlier!r}")
    for group in case.together:
        for earlier, later in zip(group, group[1:]):
            if position[earlier] > position[later]:
                problems.append(f"{later!r} before {earlier!r}")
        indices = sorted(sequence.index(phrase) for phrase in group)
        intruders = [
            sequence[i] for i in range(indices[0], indices[-1] + 1) if sequence[i] not in group
        ]
        if intruders:
            problems.append(f"{intruders[0]!r} inside {group[0].rsplit(' ', 1)[0]!r}")
    return problems


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "--converters", nargs="*", default=list(CONVERTERS), choices=list(CONVERTERS)
    )
    parser.add_argument("--keep", type=Path, help="Directory for decks and outputs")
    parser.add_argument("--verbose", action="store_true", help="List every problem")
    args = parser.parse_args(argv)

    converters = {}
    for name in args.converters:
        try:
            converters[name] = CONVERTERS[name][1]()
        except ImportError as exc:
            print(f"skipping {name}: {exc}", file=sys.stderr)

    work = args.keep or Path(tempfile.mkdtemp(prefix="pptx-reading-order-"))
    work.mkdir(parents=True, exist_ok=True)
    print("| Case | " + " | ".join(converters) + " |")
    print("| --- | " + " | ".join(":---:" for _ in converters) + " |")
    totals = dict.fromkeys(converters, 0)
    details: list[str] = []
    for case in CASES:
        deck = work / f"{case.name}.pptx"
        case.build(deck)
        _shuffle(deck, case.name)
        cells = []
        for name, convert in converters.items():
            try:
                markdown = convert(deck)
            except Exception as exc:  # a converter failing is a result
                (work / f"{case.name}.{name}.error.txt").write_text(
                    "".join(traceback.format_exception(exc)), encoding="utf-8"
                )
                cells.append("💥")
                continue
            (work / f"{case.name}.{name}.md").write_text(markdown, encoding="utf-8")
            problems = check(case, markdown)
            totals[name] += not problems
            cells.append("✅" if not problems else "❌")
            details.extend(f"- {case.name} / {name}: {problem}" for problem in problems[:3])
        print(f"| {case.description} | " + " | ".join(cells) + " |")
    print(
        "| **Passed** | "
        + " | ".join(f"**{totals[name]}/{len(CASES)}**" for name in converters)
        + " |"
    )
    if args.verbose and details:
        print("\n" + "\n".join(details))
    print(f"\nDecks and outputs: {work}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
