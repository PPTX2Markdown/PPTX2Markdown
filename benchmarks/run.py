"""Run PPTX-to-Markdown converters over feature decks and report what each preserves.

    python benchmarks/run.py                  # all converters that are installed
    python benchmarks/run.py --large 300      # also time a 300-slide deck
    python benchmarks/run.py --keep out/      # keep decks and outputs for inspection

See benchmarks/README.md for setup.
"""

from __future__ import annotations

import argparse
import json
import platform
import sys
import tempfile
import time
import traceback
from collections.abc import Callable
from importlib import metadata
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from decks import CASES, build_large  # noqa: E402

Converter = Callable[[Path], str]


def _pptx2markdown() -> Converter:
    import pptx2markdown

    def convert(path: Path) -> str:
        with pptx2markdown.parse(path) as presentation:
            return presentation.render_markdown(asset_url=lambda asset_id: f"media/{asset_id}")

    return convert


def _markitdown() -> Converter:
    from markitdown import MarkItDown

    converter = MarkItDown()
    return lambda path: converter.convert(str(path)).text_content


def _anydoc() -> Converter:
    import anydoc

    return lambda path: anydoc.to_markdown(path)


CONVERTERS: dict[str, tuple[str, Callable[[], Converter]]] = {
    "pptx2markdown (rewrite)": ("pptx2markdown", _pptx2markdown),
    "markitdown": ("markitdown", _markitdown),
    "anydoc": ("firecrawl-anydoc", _anydoc),
}


def _version(distribution: str) -> str:
    try:
        return metadata.version(distribution)
    except metadata.PackageNotFoundError:
        return "?"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "--converters", nargs="*", default=list(CONVERTERS), choices=list(CONVERTERS)
    )
    parser.add_argument("--large", type=int, default=0, help="Time a deck with N slides")
    parser.add_argument("--keep", type=Path, help="Directory for decks and outputs")
    parser.add_argument(
        "--out",
        type=Path,
        default=Path(__file__).resolve().parent / "results",
        help="Where results.md and results.json go",
    )
    args = parser.parse_args(argv)

    converters: dict[str, Converter] = {}
    versions: dict[str, str] = {}
    for name in args.converters:
        distribution, factory = CONVERTERS[name]
        try:
            converters[name] = factory()
            versions[name] = _version(distribution)
        except ImportError as exc:
            print(f"skipping {name}: {exc}", file=sys.stderr)

    work = args.keep or Path(tempfile.mkdtemp(prefix="pptx-bench-"))
    work.mkdir(parents=True, exist_ok=True)
    results: dict[str, dict[str, str]] = {}
    for case in CASES:
        deck = work / f"{case.name}.pptx"
        case.build(deck)
        results[case.name] = {}
        for name, convert in converters.items():
            try:
                markdown = convert(deck)
            except Exception as exc:  # a converter failing is a result, not a crash
                results[case.name][name] = "error"
                (work / f"{case.name}.{name}.error.txt").write_text(
                    "".join(traceback.format_exception(exc)), encoding="utf-8"
                )
                continue
            (work / f"{case.name}.{name}.md").write_text(markdown, encoding="utf-8")
            results[case.name][name] = "pass" if case.check(markdown) else "fail"

    timings: dict[str, float | str] = {}
    if args.large:
        deck = work / f"large-{args.large}.pptx"
        build_large(deck, args.large)
        for name, convert in converters.items():
            try:
                start = time.perf_counter()
                convert(deck)
                timings[name] = round(time.perf_counter() - start, 3)
            except Exception:
                timings[name] = "error"

    args.out.mkdir(parents=True, exist_ok=True)
    report = {
        "python": platform.python_version(),
        "platform": platform.platform(terse=True),
        "versions": versions,
        "cases": {case.name: case.feature for case in CASES},
        "results": results,
        "timings": {"slides": args.large, "seconds": timings} if args.large else None,
    }
    (args.out / "results.json").write_text(
        json.dumps(report, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )
    (args.out / "results.md").write_text(render(report), encoding="utf-8")
    print(render(report))
    print(f"decks and outputs: {work}", file=sys.stderr)
    return 0


def render(report: dict) -> str:
    names = list(report["versions"])
    symbol = {"pass": "✅", "fail": "❌", "error": "💥"}
    header = " | ".join(f"{name} {report['versions'][name]}" for name in names)
    lines = [
        "| Feature | " + header + " |",
        "| --- |" + " :---: |" * len(names),
    ]
    totals = dict.fromkeys(names, 0)
    for case, feature in report["cases"].items():
        row = report["results"][case]
        for name in names:
            totals[name] += row[name] == "pass"
        lines.append(f"| {feature} | " + " | ".join(symbol[row[name]] for name in names) + " |")
    count = len(report["cases"])
    lines.append("| **Passed** | " + " | ".join(f"**{totals[n]}/{count}**" for n in names) + " |")
    if report["timings"]:
        seconds = report["timings"]["seconds"]
        lines += [
            "",
            f"Time for a {report['timings']['slides']}-slide deck (text, table per slide), "
            f"Python {report['python']}, {report['platform']}:",
            "",
            "| Converter | Seconds |",
            "| --- | ---: |",
            *(f"| {name} | {seconds.get(name, '-')} |" for name in names),
        ]
    return "\n".join(lines) + "\n"


if __name__ == "__main__":
    raise SystemExit(main())
