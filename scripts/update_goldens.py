"""Check or accept the golden Markdown/JSON snapshots of the repository PPTX suite.

Every deck listed in ``tests/fixtures/golden/manifest.json`` is converted to a
Markdown and a JSON bundle; the documents must match ``expected/`` byte for byte
and every saved asset must match the hashes in the manifest.

    python scripts/update_goldens.py --check    # exit 1 and list differences
    python scripts/update_goldens.py --accept   # rewrite snapshots and hashes
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
import tempfile
from collections import Counter
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parents[1]
GOLDEN = REPO_ROOT / "tests" / "fixtures" / "golden"
MANIFEST = GOLDEN / "manifest.json"
EXPECTED = {"markdown": GOLDEN / "expected" / "markdown", "json": GOLDEN / "expected" / "json"}
SUFFIX = {"markdown": ".md", "json": ".json"}


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _convert(source: Path, output_format: str, root: Path) -> tuple[bytes, dict[str, str]]:
    import pptx2markdown

    result = pptx2markdown.convert(
        source, output_dir=root / output_format, output_format=output_format
    )
    assets = {
        asset.path.relative_to(result.bundle_dir).as_posix(): asset.sha256
        for asset in result.assets
    }
    return result.document_path.read_bytes(), assets


def render_suite() -> tuple[dict[str, Any], dict[str, dict[str, Any]]]:
    """Each case's documents, asset hashes and block counts, as converted now."""

    import pptx2markdown

    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
    rendered: dict[str, dict[str, Any]] = {}
    with tempfile.TemporaryDirectory() as directory:
        for case in manifest["cases"]:
            source = GOLDEN / case["file"]
            markdown, assets = _convert(source, "markdown", Path(directory) / case["file"])
            document, _ = _convert(source, "json", Path(directory) / case["file"])
            with pptx2markdown.parse(source, include_assets=False) as presentation:
                kinds = Counter(
                    block.kind for slide in presentation.slides for block in slide.blocks
                )
                slides = len(presentation.slides)
            rendered[case["file"]] = {
                "markdown": markdown,
                "json": document,
                "assets": assets,
                "slides": slides,
                "kinds": kinds,
            }
    return manifest, rendered


def check() -> list[str]:
    """Differences between the current conversion and the accepted snapshots."""

    manifest, rendered = render_suite()
    problems: list[str] = []
    for case in manifest["cases"]:
        name, now = case["file"], rendered[case["file"]]
        source = GOLDEN / name
        if _sha256(source.read_bytes()) != case["sha256"]:
            problems.append(f"{name}: the input deck does not match its manifest hash")
        for output_format, directory in EXPECTED.items():
            expected = directory / (Path(name).stem + SUFFIX[output_format])
            if not expected.is_file():
                problems.append(f"{name}: missing {expected.relative_to(REPO_ROOT)}")
            elif expected.read_bytes() != now[output_format]:
                problems.append(f"{name}: {output_format} output differs from {expected.name}")
            if case.get(f"{output_format}_sha256") != _sha256(now[output_format]):
                problems.append(f"{name}: {output_format} hash differs from the manifest")
        if case.get("assets", {}) != now["assets"]:
            problems.append(f"{name}: saved assets differ from the manifest")
    return problems


def accept() -> None:
    manifest, rendered = render_suite()
    for directory in EXPECTED.values():
        directory.mkdir(parents=True, exist_ok=True)
        for stale in directory.iterdir():
            stale.unlink()
    totals: Counter[str] = Counter()
    slides = 0
    for case in manifest["cases"]:
        now = rendered[case["file"]]
        for output_format, directory in EXPECTED.items():
            target = directory / (Path(case["file"]).stem + SUFFIX[output_format])
            target.write_bytes(now[output_format])
            case[f"{output_format}_sha256"] = _sha256(now[output_format])
        case["assets"] = now["assets"]
        totals.update(now["kinds"])
        slides += now["slides"]
    manifest["expected_summary"] = {
        "decks": len(manifest["cases"]),
        "slides": slides,
        "blocks": dict(sorted(totals.items())),
    }
    MANIFEST.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--check", action="store_true")
    mode.add_argument("--accept", action="store_true")
    args = parser.parse_args(argv)
    if args.accept:
        accept()
        print(f"accepted snapshots in {GOLDEN.relative_to(REPO_ROOT)}/expected")
        return 0
    problems = check()
    for problem in problems:
        print(problem)
    if problems:
        print("review the changes, then run: python scripts/update_goldens.py --accept")
        return 1
    print("golden snapshots are current")
    return 0


if __name__ == "__main__":
    sys.exit(main())
