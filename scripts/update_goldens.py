from __future__ import annotations

import argparse
import difflib
import hashlib
import json
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any

import pptx2markdown

REPO_ROOT = Path(__file__).resolve().parents[1]
FIXTURE_DIR = REPO_ROOT / "tests" / "fixtures" / "golden"
MANIFEST_PATH = FIXTURE_DIR / "manifest.json"
DIFF_LINE_LIMIT = 200


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def load_manifest() -> dict[str, Any]:
    return json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))


def result_path(output_dir: Path, stem: str, output_format: str) -> Path:
    extension = "json" if output_format == "json" else "md"
    return output_dir / stem / f"{stem}.{extension}"


def asset_hashes(package_dir: Path, result: Path) -> dict[str, str]:
    return {
        path.relative_to(package_dir).as_posix(): sha256(path)
        for path in sorted(package_dir.rglob("*"))
        if path.is_file() and path != result
    }


def print_snapshot_diff(snapshot: Path, actual: Path) -> None:
    expected_text = snapshot.read_text(encoding="utf-8") if snapshot.exists() else ""
    actual_text = actual.read_text(encoding="utf-8")
    diff = list(
        difflib.unified_diff(
            expected_text.splitlines(),
            actual_text.splitlines(),
            fromfile=str(snapshot.relative_to(REPO_ROOT)),
            tofile=f"generated/{actual.name}",
            lineterm="",
        )
    )
    for line in diff[:DIFF_LINE_LIMIT]:
        print(line)
    if len(diff) > DIFF_LINE_LIMIT:
        print(f"... {len(diff) - DIFF_LINE_LIMIT} additional diff lines omitted")


def convert_all(pptx_paths: list[Path], temp_root: Path) -> tuple[dict[str, Path], dict[str, Any]]:
    outputs: dict[str, Path] = {}
    summaries: dict[str, Any] = {}
    for output_format in ("json", "markdown"):
        output_dir = temp_root / f"{output_format}-output"
        exit_code = pptx2markdown.convert(
            pptx_paths,
            output_dir=output_dir,
            work_dir=temp_root / f"{output_format}-work",
            output_format=output_format,
        )
        if exit_code != 0:
            raise RuntimeError(f"{output_format} conversion failed with exit code {exit_code}")
        conversion_manifest = json.loads(
            (output_dir / "convert_manifest.json").read_text(encoding="utf-8")
        )
        outputs[output_format] = output_dir
        summaries[output_format] = conversion_manifest["summary"]
    if summaries["json"] != summaries["markdown"]:
        raise RuntimeError("JSON and Markdown conversion summaries differ")
    return outputs, summaries["json"]


def verify_input_hashes(cases: list[dict[str, Any]]) -> list[Path]:
    paths: list[Path] = []
    for case in cases:
        path = FIXTURE_DIR / case["file"]
        if not path.is_file():
            raise RuntimeError(f"missing golden PPTX: {path.relative_to(REPO_ROOT)}")
        actual = sha256(path)
        if actual != case["sha256"]:
            raise RuntimeError(
                f"golden PPTX changed: {case['file']}\n"
                "Review its provenance and update the input hash manually before accepting "
                "new parser snapshots."
            )
        paths.append(path)
    return paths


def expected_snapshot_paths(manifest: dict[str, Any], stem: str) -> dict[str, Path]:
    return {
        "json": FIXTURE_DIR / manifest["snapshot_roots"]["json"] / f"{stem}.json",
        "markdown": FIXTURE_DIR / manifest["snapshot_roots"]["markdown"] / f"{stem}.md",
    }


def collect_changes(
    manifest: dict[str, Any], outputs: dict[str, Path], summary: dict[str, Any]
) -> list[str]:
    changes: list[str] = []
    if manifest["expected_summary"] != summary:
        changes.append("batch conversion summary changed")

    for case in manifest["cases"]:
        stem = Path(case["file"]).stem
        snapshots = expected_snapshot_paths(manifest, stem)
        actual_assets: dict[str, dict[str, str]] = {}
        for output_format in ("json", "markdown"):
            actual = result_path(outputs[output_format], stem, output_format)
            snapshot = snapshots[output_format]
            if not snapshot.exists() or snapshot.read_bytes() != actual.read_bytes():
                changes.append(f"{case['file']}: {output_format} snapshot changed")
                print_snapshot_diff(snapshot, actual)

            expected_digest = case[f"{output_format}_sha256"]
            if sha256(actual) != expected_digest:
                changes.append(f"{case['file']}: {output_format} hash changed")
            actual_assets[output_format] = asset_hashes(actual.parent, actual)

        if actual_assets["json"] != actual_assets["markdown"]:
            raise RuntimeError(f"asset outputs differ by format: {case['file']}")
        if actual_assets["json"] != case["assets"]:
            changes.append(f"{case['file']}: emitted assets changed")
    return changes


def accept_outputs(
    manifest: dict[str, Any], outputs: dict[str, Path], summary: dict[str, Any]
) -> None:
    manifest["expected_summary"] = summary
    expected_names: dict[str, set[str]] = {"json": set(), "markdown": set()}

    for case in manifest["cases"]:
        stem = Path(case["file"]).stem
        snapshots = expected_snapshot_paths(manifest, stem)
        assets_by_format: dict[str, dict[str, str]] = {}
        for output_format in ("json", "markdown"):
            actual = result_path(outputs[output_format], stem, output_format)
            snapshot = snapshots[output_format]
            snapshot.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(actual, snapshot)
            case[f"{output_format}_sha256"] = sha256(actual)
            expected_names[output_format].add(snapshot.name)
            assets_by_format[output_format] = asset_hashes(actual.parent, actual)
        if assets_by_format["json"] != assets_by_format["markdown"]:
            raise RuntimeError(f"asset outputs differ by format: {case['file']}")
        case["assets"] = dict(sorted(assets_by_format["json"].items()))

    for output_format, relative_dir in manifest["snapshot_roots"].items():
        snapshot_dir = FIXTURE_DIR / relative_dir
        extension = "json" if output_format == "json" else "md"
        for stale in snapshot_dir.glob(f"*.{extension}"):
            if stale.name not in expected_names[output_format]:
                stale.unlink()

    MANIFEST_PATH.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Check or intentionally refresh repository golden parser outputs."
    )
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument(
        "--check",
        action="store_true",
        help="fail if generated JSON, Markdown, hashes, assets, or summary changed",
    )
    mode.add_argument(
        "--accept",
        action="store_true",
        help="replace reviewed snapshots and update output hashes in the manifest",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    manifest = load_manifest()
    cases = manifest["cases"]
    try:
        pptx_paths = verify_input_hashes(cases)
        with tempfile.TemporaryDirectory(prefix="pptx2markdown-goldens-") as temp_dir:
            outputs, summary = convert_all(pptx_paths, Path(temp_dir))
            changes = collect_changes(manifest, outputs, summary)
            if args.accept:
                accept_outputs(manifest, outputs, summary)
                print(f"Accepted {len(cases)} golden JSON/Markdown pairs and their asset hashes.")
                return 0
    except (OSError, RuntimeError, ValueError, json.JSONDecodeError) as exc:
        print(f"golden update failed: {exc}", file=sys.stderr)
        return 2

    if changes:
        print("\nGolden outputs are stale:")
        for change in changes:
            print(f"- {change}")
        print("Review the diff, then run: python scripts/update_goldens.py --accept")
        return 1

    print(f"Golden outputs are current: {len(cases)} PPTX files, two output formats.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
