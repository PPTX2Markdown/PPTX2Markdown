"""Compare the legacy converter (src/) with the rewrite (rewrite/src/) on the same inputs.

Both engines ship a top-level ``pptx2markdown`` package, so each one runs in its own
subprocess with its own PYTHONPATH and the harness only reads their JSON output.
The legacy output is treated as the reference.

Usage:
    .venv/bin/python scripts/parity.py                      # repository fixtures
    .venv/bin/python scripts/parity.py pptx_samples deck.pptx
"""

from __future__ import annotations

import argparse
import difflib
import hashlib
import json
import os
import re
import shlex
import shutil
import subprocess
import sys
import time
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Sequence

REPO_ROOT = Path(__file__).resolve().parents[1]
ENGINES = ("legacy", "rewrite")
ENGINE_SOURCES = {
    "legacy": REPO_ROOT / "src",
    "rewrite": REPO_ROOT / "rewrite" / "src",
}
DEFAULT_INPUT_DIRS = (
    REPO_ROOT / "tests" / "fixtures" / "golden",
    REPO_ROOT / "tests" / "fixtures",
)
DEFAULT_REPORT_DIR = REPO_ROOT / ".pptx2markdown" / "parity"
NO_SHAPE_KEY = "<no-shape-id>"
TOKEN_PREVIEW = 12
SLIDE_DETAIL_LIMIT = 15

REWRITE_TEXT_FIELDS = {
    "text": "text",
    "heading": "text",
    "math": "latex",
    "chart": "markdown",
    "smartart": "markdown",
    "table": "markdown",
    "attachment": "label",
    "unsupported": "message",
}

_IMAGE = re.compile(r"!\[[^\]]*\]\([^)]*\)")
_LINK = re.compile(r"\[([^\]]*)\]\([^)]*\)")
_TABLE_RULE = re.compile(r"^[\s|:-]*-{3,}[\s|:-]*$", re.MULTILINE)
_WORD = re.compile(r"\w+")
_SPACES = re.compile(r"[ \t]+")
_EMPHASIS = re.compile(r"(?<!\\)\*+")
_ESCAPE = re.compile(r"\\([!-/:-@\[-`{-~])")


# ---------------------------------------------------------------------------
# Normalization: both engines -> {"slides": [{page, hidden, notes, blocks}]}
# where each block is {kind, text, key, level}.
# ---------------------------------------------------------------------------


def tokens(text: str | None) -> list[str]:
    """Content words with Markdown syntax, link targets and image paths removed."""
    if not text:
        return []
    text = _IMAGE.sub(" ", text)
    text = _LINK.sub(r"\1", text)
    text = _TABLE_RULE.sub(" ", text)
    # Emphasis markers and backslash escapes are formatting, not content: "**도구**의"
    # is the same word as "도구의".
    text = _EMPHASIS.sub("", text)
    text = _ESCAPE.sub(r"\1", text)
    return _WORD.findall(text)


def _shape_key(block: dict[str, Any]) -> str | None:
    shape_id = block.get("shape_id")
    if not shape_id:
        return None
    return f"{block.get('source_part') or 'slide'}:{shape_id}"


def normalize_legacy(data: dict[str, Any]) -> dict[str, Any]:
    slides = []
    for slide in data.get("slides", []):
        blocks = [
            {
                "kind": block["kind"],
                "text": "" if block["kind"] == "image" else block.get("content", ""),
                "key": _shape_key(block),
                "level": block.get("heading_level"),
            }
            for block in slide.get("blocks", [])
        ]
        slides.append(
            {
                "page": int(slide["page"]),
                "hidden": bool(slide.get("hidden")),
                "notes": slide.get("notes"),
                "blocks": blocks,
            }
        )
    return {"slides": slides}


def _rewrite_block_text(block: dict[str, Any]) -> str:
    kind = block["kind"]
    if kind == "image":
        return ""
    if kind == "list":
        return "\n".join(
            f"{'  ' * int(item.get('level', 0))}- {item.get('text', '')}"
            for item in block.get("items", [])
        )
    if kind == "attachment":
        # Same words as the legacy "[attachment: name](path)" link label.
        return f"attachment: {block.get('label', '')}"
    return str(block.get(REWRITE_TEXT_FIELDS.get(kind, "text"), ""))


def normalize_rewrite(data: dict[str, Any]) -> dict[str, Any]:
    slides = []
    for slide in data.get("slides", []):
        blocks = [
            {
                "kind": block["kind"],
                "text": _rewrite_block_text(block),
                "key": _shape_key(block),
                "level": block.get("level") if block["kind"] == "heading" else None,
            }
            for block in slide.get("blocks", [])
        ]
        slides.append(
            {
                "page": int(slide["page"]),
                "hidden": bool(slide.get("hidden")),
                "notes": slide.get("notes"),
                "blocks": blocks,
            }
        )
    return {"slides": slides}


def display_lines(text: str) -> list[str]:
    """Readable, diff-friendly lines: padding and table rules are canonicalized."""
    text = _IMAGE.sub("<image>", text)
    lines = []
    for line in text.splitlines():
        line = _SPACES.sub(" ", line).strip()
        if not line:
            continue
        lines.append("|---|" if _TABLE_RULE.fullmatch(line) else line)
    return lines


def render_canonical(document: dict[str, Any]) -> str:
    """One line per block header so `diff -u legacy.txt rewrite.txt` is readable."""
    lines: list[str] = []
    for slide in document["slides"]:
        hidden = " (hidden)" if slide["hidden"] else ""
        lines.append(f"=== Page {slide['page']}{hidden} ===")
        for block in slide["blocks"]:
            label = block["kind"] + (f" h{block['level']}" if block["level"] else "")
            lines.append(f"[{label}] {block['key'] or '-'}")
            lines.extend(f"    {line}" for line in display_lines(block["text"]))
        if slide["notes"]:
            lines.append("[notes]")
            lines.extend(f"    {line}" for line in display_lines(slide["notes"]))
    return "\n".join(lines) + "\n"


# ---------------------------------------------------------------------------
# Comparison
# ---------------------------------------------------------------------------


def _group_by_shape(blocks: Sequence[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    groups: dict[str, list[dict[str, Any]]] = {}
    for block in blocks:
        groups.setdefault(block["key"] or NO_SHAPE_KEY, []).append(block)
    return groups


def _group_signature(blocks: Sequence[dict[str, Any]]) -> tuple[list[str], list[str]]:
    kinds = [b["kind"] + (f"/h{b['level']}" if b["level"] else "") for b in blocks]
    words = [word for b in blocks for word in tokens(b["text"])]
    return kinds, words


def _preview(text: str) -> str:
    flat = " / ".join(display_lines(text))
    return flat if len(flat) <= 80 else flat[:77] + "..."


def _group_summary(blocks: Sequence[dict[str, Any]]) -> dict[str, Any]:
    kinds, _ = _group_signature(blocks)
    return {"kinds": kinds, "text": _preview("\n".join(b["text"] for b in blocks))}


def _counter_preview(counter: Counter[str]) -> list[str]:
    return list(counter.elements())[:TOKEN_PREVIEW]


def compare_shapes(
    legacy_blocks: Sequence[dict[str, Any]], rewrite_blocks: Sequence[dict[str, Any]]
) -> list[dict[str, Any]]:
    """Per-shape comparison. Returns one row per shape in either engine."""
    legacy_groups = _group_by_shape(legacy_blocks)
    rewrite_groups = _group_by_shape(rewrite_blocks)
    keys = list(legacy_groups) + [k for k in rewrite_groups if k not in legacy_groups]
    rows = []
    for key in keys:
        ref = legacy_groups.get(key)
        cand = rewrite_groups.get(key)
        if cand is None:
            rows.append({"key": key, "status": "missing", "legacy": _group_summary(ref)})
            continue
        if ref is None:
            rows.append({"key": key, "status": "extra", "rewrite": _group_summary(cand)})
            continue
        ref_kinds, ref_words = _group_signature(ref)
        cand_kinds, cand_words = _group_signature(cand)
        row: dict[str, Any] = {"key": key, "legacy_kinds": ref_kinds}
        if ref_kinds == cand_kinds and ref_words == cand_words:
            row["status"] = "same"
        else:
            ref_counter, cand_counter = Counter(ref_words), Counter(cand_words)
            row.update(
                status="differs",
                rewrite_kinds=cand_kinds,
                missing_tokens=_counter_preview(ref_counter - cand_counter),
                extra_tokens=_counter_preview(cand_counter - ref_counter),
                legacy=_group_summary(ref),
                rewrite=_group_summary(cand),
            )
        rows.append(row)
    return rows


def _first_seen_keys(blocks: Sequence[dict[str, Any]]) -> list[str]:
    seen: dict[str, None] = {}
    for block in blocks:
        if block["key"]:
            seen.setdefault(block["key"], None)
    return list(seen)


def order_score(
    legacy_blocks: Sequence[dict[str, Any]], rewrite_blocks: Sequence[dict[str, Any]]
) -> float:
    """Agreement of the shape order over shapes both engines emitted (1.0 = identical)."""
    ref_keys = _first_seen_keys(legacy_blocks)
    cand_keys = _first_seen_keys(rewrite_blocks)
    shared = set(ref_keys) & set(cand_keys)
    ref_seq = [k for k in ref_keys if k in shared]
    cand_seq = [k for k in cand_keys if k in shared]
    if len(ref_seq) < 2:
        return 1.0
    return difflib.SequenceMatcher(None, ref_seq, cand_seq, autojunk=False).ratio()


def _headings(blocks: Sequence[dict[str, Any]]) -> list[list[Any]]:
    return [
        [block["level"], " ".join(tokens(block["text"]))]
        for block in blocks
        if block["kind"] == "heading"
    ]


def _ratio(numerator: int, denominator: int) -> float:
    return 1.0 if denominator == 0 else numerator / denominator


def compare_slide(legacy: dict[str, Any] | None, rewrite: dict[str, Any] | None) -> dict[str, Any]:
    page = (legacy or rewrite or {}).get("page")
    legacy_blocks = legacy["blocks"] if legacy else []
    rewrite_blocks = rewrite["blocks"] if rewrite else []
    ref_words = Counter(w for b in legacy_blocks for w in tokens(b["text"]))
    cand_words = Counter(w for b in rewrite_blocks for w in tokens(b["text"]))
    common = sum((ref_words & cand_words).values())
    ref_total = sum(ref_words.values())
    cand_total = sum(cand_words.values())

    result: dict[str, Any] = {
        "page": page,
        "tokens": {"legacy": ref_total, "rewrite": cand_total, "common": common},
        "recall": _ratio(common, ref_total),
        "precision": _ratio(common, cand_total),
        "kinds": {
            "legacy": dict(Counter(b["kind"] for b in legacy_blocks)),
            "rewrite": dict(Counter(b["kind"] for b in rewrite_blocks)),
        },
    }
    if legacy is None or rewrite is None:
        result["status"] = "missing_in_rewrite" if rewrite is None else "missing_in_legacy"
        result["passed"] = False
        result["shapes"] = compare_shapes(legacy_blocks, rewrite_blocks)
        return result

    result["order"] = order_score(legacy_blocks, rewrite_blocks)
    headings = {"legacy": _headings(legacy_blocks), "rewrite": _headings(rewrite_blocks)}
    result["headings_equal"] = headings["legacy"] == headings["rewrite"]
    if not result["headings_equal"]:
        result["headings"] = headings
    result["notes_equal"] = tokens(legacy["notes"]) == tokens(rewrite["notes"])
    result["hidden_equal"] = legacy["hidden"] == rewrite["hidden"]
    if ref_words != cand_words:
        result["missing_tokens"] = _counter_preview(ref_words - cand_words)
        result["extra_tokens"] = _counter_preview(cand_words - ref_words)
    result["shapes"] = compare_shapes(legacy_blocks, rewrite_blocks)
    result["status"] = "compared"
    result["passed"] = (
        ref_words == cand_words
        and result["order"] == 1.0
        and result["kinds"]["legacy"] == result["kinds"]["rewrite"]
        and result["headings_equal"]
        and result["notes_equal"]
        and result["hidden_equal"]
        and all(row["status"] == "same" for row in result["shapes"])
    )
    return result


def compare_documents(legacy: dict[str, Any], rewrite: dict[str, Any]) -> dict[str, Any]:
    ref = {s["page"]: s for s in legacy["slides"]}
    cand = {s["page"]: s for s in rewrite["slides"]}
    slides = [compare_slide(ref.get(p), cand.get(p)) for p in sorted(ref.keys() | cand.keys())]

    common = sum(s["tokens"]["common"] for s in slides)
    ref_total = sum(s["tokens"]["legacy"] for s in slides)
    cand_total = sum(s["tokens"]["rewrite"] for s in slides)
    kinds = {
        engine: dict(sum((Counter(s["kinds"][engine]) for s in slides), Counter()))
        for engine in ENGINES
    }
    shape_parity: dict[str, dict[str, int]] = {}
    for slide in slides:
        for row in slide["shapes"]:
            if row["status"] == "extra":
                continue
            ref_kinds = row["legacy"]["kinds"] if "legacy" in row else row["legacy_kinds"]
            for kind in {k.split("/")[0] for k in ref_kinds}:
                stats = shape_parity.setdefault(kind, {"shapes": 0, "same": 0})
                stats["shapes"] += 1
                stats["same"] += int(row["status"] == "same")
    return {
        "slides": slides,
        "summary": {
            "slides": {"legacy": len(ref), "rewrite": len(cand)},
            "slides_passed": sum(1 for s in slides if s["passed"]),
            "slides_compared": len(slides),
            "recall": _ratio(common, ref_total),
            "precision": _ratio(common, cand_total),
            "kinds": kinds,
            "shape_parity": shape_parity,
        },
    }


# ---------------------------------------------------------------------------
# Engine execution
# ---------------------------------------------------------------------------


@dataclass
class EngineRun:
    engine: str
    status: str
    seconds: float
    document: dict[str, Any] | None = None
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def engine_command(
    engine: str, python: str, source: Path, out_dir: Path, work_dir: Path, extra: Sequence[str]
) -> list[str]:
    if engine == "legacy":
        return [
            python,
            "-m",
            "pptx2markdown.main_converter.run_pptx_to_markdown",
            str(source),
            "--output-dir",
            str(out_dir),
            "--work-dir",
            str(work_dir),
            "--output-format",
            "json",
            *extra,
        ]
    return [
        python,
        "-m",
        "pptx2markdown",
        str(source),
        "--output-dir",
        str(out_dir),
        "--output-format",
        "json",
        "--tolerant",
        *extra,
    ]


def _stderr_tail(proc: subprocess.CompletedProcess[str], limit: int = 5) -> str:
    lines = [line for line in (proc.stderr or proc.stdout or "").splitlines() if line.strip()]
    return "\n".join(lines[-limit:]) or f"exit code {proc.returncode}"


def _read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _pick_by_source(candidates: list[tuple[Path, dict[str, Any]]], source: Path) -> Path | None:
    for path, data in candidates:
        if (data.get("source") or {}).get("name") == source.name:
            return path
    return candidates[0][0] if len(candidates) == 1 else None


def load_legacy_output(
    out_dir: Path, source: Path, proc: subprocess.CompletedProcess[str], seconds: float
) -> EngineRun:
    errors: list[str] = []
    manifest_path = out_dir / "convert_manifest.json"
    if manifest_path.is_file():
        for package in _read_json(manifest_path).get("packages", []):
            if package.get("status") == "failed":
                errors.append(f"{package.get('stage', 'input')}: {package.get('error')}")
            for slide in package.get("slides", []):
                if slide.get("status") == "failed":
                    errors.append(f"slide {slide.get('page')}: {slide.get('error')}")
    documents = []
    for path in sorted(out_dir.glob("*/*.json")):
        data = _read_json(path)
        if isinstance(data, dict) and "slides" in data:
            documents.append((path, data))
    chosen = _pick_by_source(documents, source)
    if chosen is None:
        return EngineRun("legacy", "failed", seconds, errors=errors or [_stderr_tail(proc)])
    status = "partial" if errors else "ok"
    return EngineRun("legacy", status, seconds, normalize_legacy(_read_json(chosen)), errors)


def load_rewrite_output(
    out_dir: Path, source: Path, proc: subprocess.CompletedProcess[str], seconds: float
) -> EngineRun:
    errors = [
        line.removeprefix("ERROR: ")
        for line in (proc.stderr or "").splitlines()
        if line.startswith("ERROR:")
    ]
    manifests = [
        (path, _read_json(path)) for path in sorted(out_dir.glob("*/bundle-manifest.json"))
    ]
    chosen = _pick_by_source(manifests, source)
    if chosen is None:
        return EngineRun("rewrite", "failed", seconds, errors=errors or [_stderr_tail(proc)])
    data = _read_json(chosen.parent / _read_json(chosen)["document"])
    warnings = [
        f"slide {w.get('slide')}: {w.get('code')}: {w.get('message')}"
        for w in data.get("warnings", [])
    ]
    status = "partial" if errors else "ok"
    return EngineRun("rewrite", status, seconds, normalize_rewrite(data), errors, warnings)


def run_engine(
    engine: str,
    source: Path,
    run_dir: Path,
    python: str,
    extra: Sequence[str],
    timeout: float,
) -> EngineRun:
    engine_dir = run_dir / engine
    shutil.rmtree(engine_dir, ignore_errors=True)
    out_dir = engine_dir / "output"
    work_dir = engine_dir / "work"
    work_dir.mkdir(parents=True)
    env = os.environ.copy()
    env["PYTHONPATH"] = os.pathsep.join(
        p for p in (str(ENGINE_SOURCES[engine]), env.get("PYTHONPATH")) if p
    )
    env.setdefault("PYTHONUTF8", "1")
    command = engine_command(engine, python, source, out_dir, work_dir, extra)
    started = time.perf_counter()
    try:
        proc = subprocess.run(
            command,
            cwd=work_dir,
            env=env,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        elapsed = time.perf_counter() - started
        return EngineRun(engine, "failed", elapsed, errors=[f"timed out after {timeout:g}s"])
    elapsed = time.perf_counter() - started
    (engine_dir / "log.txt").write_text(
        f"$ {shlex.join(command)}\n\n{proc.stdout}\n{proc.stderr}", encoding="utf-8"
    )
    loader = load_legacy_output if engine == "legacy" else load_rewrite_output
    try:
        return loader(out_dir, source, proc, elapsed)
    except (OSError, ValueError, KeyError) as exc:
        return EngineRun(engine, "failed", elapsed, errors=[f"unreadable output: {exc}"])


# ---------------------------------------------------------------------------
# Reporting
# ---------------------------------------------------------------------------


def _display_path(path: Path) -> str:
    try:
        return path.relative_to(REPO_ROOT).as_posix()
    except ValueError:
        return str(path)


def _engine_cell(run: EngineRun) -> str:
    if run.status == "partial":
        return f"partial ({len(run.errors)} err)"
    return run.status


def _pct(value: float | None) -> str:
    return "-" if value is None else f"{value * 100:.1f}%"


def _kind_delta(kinds: dict[str, dict[str, int]]) -> str:
    names = sorted(kinds["legacy"].keys() | kinds["rewrite"].keys())
    parts = [
        f"{name} {kinds['legacy'].get(name, 0)}→{kinds['rewrite'].get(name, 0)}"
        for name in names
        if kinds["legacy"].get(name, 0) != kinds["rewrite"].get(name, 0)
    ]
    return ", ".join(parts)


def render_summary(report: dict[str, Any]) -> str:
    totals = report["totals"]
    lines = [
        "## Parity summary (reference: legacy)",
        "",
        f"- Decks: {totals['decks']} (compared: {totals['decks_compared']})",
        f"- Slides passing: {totals['slides_passed']}/{totals['slides_compared']}",
        f"- Token recall: {_pct(totals['recall'])}, precision: {_pct(totals['precision'])}",
        "",
        "| Deck | Slides L/R | Legacy | Rewrite | Recall | Precision | Passing | Time L/R |",
        "| --- | --- | --- | --- | --- | --- | --- | --- |",
    ]
    for deck in report["decks"]:
        legacy, rewrite = deck["engines"]["legacy"], deck["engines"]["rewrite"]
        summary = (deck.get("comparison") or {}).get("summary")
        if summary:
            slides = f"{summary['slides']['legacy']}/{summary['slides']['rewrite']}"
            recall, precision = _pct(summary["recall"]), _pct(summary["precision"])
            passing = f"{summary['slides_passed']}/{summary['slides_compared']}"
        else:
            slides, recall, precision, passing = "-", "-", "-", "-"
        lines.append(
            f"| {deck['name']} | {slides} | {_engine_cell(EngineRun(**legacy))} "
            f"| {_engine_cell(EngineRun(**rewrite))} | {recall} | {precision} | {passing} "
            f"| {legacy['seconds']:.1f}s/{rewrite['seconds']:.1f}s |"
        )
    lines += [
        "",
        "| Kind | Legacy blocks | Rewrite blocks | Identical shapes |",
        "| --- | --- | --- | --- |",
    ]
    kinds = totals["kinds"]
    for name in sorted(kinds["legacy"].keys() | kinds["rewrite"].keys()):
        parity = totals["shape_parity"].get(name)
        same = f"{parity['same']}/{parity['shapes']}" if parity else "-"
        lines.append(
            f"| {name} | {kinds['legacy'].get(name, 0)} | {kinds['rewrite'].get(name, 0)} "
            f"| {same} |"
        )
    return "\n".join(lines) + "\n"


def _slide_detail(slide: dict[str, Any]) -> list[str]:
    head = f"- **Page {slide['page']}**"
    if slide["status"] != "compared":
        return [f"{head}: {slide['status'].replace('_', ' ')}"]
    facts = [f"recall {_pct(slide['recall'])}", f"precision {_pct(slide['precision'])}"]
    if slide["order"] < 1.0:
        facts.append(f"order {slide['order']:.2f}")
    lines = [f"{head}: " + ", ".join(facts)]
    delta = _kind_delta(slide["kinds"])
    if delta:
        lines.append(f"  - kinds: {delta}")
    if not slide["headings_equal"]:
        lines.append(f"  - headings legacy: `{slide['headings']['legacy']}`")
        lines.append(f"  - headings rewrite: `{slide['headings']['rewrite']}`")
    if not slide["notes_equal"]:
        lines.append("  - speaker notes differ")
    if not slide["hidden_equal"]:
        lines.append("  - hidden flag differs")
    if slide.get("missing_tokens"):
        lines.append(f"  - missing in rewrite: `{' '.join(slide['missing_tokens'])}`")
    if slide.get("extra_tokens"):
        lines.append(f"  - extra in rewrite: `{' '.join(slide['extra_tokens'])}`")
    for row in slide["shapes"]:
        if row["status"] == "same":
            continue
        if row["status"] == "missing":
            lines.append(
                f"  - shape `{row['key']}` missing: {row['legacy']['kinds']} "
                f"“{row['legacy']['text']}”"
            )
        elif row["status"] == "extra":
            lines.append(
                f"  - shape `{row['key']}` extra: {row['rewrite']['kinds']} "
                f"“{row['rewrite']['text']}”"
            )
        else:
            lines.append(
                f"  - shape `{row['key']}`: {row['legacy']['kinds']} → "
                f"{row['rewrite']['kinds']}; “{row['legacy']['text']}” → "
                f"“{row['rewrite']['text']}”"
            )
    return lines


def render_details(report: dict[str, Any], report_dir: Path) -> str:
    lines = ["## Deck details", ""]
    for deck in report["decks"]:
        lines.append(f"### {deck['name']}")
        lines.append("")
        lines.append(f"Source: `{deck['source']}`")
        for engine in ENGINES:
            run = deck["engines"][engine]
            for error in run["errors"]:
                lines.append(f"- {engine} error: {error}")
            for warning in run["warnings"][:5]:
                lines.append(f"- {engine} warning: {warning}")
            if len(run["warnings"]) > 5:
                lines.append(f"- {engine}: {len(run['warnings']) - 5} more warnings")
        comparison = deck.get("comparison")
        if comparison:
            run_dir = report_dir / "runs" / deck["id"]
            lines.append(
                f"- Diff: `diff -u '{run_dir / 'legacy.txt'}' '{run_dir / 'rewrite.txt'}'`"
            )
            failing = [s for s in comparison["slides"] if not s["passed"]]
            for slide in failing[:SLIDE_DETAIL_LIMIT]:
                lines.extend(_slide_detail(slide))
            if len(failing) > SLIDE_DETAIL_LIMIT:
                lines.append(f"- … {len(failing) - SLIDE_DETAIL_LIMIT} more failing slides")
            if not failing:
                lines.append("- All slides identical.")
        lines.append("")
    return "\n".join(lines)


def build_report(decks: list[dict[str, Any]]) -> dict[str, Any]:
    compared = [d["comparison"]["summary"] for d in decks if d.get("comparison")]
    kinds: dict[str, Counter[str]] = {engine: Counter() for engine in ENGINES}
    shape_parity: dict[str, dict[str, int]] = {}
    common = ref_total = cand_total = 0
    for deck in decks:
        for slide in (deck.get("comparison") or {}).get("slides", []):
            common += slide["tokens"]["common"]
            ref_total += slide["tokens"]["legacy"]
            cand_total += slide["tokens"]["rewrite"]
    for summary in compared:
        for engine in ENGINES:
            kinds[engine].update(summary["kinds"][engine])
        for name, stats in summary["shape_parity"].items():
            total = shape_parity.setdefault(name, {"shapes": 0, "same": 0})
            total["shapes"] += stats["shapes"]
            total["same"] += stats["same"]
    return {
        "reference": "legacy",
        "totals": {
            "decks": len(decks),
            "decks_compared": len(compared),
            "slides_passed": sum(s["slides_passed"] for s in compared),
            "slides_compared": sum(s["slides_compared"] for s in compared),
            "recall": _ratio(common, ref_total),
            "precision": _ratio(common, cand_total),
            "kinds": {engine: dict(counter) for engine, counter in kinds.items()},
            "shape_parity": shape_parity,
        },
        "decks": decks,
    }


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def collect_inputs(raw_inputs: Sequence[str]) -> list[Path]:
    candidates: list[Path] = []
    roots = [Path(raw) for raw in raw_inputs] or list(DEFAULT_INPUT_DIRS)
    for root in roots:
        if root.is_dir():
            candidates.extend(
                path
                for path in sorted(root.glob("*.pptx"))
                if path.is_file() and not path.name.startswith("~$")
            )
        elif root.is_file() and root.suffix.lower() in {".pptx", ".ppt"}:
            candidates.append(root)
        else:
            raise SystemExit(f"not a .pptx/.ppt file or directory: {root}")
    unique: dict[Path, None] = {}
    for path in candidates:
        unique.setdefault(path.resolve(), None)
    return list(unique)


def deck_id(source: Path) -> str:
    stem = re.sub(r"[^\w.-]+", "_", source.stem).strip("._") or "deck"
    digest = hashlib.sha1(str(source).encode("utf-8")).hexdigest()[:6]
    return f"{stem}-{digest}"


def parse_args(argv: Sequence[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Compare legacy and rewrite conversions (legacy is the reference)."
    )
    parser.add_argument(
        "inputs",
        nargs="*",
        help="PPTX files or directories (default: tests/fixtures and tests/fixtures/golden)",
    )
    parser.add_argument(
        "--report-dir",
        default=str(DEFAULT_REPORT_DIR),
        help="Where report.md, report.json and per-deck runs go (default: %(default)s)",
    )
    parser.add_argument("--legacy-python", default=sys.executable)
    parser.add_argument("--rewrite-python", default=sys.executable)
    parser.add_argument(
        "--legacy-args",
        default="",
        help="Extra legacy CLI arguments, e.g. '--placeholder-inheritance none'",
    )
    parser.add_argument("--rewrite-args", default="", help="Extra rewrite CLI arguments")
    parser.add_argument("--jobs", type=int, default=min(8, os.cpu_count() or 1))
    parser.add_argument("--timeout", type=float, default=300.0, help="Seconds per engine run")
    parser.add_argument(
        "--baseline",
        help="JSON baseline; exit 1 when the rewrite does worse on any of its numbers",
    )
    parser.add_argument(
        "--update-baseline",
        action="store_true",
        help="Write the current numbers to --baseline instead of checking them",
    )
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    for engine, src in ENGINE_SOURCES.items():
        if not (src / "pptx2markdown").is_dir():
            print(f"{engine} sources not found: {src}", file=sys.stderr)
            return 2
    sources = collect_inputs(args.inputs)
    if not sources:
        print("no .pptx inputs found", file=sys.stderr)
        return 2

    report_dir = Path(args.report_dir).expanduser().resolve()
    pythons = {"legacy": args.legacy_python, "rewrite": args.rewrite_python}
    extras = {"legacy": shlex.split(args.legacy_args), "rewrite": shlex.split(args.rewrite_args)}
    ids = {source: deck_id(source) for source in sources}

    with ThreadPoolExecutor(max_workers=max(1, args.jobs)) as pool:
        futures = {
            (source, engine): pool.submit(
                run_engine,
                engine,
                source,
                report_dir / "runs" / ids[source],
                pythons[engine],
                extras[engine],
                args.timeout,
            )
            for source in sources
            for engine in ENGINES
        }
        runs = {key: future.result() for key, future in futures.items()}

    decks = []
    for source in sources:
        run_dir = report_dir / "runs" / ids[source]
        engine_runs = {engine: runs[(source, engine)] for engine in ENGINES}
        for engine, run in engine_runs.items():
            if run.document is not None:
                (run_dir / f"{engine}.txt").write_text(
                    render_canonical(run.document), encoding="utf-8"
                )
        legacy_doc = engine_runs["legacy"].document
        rewrite_doc = engine_runs["rewrite"].document
        decks.append(
            {
                "id": ids[source],
                "name": source.name,
                "source": _display_path(source),
                "engines": {
                    engine: {k: v for k, v in asdict(run).items() if k != "document"}
                    for engine, run in engine_runs.items()
                },
                "comparison": (
                    compare_documents(legacy_doc, rewrite_doc)
                    if legacy_doc is not None and rewrite_doc is not None
                    else None
                ),
            }
        )

    report = build_report(decks)
    summary = render_summary(report)
    report_dir.mkdir(parents=True, exist_ok=True)
    (report_dir / "report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    (report_dir / "report.md").write_text(
        summary + "\n" + render_details(report, report_dir), encoding="utf-8"
    )
    print(summary)
    print(f"Details: {report_dir / 'report.md'}")

    if args.baseline:
        baseline_path = Path(args.baseline)
        current = baseline_metrics(report)
        if args.update_baseline:
            baseline_path.write_text(
                json.dumps(current, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
                encoding="utf-8",
            )
            print(f"Baseline written: {baseline_path}")
            return 0
        regressions = compare_baseline(
            json.loads(baseline_path.read_text(encoding="utf-8")), current
        )
        if regressions:
            print("\nParity regressed against the baseline:", file=sys.stderr)
            for line in regressions:
                print(f"- {line}", file=sys.stderr)
            return 1
        print("\nNo parity regression against the baseline.")
    return 0


def baseline_metrics(report: dict[str, Any]) -> dict[str, Any]:
    """The numbers a change must not make worse."""
    totals = report["totals"]
    return {
        "slides_passed": totals["slides_passed"],
        "slides_compared": totals["slides_compared"],
        "recall": round(totals["recall"], 4),
        "precision": round(totals["precision"], 4),
        "same_shapes": {kind: value["same"] for kind, value in totals["shape_parity"].items()},
        "failed_runs": sorted(
            f"{deck['name']}:{engine}"
            for deck in report["decks"]
            for engine, run in deck["engines"].items()
            if run["status"] != "ok"
        ),
    }


def compare_baseline(baseline: dict[str, Any], current: dict[str, Any]) -> list[str]:
    """Regressions only; improvements pass and can be recorded with --update-baseline."""
    regressions = []
    for key in ("slides_passed", "recall", "precision"):
        if current[key] < baseline[key]:
            regressions.append(f"{key}: {baseline[key]} -> {current[key]}")
    for kind, same in baseline["same_shapes"].items():
        now = current["same_shapes"].get(kind, 0)
        if now < same:
            regressions.append(f"identical {kind} shapes: {same} -> {now}")
    new_failures = sorted(set(current["failed_runs"]) - set(baseline["failed_runs"]))
    if new_failures:
        regressions.append("runs that now fail: " + ", ".join(new_failures))
    return regressions


if __name__ == "__main__":
    raise SystemExit(main())
