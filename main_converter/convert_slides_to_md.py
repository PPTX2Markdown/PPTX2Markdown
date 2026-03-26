#!/usr/bin/env python3
"""
Convert extracted PPTX package(s) to markdown.

Usage:
  python convert_slides_to_md.py [package_name|file.pptx ...]

Rules:
  - If no positional args are provided, process all package roots in ./target_slides.
  - Package root example: ./target_slides/sample1
  - If a .pptx file is provided, it is extracted into ./target_slides/<stem>/ first.
  - Each package must contain: ./ppt/slides
  - Output is always written to ./output (created automatically).
  - Input slide XML order is assumed to be the final reading order.
"""

from __future__ import annotations

import argparse
import importlib
import json
import os
import re
import shutil
import subprocess
import sys
import time
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple
import xml.etree.ElementTree as ET

from pydantic import BaseModel, ConfigDict, Field

from heading_rules import (
    HeadingPolicy,
    clean_heading_text_for_render as hr_clean_heading_text_for_render,
    infer_heading_depth_fallback as hr_infer_heading_depth_fallback,
    is_body_like_long_sentence as hr_is_body_like_long_sentence,
    looks_like_multi_numbered_items as hr_looks_like_multi_numbered_items,
    normalize_single_heading_to_h1,
    strict_heading_depth_from_placeholder as hr_strict_heading_depth_from_placeholder,
)


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
REL_NS = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}

_IMAGE_MARKDOWN_PIPELINE_MODULE: Optional[object] = None
_IMAGE_MARKDOWN_PIPELINE_IMPORT_ERROR: Optional[str] = None
_IMAGE_VLM_DEBUG_JSON = os.getenv("IMAGE_VLM_DEBUG_JSON", "").strip().lower() in {"1", "true", "yes", "on"}
DEFAULT_IMAGE_VLM_PROMPT = (
    "Convert this image into concise Markdown for RAG ingestion.\n"
    "- Extract visible text and table content faithfully.\n"
    "- If the image is mainly a table, recreate it as a Markdown table and keep readable cell text.\n"
    "- Preserve visible headings, paragraphs, bullet lists, numbered lists, and code-like text.\n"
    "- If the image is a chart, diagram, infographic, or screenshot, summarize only the useful visible content in Markdown.\n"
    "- If some text is unreadable, omit it instead of guessing.\n"
    "- If the image does not contain useful documentable information, answer exactly: 불필요한 정보\n"
    "- Do not wrap the answer in triple backticks."
)
DEFAULT_IMAGE_VLM_MAX_NEW_TOKENS = 1024


def run_structure_analysis_stage(
    repo_root: Path,
    package_name: str,
    slide_xmls: Sequence[Path],
    strict: bool = False,
) -> Tuple[Dict[str, Path], Path]:
    ro_script = repo_root / "structure_analyzer" / "extract_structure_analysis.py"
    if not ro_script.exists():
        raise FileNotFoundError(f"structure_analyzer script not found: {ro_script}")

    # Keep XML-stage artifacts as persistent package-scoped outputs for debugging/reuse.
    ro_output = repo_root / "structure_analyzer" / "output" / package_name
    if ro_output.exists():
        shutil.rmtree(ro_output)
    ro_output.mkdir(parents=True, exist_ok=True)
    cmd = [
        "python3",
        str(ro_script),
        "--mode",
        "xml",
        "--output-dir",
        str(ro_output),
    ]
    if strict:
        cmd.append("--strict")
    cmd.extend(str(p.resolve()) for p in slide_xmls)

    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        raise RuntimeError(
            "structure_analysis stage failed\n"
            f"cmd: {' '.join(cmd)}\n"
            f"stdout:\n{proc.stdout}\n"
            f"stderr:\n{proc.stderr}"
        )

    manifest_path = ro_output / "structure_analysis_manifest.json"
    if not manifest_path.exists():
        legacy_manifest = ro_output / "reading_order_manifest.json"
        if legacy_manifest.exists():
            manifest_path = legacy_manifest
        else:
            raise FileNotFoundError(f"structure_analysis manifest not found: {manifest_path}")
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    failed = manifest.get("failed", [])
    if isinstance(failed, list) and failed:
        raise RuntimeError(f"structure_analysis stage reported failures: {json.dumps(failed, ensure_ascii=False)}")

    mapping: Dict[str, Path] = {}
    processed = manifest.get("processed", [])
    if isinstance(processed, list):
        for row in processed:
            if not isinstance(row, dict):
                continue
            src = row.get("input_xml")
            out = row.get("output_xml")
            if isinstance(src, str) and isinstance(out, str):
                mapping[str(Path(src).resolve())] = Path(out).resolve()
    return mapping, ro_output


def ensure_imports(repo_root: Path) -> None:
    # Ensure project root is importable for local packages.
    candidates = [
        repo_root,
    ]
    for cand in candidates:
        if cand.exists() and cand.is_dir() and str(cand) not in sys.path:
            sys.path.insert(0, str(cand))


def default_target_dirs(base_dir: Path) -> List[Path]:
    # Always anchor under main_converter/.
    local_target = base_dir / "target_slides"
    if local_target.exists() and not local_target.is_dir():
        raise NotADirectoryError(f"target_slides path exists but is not a directory: {local_target}")
    local_target.mkdir(parents=True, exist_ok=True)
    return [local_target]


def default_pptx_input_dirs(base_dir: Path) -> List[Path]:
    # Always anchor under main_converter/.
    local_target = base_dir / "target_pptx"
    if local_target.exists() and not local_target.is_dir():
        raise NotADirectoryError(f"target_pptx path exists but is not a directory: {local_target}")
    local_target.mkdir(parents=True, exist_ok=True)
    return [local_target]


def natural_key(name: str) -> Tuple:
    parts = re.split(r"(\d+)", name)
    out: List[object] = []
    for part in parts:
        if part.isdigit():
            out.append(int(part))
        else:
            out.append(part.lower())
    return tuple(out)


def local_name(tag: str) -> str:
    return tag.split("}", 1)[-1]


def normalize_text(s: str) -> str:
    s = s.replace("\n", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def pick_packages(target_dirs: Sequence[Path], raw_inputs: Sequence[str]) -> List[Path]:
    pkgs: List[Path] = []
    if raw_inputs:
        for item in raw_inputs:
            p = Path(item)
            candidates = [p]
            for td in target_dirs:
                candidates.append(td / item)
            picked = None
            for c in candidates:
                if c.exists() and c.is_dir():
                    picked = c.resolve()
                    break
            if picked is not None:
                pkgs.append(picked)
    else:
        for target_dir in target_dirs:
            for p in sorted(target_dir.iterdir(), key=lambda x: natural_key(x.name)):
                if p.is_dir():
                    pkgs.append(p.resolve())

    out: List[Path] = []
    uniq: Dict[str, Path] = {}
    for pkg in pkgs:
        slides_dir = pkg / "ppt" / "slides"
        if slides_dir.exists() and slides_dir.is_dir():
            uniq[str(pkg)] = pkg
    for _, v in uniq.items():
        out.append(v)
    return out


def sanitize_package_name(name: str) -> str:
    cleaned = re.sub(r"[^0-9A-Za-z._-]+", "_", name).strip("._")
    return cleaned or "package"


def preferred_target_dir(base_dir: Path) -> Path:
    cands = default_target_dirs(base_dir)
    if cands:
        return cands[0]
    return base_dir / "target_slides"


def preferred_pptx_input_dir(base_dir: Path) -> Path:
    cands = default_pptx_input_dirs(base_dir)
    if cands:
        return cands[0]
    return base_dir / "target_pptx"


def package_marker_path(pkg_dir: Path) -> Path:
    return pkg_dir / ".pptx_source.json"


def package_marker_matches(pkg_dir: Path, pptx_path: Path) -> bool:
    marker = package_marker_path(pkg_dir)
    if not marker.exists():
        return False
    try:
        payload = json.loads(marker.read_text(encoding="utf-8"))
    except Exception:
        return False
    if not isinstance(payload, dict):
        return False
    slides_dir = pkg_dir / "ppt" / "slides"
    if not slides_dir.exists():
        return False
    try:
        stat = pptx_path.stat()
    except OSError:
        return False
    return (
        payload.get("source_path") == str(pptx_path.resolve())
        and payload.get("size") == stat.st_size
        and payload.get("mtime_ns") == stat.st_mtime_ns
    )


def safe_extract_pptx(pptx_path: Path, dest_dir: Path) -> None:
    with zipfile.ZipFile(pptx_path) as zf:
        for member in zf.infolist():
            member_path = Path(member.filename)
            if member_path.is_absolute() or ".." in member_path.parts:
                raise ValueError(f"Unsafe archive entry: {member.filename}")
        zf.extractall(dest_dir)


def extract_pptx_to_target(
    pptx_path: Path,
    extraction_root: Path,
    staged_pptx_root: Path,
    allow_replace_unmanaged: bool = False,
) -> Path:
    stat = pptx_path.stat()
    extraction_root.mkdir(parents=True, exist_ok=True)
    pkg_name = sanitize_package_name(pptx_path.stem)
    pkg_dir = extraction_root / pkg_name

    # Raw extraction target should be a real managed directory, never a symlink.
    # Remove legacy symlink targets (including valid/broken links) before extraction.
    if pkg_dir.is_symlink():
        try:
            pkg_dir.unlink()
        except PermissionError:
            if allow_replace_unmanaged:
                # Some bind-mounted filesystems disallow unlinking symlinks.
                # Fall back to an alternate managed directory name.
                suffix = "__raw"
                idx = 0
                while True:
                    candidate_name = f"{pkg_name}{suffix}" if idx == 0 else f"{pkg_name}{suffix}{idx}"
                    candidate = extraction_root / candidate_name
                    if not candidate.exists() and not candidate.is_symlink():
                        pkg_name = candidate_name
                        pkg_dir = candidate
                        break
                    idx += 1
            else:
                raise

    if package_marker_matches(pkg_dir, pptx_path):
        staged_pptx_root.mkdir(parents=True, exist_ok=True)
        staged_pptx = staged_pptx_root / f"{pkg_name}.pptx"
        if staged_pptx.resolve() != pptx_path.resolve():
            shutil.copy2(pptx_path, staged_pptx)
        return pkg_dir.resolve()

    marker = package_marker_path(pkg_dir)
    if pkg_dir.exists():
        if marker.exists():
            shutil.rmtree(pkg_dir)
        elif allow_replace_unmanaged:
            shutil.rmtree(pkg_dir)
        else:
            raise FileExistsError(
                f"target package directory already exists and is not managed by converter: {pkg_dir}"
            )

    pkg_dir.mkdir(parents=True, exist_ok=True)
    safe_extract_pptx(pptx_path, pkg_dir)

    slides_dir = pkg_dir / "ppt" / "slides"
    if not slides_dir.exists() or not slides_dir.is_dir():
        shutil.rmtree(pkg_dir, ignore_errors=True)
        raise ValueError(f"Not a valid pptx package after extraction: {pptx_path}")

    marker.write_text(
        json.dumps(
            {
                "source_path": str(pptx_path.resolve()),
                "size": stat.st_size,
                "mtime_ns": stat.st_mtime_ns,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    staged_pptx_root.mkdir(parents=True, exist_ok=True)
    staged_pptx = staged_pptx_root / f"{pkg_name}.pptx"
    if staged_pptx.resolve() != pptx_path.resolve():
        shutil.copy2(pptx_path, staged_pptx)

    return pkg_dir.resolve()


def is_ignored_pptx_file(path: Path) -> bool:
    # Skip Office lock/temp files like "~$sample1.pptx".
    return path.name.startswith("~$")


def resolve_input_pptx_path(cwd: Path, item: str) -> Tuple[Optional[Path], List[Path]]:
    search_roots = [cwd]
    for root in default_pptx_input_dirs(cwd):
        if root not in search_roots:
            search_roots.append(root)
    for root in default_target_dirs(cwd):
        if root not in search_roots:
            search_roots.append(root)

    candidates: List[Path] = []
    seen: set[str] = set()

    def add_candidate(path: Path) -> None:
        key = str(path)
        if key not in seen:
            seen.add(key)
            candidates.append(path)

    raw_path = Path(item)
    add_candidate(raw_path)
    for root in search_roots:
        add_candidate(root / item)

    for cand in candidates:
        if cand.exists() and cand.is_file() and cand.suffix.lower() == ".pptx":
            if is_ignored_pptx_file(cand):
                continue
            return cand.resolve(), candidates
        if cand.suffix.lower() != ".pptx":
            cand_pptx = cand.with_suffix(".pptx")
            if cand_pptx.exists() and cand_pptx.is_file() and not is_ignored_pptx_file(cand_pptx):
                return cand_pptx.resolve(), candidates
    return None, candidates


def prepare_package_inputs(
    cwd: Path,
    raw_inputs: Sequence[str],
    force_extract: bool = False,
) -> Tuple[List[str], List[Dict[str, object]]]:
    if not raw_inputs:
        return list(raw_inputs), []

    prepared: List[str] = []
    missing_inputs: List[Dict[str, object]] = []
    extraction_root = preferred_target_dir(cwd)
    staged_pptx_root = preferred_pptx_input_dir(cwd)

    for item in raw_inputs:
        picked_file, candidates = resolve_input_pptx_path(cwd, item)
        if picked_file is None:
            missing_inputs.append(
                {
                    "input": item,
                    "checked": [str(path.resolve()) if path.is_absolute() else str((cwd / path).resolve()) for path in candidates],
                }
            )
            continue
        pkg_dir = extract_pptx_to_target(
            picked_file,
            extraction_root,
            staged_pptx_root=staged_pptx_root,
            allow_replace_unmanaged=force_extract,
        )
        prepared.append(str(pkg_dir))
    return prepared, missing_inputs


def collect_target_pptx_inputs(cwd: Path) -> List[str]:
    files: List[Path] = []
    for root in default_pptx_input_dirs(cwd):
        for path in root.glob("*.pptx"):
            if path.is_file() and not is_ignored_pptx_file(path):
                files.append(path.resolve())
    files = sorted(files, key=lambda p: natural_key(p.name))
    uniq: Dict[str, Path] = {str(p): p for p in files}
    return [str(p) for p in uniq.values()]


def parse_slide_number(filename: str, default_idx: int) -> int:
    m = re.search(r"slide(\d+)", filename, re.IGNORECASE)
    if m:
        return int(m.group(1))
    return default_idx


def slide_base_name(slide_xml: Path) -> str:
    m = re.search(r"(slide\d+)", slide_xml.stem, re.IGNORECASE)
    if m:
        return f"{m.group(1)}.xml"
    return slide_xml.name


def find_sidecar_json(slide_xml: Path) -> Optional[Path]:
    # e.g. slide2.reordered.xml -> slide2.structure_analysis.json
    stem = slide_xml.stem
    candidates = []
    if stem.endswith(".reordered"):
        candidates.append(slide_xml.with_name(f"{stem[:-len('.reordered')]}.structure_analysis.json"))
        candidates.append(slide_xml.with_name(f"{stem[:-len('.reordered')]}.reading_order.json"))
    candidates.append(slide_xml.with_name(f"{stem}.structure_analysis.json"))
    candidates.append(slide_xml.with_name(f"{stem}.reading_order.json"))
    for c in candidates:
        if c.exists() and c.is_file():
            return c
    return None


def load_heading_hints(slide_xml: Path) -> Dict[str, Dict[str, object]]:
    """
    Load heading hints produced by structure analysis stage.
    key: shape_id (string)
    value: subset of hint fields
    """
    sidecar = find_sidecar_json(slide_xml)
    if sidecar is None:
        return {}
    try:
        payload = json.loads(sidecar.read_text(encoding="utf-8"))
    except Exception:
        return {}
    rows = payload.get("structure_order")
    if not isinstance(rows, list):
        rows = payload.get("reading_order")
    if not isinstance(rows, list):
        return {}
    out: Dict[str, Dict[str, object]] = {}
    for row in rows:
        if not isinstance(row, dict):
            continue
        sid = str(row.get("shape_id", "")).strip()
        if not sid:
            continue
        out[sid] = {
            "is_heading_candidate": bool(row.get("is_heading_candidate", False)),
            "heading_score": float(row.get("heading_score", 0.0)),
            "heading_depth_hint": row.get("heading_depth_hint"),
            "font_pt": row.get("font_pt"),
            "ph_type": row.get("ph_type"),
            "is_title_placeholder": bool(row.get("is_title_placeholder", False)),
        }
    return out


def rels_from_sidecar(slide_xml: Path) -> Optional[Path]:
    sidecar = find_sidecar_json(slide_xml)
    if sidecar is None:
        return None
    try:
        payload = json.loads(sidecar.read_text(encoding="utf-8"))
    except Exception:
        return None
    src = payload.get("input_xml")
    if not isinstance(src, str) or not src:
        return None
    src_path = Path(src)
    if not src_path.exists():
        return None
    rels = src_path.parent / "_rels" / f"{src_path.name}.rels"
    if rels.exists():
        return rels

    # Fallback: try finding the original slide XML under */ppt/slides/.
    base_name = slide_base_name(slide_xml)
    repo_root = Path(__file__).resolve().parent.parent
    for cand in repo_root.rglob(base_name):
        if "/ppt/slides/" not in cand.as_posix():
            continue
        rels2 = cand.parent / "_rels" / f"{cand.name}.rels"
        if rels2.exists():
            return rels2
    return None


def fallback_rels(slide_xml: Path) -> Optional[Path]:
    base = slide_base_name(slide_xml)
    cands = [
        slide_xml.parent / "_rels" / f"{base}.rels",
        slide_xml.with_name(f"{base}.rels"),
        slide_xml.parent / "_rels" / f"{slide_xml.name}.rels",
    ]
    for c in cands:
        if c.exists() and c.is_file():
            return c
    return None


def source_slide_rels(source_slide_xml: Optional[Path]) -> Optional[Path]:
    if source_slide_xml is None or not source_slide_xml.exists():
        return None
    rels = source_slide_xml.parent / "_rels" / f"{source_slide_xml.name}.rels"
    if rels.exists() and rels.is_file():
        return rels
    return None


def build_rels_map(rels_path: Optional[Path]) -> Dict[str, str]:
    if rels_path is None or not rels_path.exists():
        return {}
    root = ET.parse(rels_path).getroot()
    out: Dict[str, str] = {}
    for rel in root.findall("rel:Relationship", REL_NS):
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if rid and target:
            out[rid] = target
    return out


def choose_rels_in_package(
    slide_xml: Path,
    source_slide_xml: Optional[Path] = None,
) -> Optional[Path]:
    # Strictly stay inside same ppt package to avoid cross-package mismatches.
    sidecar = rels_from_sidecar(slide_xml)
    if sidecar:
        return sidecar
    fallback = fallback_rels(slide_xml)
    if fallback:
        return fallback
    return source_slide_rels(source_slide_xml)


def resolve_image_path(
    rels_map: Dict[str, str],
    rels_path: Optional[Path],
    r_embed: Optional[str],
) -> Tuple[str, Optional[str]]:
    if not r_embed:
        return "[unresolved-image]", "missing r:embed"
    if rels_path is None:
        return f"[unresolved-image:{r_embed}]", "missing slide rels file"
    if not rels_map:
        return f"[unresolved-image:{r_embed}]", "empty or unreadable rels map"
    target = rels_map.get(r_embed)
    if not target:
        return f"[unresolved-image:{r_embed}]", f"relationship not found: {r_embed}"

    abs_path = (rels_path.parent.parent / target).resolve()
    repo_root = Path(__file__).resolve().parent.parent
    pkg_name: Optional[str] = None
    try:
        if rels_path.parents[2].name == "ppt":
            pkg_name = rels_path.parents[3].name
    except IndexError:
        pkg_name = None
    if pkg_name:
        target_name = Path(target).name
        remapped = (repo_root / "main_converter" / "target_slides" / pkg_name / "ppt" / "media" / target_name)
        if remapped.exists():
            abs_path = remapped.absolute()
    return str(abs_path), None


def relativize_markdown_path(path: str, output_dir: Optional[Path]) -> str:
    if path.startswith("[unresolved-image") or output_dir is None:
        return path
    try:
        return os.path.relpath(path, start=str(output_dir))
    except Exception:
        return path


def _rewrite_markdown_for_per_slide(md_text: str) -> str:
    # Per-slide markdown lives in <package>/per_slide, so media links move one level up.
    return re.sub(r"\]\((media/)", r"](../\1", md_text)


def copy_media_asset(
    path: str,
    media_dir: Optional[Path],
    copied_media: Optional[Dict[str, Path]] = None,
) -> str:
    if path.startswith("[unresolved-image") or media_dir is None:
        return path

    src = Path(path)
    if not src.exists() or not src.is_file():
        return path

    try:
        src_key = str(src.resolve())
    except Exception:
        src_key = str(src)

    if copied_media is not None and src_key in copied_media:
        return str(copied_media[src_key])

    media_dir.mkdir(parents=True, exist_ok=True)
    dest = media_dir / src.name
    if dest.exists():
        try:
            same_file = dest.resolve() == src.resolve()
        except Exception:
            same_file = False
        if not same_file:
            stem = src.stem
            suffix = src.suffix
            n = 2
            while dest.exists():
                dest = media_dir / f"{stem}-{n}{suffix}"
                n += 1

    shutil.copy2(src, dest)
    if copied_media is not None:
        copied_media[src_key] = dest
    return str(dest)


def copy_debug_image_asset(
    path: str,
    debug_dir: Optional[Path],
    copied_debug_images: Optional[Dict[str, Path]] = None,
) -> Optional[str]:
    if path.startswith("[unresolved-image") or debug_dir is None:
        return None

    src = Path(path)
    if not src.exists() or not src.is_file():
        return None

    try:
        src_key = str(src.resolve())
    except Exception:
        src_key = str(src)

    if copied_debug_images is not None and src_key in copied_debug_images:
        return str(copied_debug_images[src_key])

    debug_dir.mkdir(parents=True, exist_ok=True)
    dest = debug_dir / src.name
    if dest.exists():
        try:
            same_file = dest.resolve() == src.resolve()
        except Exception:
            same_file = False
        if not same_file:
            stem = src.stem
            suffix = src.suffix
            n = 2
            while dest.exists():
                dest = debug_dir / f"{stem}-{n}{suffix}"
                n += 1

    shutil.copy2(src, dest)
    if copied_debug_images is not None:
        copied_debug_images[src_key] = dest
    return str(dest)


def format_markdown_image(
    path: str,
    output_dir: Optional[Path],
    alt_text: str = "image",
    media_dir: Optional[Path] = None,
    copied_media: Optional[Dict[str, Path]] = None,
) -> str:
    if path.startswith("[unresolved-image"):
        return path
    path = copy_media_asset(path, media_dir=media_dir, copied_media=copied_media)
    path = relativize_markdown_path(path, output_dir)
    return f"![{alt_text}]({path})"


def paragraph_text(paragraph: ET.Element) -> str:
    runs = []
    for t in paragraph.findall(".//a:t", NS):
        if t.text:
            runs.append(t.text.strip())
    return normalize_text(" ".join(x for x in runs if x))


def paragraph_level(paragraph: ET.Element) -> Optional[int]:
    p_pr = paragraph.find("./a:pPr", NS)
    if p_pr is None:
        return None
    lvl = p_pr.attrib.get("lvl")
    if lvl is None:
        return None
    try:
        return int(lvl)
    except ValueError:
        return None


def paragraph_has_list_semantics(paragraph: ET.Element) -> bool:
    p_pr = paragraph.find("./a:pPr", NS)
    if p_pr is None:
        return False
    if p_pr.attrib.get("lvl") is not None:
        return True
    if p_pr.find("./a:buChar", NS) is not None:
        return True
    if p_pr.find("./a:buAutoNum", NS) is not None:
        return True
    return False


def promote_plain_text_to_list(
    blocks: Sequence[Tuple[str, str, Optional[int]]],
) -> List[Tuple[str, str, Optional[int]]]:
    if not any(kind in {"list_ul", "list_ol"} for kind, _, _ in blocks):
        return list(blocks)

    text_blocks = [(idx, text) for idx, (kind, text, _) in enumerate(blocks) if kind == "text"]
    if len(text_blocks) < 3:
        return list(blocks)

    first_list_kind = next((kind for kind, _, _ in blocks if kind in {"list_ul", "list_ol"}), "list_ul")
    promoted = list(blocks)
    for idx, text in text_blocks:
        if len(text) > 80 or text.endswith((".", ":")):
            continue
        promoted[idx] = (first_list_kind, text, 0)
    return promoted


def normalize_list_levels(blocks: Sequence[Tuple[str, str, Optional[int]]]) -> List[Tuple[str, str, Optional[int]]]:
    levels = sorted({int(level or 0) for kind, _, level in blocks if kind in {"list_ul", "list_ol"}})
    if not levels:
        return list(blocks)
    remap = {level: idx for idx, level in enumerate(levels)}
    normalized: List[Tuple[str, str, Optional[int]]] = []
    for kind, text, level in blocks:
        if kind not in {"list_ul", "list_ol"}:
            normalized.append((kind, text, level))
            continue
        mapped = remap[int(level or 0)]
        normalized.append((kind, text, mapped))
    return normalized


def extract_shape_blocks(shape_elem: ET.Element) -> List[Tuple[str, str, Optional[int]]]:
    blocks: List[Tuple[str, str, Optional[int]]] = []
    for p in shape_elem.findall(".//p:txBody/a:p", NS):
        text = paragraph_text(p)
        if not text:
            continue
        p_pr = p.find("./a:pPr", NS)
        has_auto_num = p_pr is not None and p_pr.find("./a:buAutoNum", NS) is not None
        if paragraph_has_list_semantics(p):
            level = paragraph_level(p)
            blocks.append(("list_ol" if has_auto_num else "list_ul", text, 0 if level is None else level))
        else:
            blocks.append(("text", text, None))
    return blocks


def render_shape_blocks(blocks: Sequence[Tuple[str, str, Optional[int]]]) -> str:
    if not blocks:
        return ""
    blocks = normalize_list_levels(promote_plain_text_to_list(blocks))

    rendered: List[str] = []
    ordered_counters: Dict[int, int] = {}
    for idx, (kind, text, level) in enumerate(blocks):
        if kind in {"list_ul", "list_ol"}:
            indent = "  " * max(0, int(level or 0))
            if kind == "list_ol":
                clean_text = re.sub(r"^\d+\s*\.\s*", "", text).strip() or text
                lvl = max(0, int(level or 0))
                ordered_counters[lvl] = ordered_counters.get(lvl, 0) + 1
                for deeper in [k for k in ordered_counters.keys() if k > lvl]:
                    del ordered_counters[deeper]
                rendered.append(f"{indent}{ordered_counters[lvl]}. {clean_text}")
            else:
                rendered.append(f"{indent}- {text}")
            continue

        prev_kind = blocks[idx - 1][0] if idx > 0 else None
        next_kind = blocks[idx + 1][0] if idx + 1 < len(blocks) else None
        if prev_kind in {"list_ul", "list_ol"} and next_kind in {"list_ul", "list_ol"}:
            prev_level = max(0, int(blocks[idx - 1][2] or 0))
            indent = "  " * prev_level
            if prev_kind == "list_ol":
                clean_text = re.sub(r"^\d+\s*\.\s*", "", text).strip() or text
                ordered_counters[prev_level] = ordered_counters.get(prev_level, 0) + 1
                for deeper in [k for k in ordered_counters.keys() if k > prev_level]:
                    del ordered_counters[deeper]
                rendered.append(f"{indent}{ordered_counters[prev_level]}. {clean_text}")
            else:
                rendered.append(f"{indent}- {text}")
        else:
            rendered.append(text)
    return "\n".join(rendered).strip()


def graphic_frame_kind(graphic_frame: ET.Element) -> Optional[str]:
    graphic_data = graphic_frame.find("./a:graphic/a:graphicData", NS)
    if graphic_data is None:
        return None
    uri = graphic_data.attrib.get("uri", "").strip()
    if uri.endswith("/diagram"):
        return "diagram"
    if uri.endswith("/chart"):
        return "chart"
    return None


def diagram_data_path(graphic_frame: ET.Element, rels_path: Optional[Path], rels_map: Dict[str, str]) -> Optional[Path]:
    if rels_path is None:
        return None
    rel_ids = graphic_frame.find("./a:graphic/a:graphicData/dgm:relIds", NS)
    if rel_ids is None:
        return None
    dm_rid = rel_ids.attrib.get(f"{{{NS['r']}}}dm")
    if not dm_rid:
        return None
    target = rels_map.get(dm_rid)
    if not target:
        return None
    data_path = (rels_path.parent.parent / target).resolve()
    if data_path.exists() and data_path.is_file():
        return data_path
    return None


def extract_diagram_texts(diagram_data_xml: Path) -> List[str]:
    try:
        root = ET.parse(diagram_data_xml).getroot()
    except Exception:
        return []

    texts: List[str] = []
    seen: set = set()
    for pt in root.findall(".//dgm:pt", NS):
        raw = "".join(t.text or "" for t in pt.findall(".//a:t", NS))
        text = normalize_text(raw)
        if not text or text in seen:
            continue
        seen.add(text)
        texts.append(text)
    return texts


def format_diagram_as_markdown(texts: Sequence[str]) -> Optional[str]:
    cleaned = [normalize_text(text) for text in texts if normalize_text(text)]
    if not cleaned:
        return None
    if len(cleaned) == 1:
        return cleaned[0]
    return "\n".join(f"- {text}" for text in cleaned)


def infer_heading_depth_fallback(
    text: str,
    text_block_index: int,
    font_pt: Optional[float] = None,
) -> Optional[int]:
    return hr_infer_heading_depth_fallback(
        text=text,
        text_block_index=text_block_index,
        font_pt=font_pt,
    )


def clean_heading_text_for_render(text: str) -> str:
    return hr_clean_heading_text_for_render(text)


def looks_like_multi_numbered_items(text: str) -> bool:
    return hr_looks_like_multi_numbered_items(text)


def strict_heading_depth_from_placeholder(ph_type: Optional[str]) -> Optional[int]:
    return hr_strict_heading_depth_from_placeholder(ph_type)


def is_body_like_long_sentence(text: str) -> bool:
    return hr_is_body_like_long_sentence(text)


def normalize_triangle_bullet(text: str) -> str:
    raw = re.sub(r"\s+", " ", (text or "").strip())
    if not raw:
        return ""
    raw = re.sub(r"^(\d+)\s+\.\s*", r"\1. ", raw)
    if re.match(r"^▶+\s*", raw):
        return re.sub(r"^▶+\s*", "- ", raw)
    return raw


def split_triangle_bullets(text: str) -> List[str]:
    lines = [line.strip() for line in (text or "").splitlines() if line.strip()]
    if not lines:
        return []

    rendered_lines: List[str] = []
    for line in lines:
        raw = re.sub(r"\s+", " ", line).strip()
        if not raw:
            continue
        if "▶" not in raw:
            normalized = normalize_triangle_bullet(raw)
            if normalized:
                rendered_lines.append(normalized)
            continue

        parts = [p.strip() for p in re.split(r"\s*▶+\s*", raw) if p.strip()]
        if not parts:
            continue
        if raw.startswith("-"):
            parts[0] = re.sub(r"^-\s*", "", parts[0]).strip()
        rendered_lines.extend(f"- {part}" for part in parts if part)

    return rendered_lines


def shape_id_of(elem: ET.Element) -> str:
    c_nv_pr = elem.find(".//p:cNvPr", NS)
    if c_nv_pr is None:
        return ""
    return c_nv_pr.attrib.get("id", "")


def parse_int(v: Optional[str], default: int = 10**18) -> int:
    if v is None:
        return default
    try:
        return int(v)
    except ValueError:
        return default


def first_off(elem: ET.Element) -> Optional[ET.Element]:
    for p in (
        "./p:spPr/a:xfrm/a:off",
        "./p:grpSpPr/a:xfrm/a:off",
        "./p:xfrm/a:off",
        ".//a:off",
    ):
        off = elem.find(p, NS)
        if off is not None:
            return off
    return None


def first_ext(elem: ET.Element) -> Optional[ET.Element]:
    for p in (
        "./p:spPr/a:xfrm/a:ext",
        "./p:grpSpPr/a:xfrm/a:ext",
        "./p:xfrm/a:ext",
        ".//a:ext",
    ):
        ext = elem.find(p, NS)
        if ext is not None:
            return ext
    return None


def extract_bbox_emu(elem: ET.Element) -> Optional[Tuple[int, int, int, int]]:
    off = first_off(elem)
    ext = first_ext(elem)
    if off is None or ext is None:
        return None
    x = parse_int(off.attrib.get("x"))
    y = parse_int(off.attrib.get("y"))
    w = parse_int(ext.attrib.get("cx"))
    h = parse_int(ext.attrib.get("cy"))
    if any(v >= 10**18 for v in (x, y, w, h)) or w <= 0 or h <= 0:
        return None
    return (x, y, x + w, y + h)


def intersection_area(a: Tuple[int, int, int, int], b: Tuple[int, int, int, int]) -> int:
    left = max(a[0], b[0])
    top = max(a[1], b[1])
    right = min(a[2], b[2])
    bottom = min(a[3], b[3])
    if right <= left or bottom <= top:
        return 0
    return (right - left) * (bottom - top)


def bbox_area(bbox: Tuple[int, int, int, int]) -> int:
    return max(0, bbox[2] - bbox[0]) * max(0, bbox[3] - bbox[1])


def bbox_center(bbox: Tuple[int, int, int, int]) -> Tuple[int, int]:
    return ((bbox[0] + bbox[2]) // 2, (bbox[1] + bbox[3]) // 2)


def bbox_contains_point(bbox: Tuple[int, int, int, int], point: Tuple[int, int]) -> bool:
    return bbox[0] <= point[0] <= bbox[2] and bbox[1] <= point[1] <= bbox[3]


def overlay_link_text(
    path: str,
    output_dir: Optional[Path],
    media_dir: Optional[Path] = None,
    copied_media: Optional[Dict[str, Path]] = None,
) -> str:
    if path.startswith("[unresolved-image"):
        return path
    path = copy_media_asset(path, media_dir=media_dir, copied_media=copied_media)
    path = relativize_markdown_path(path, output_dir)
    return f"[image]({path})"


def annotate_generated_image_markdown(markdown: str, image_path: str) -> str:
    image_name = Path(image_path).name
    body = markdown.strip()
    if not body:
        return f"[image-vlm-source: {image_name}]"
    return f"[image-vlm-source: {image_name}]\n\n{body}"


def overlay_content_text(
    path: str,
    output_dir: Optional[Path],
    media_dir: Optional[Path] = None,
    copied_media: Optional[Dict[str, Path]] = None,
    image_vlm_model: Optional[str] = None,
    image_vlm_prompt: str = DEFAULT_IMAGE_VLM_PROMPT,
    image_vlm_max_new_tokens: int = DEFAULT_IMAGE_VLM_MAX_NEW_TOKENS,
) -> Tuple[str, Optional[str], bool, bool, bool]:
    if path.startswith("[unresolved-image"):
        return path, None, False, False, False

    if image_vlm_model:
        image_md, image_warn, unavailable, result = convert_picture_to_markdown(
            path,
            model_spec=image_vlm_model,
            prompt=image_vlm_prompt,
            max_new_tokens=image_vlm_max_new_tokens,
        )
        if image_md is not None:
            return annotate_generated_image_markdown(image_md, path), None, unavailable, True, False
        skipped_no_markdown = isinstance(result, dict) and str(result.get("status", "")) == "no_markdown"
        return (
            overlay_link_text(path, output_dir, media_dir=media_dir, copied_media=copied_media),
            image_warn,
            unavailable,
            False,
            skipped_no_markdown,
        )

    return (
        overlay_link_text(path, output_dir, media_dir=media_dir, copied_media=copied_media),
        None,
        False,
        False,
        False,
    )


def collect_table_overlay_pictures(
    sp_tree: ET.Element,
    slide_xml: Path,
    rels_map: Dict[str, str],
    rels_path: Optional[Path],
) -> Tuple[Dict[str, List[Dict[str, object]]], set, List[str], int, int]:
    table_bboxes: List[Tuple[str, Tuple[int, int, int, int]]] = []
    picture_infos: List[Dict[str, object]] = []

    for child in list(sp_tree):
        tag = local_name(child.tag)
        if tag == "graphicFrame":
            tbl = child.find(".//a:tbl", NS)
            bbox = extract_bbox_emu(child)
            sid = shape_id_of(child)
            if tbl is not None and bbox is not None and sid:
                table_bboxes.append((sid, bbox))
        elif tag == "pic":
            bbox = extract_bbox_emu(child)
            sid = shape_id_of(child)
            if bbox is None or not sid:
                continue
            blip = child.find(".//a:blip", NS)
            embed = blip.attrib.get(f"{{{NS['r']}}}embed") if blip is not None else None
            path, warn = resolve_image_path(rels_map, rels_path, embed)
            picture_infos.append(
                {
                    "shape_id": sid,
                    "bbox": bbox,
                    "path": path,
                    "warn": warn,
                }
            )

    by_table: Dict[str, List[Dict[str, object]]] = {}
    consumed: set = set()
    warnings: List[str] = []
    resolved = 0
    unresolved = 0
    for _, table_bbox in table_bboxes:
        table_area = bbox_area(table_bbox)
        if table_area <= 0:
            continue
        for pic_info in picture_infos:
            pic_id = str(pic_info["shape_id"])
            pic_bbox = pic_info["bbox"]
            pic_area = bbox_area(pic_bbox)
            if pic_area <= 0:
                continue
            overlap = intersection_area(table_bbox, pic_bbox)
            overlap_ratio = overlap / pic_area
            center_inside = bbox_contains_point(table_bbox, bbox_center(pic_bbox))
            if center_inside and overlap_ratio >= 0.8:
                consumed.add(pic_id)
                by_table.setdefault(_, []).append(pic_info)
                warn = pic_info.get("warn")
                if isinstance(warn, str) and warn:
                    unresolved += 1
                    warnings.append(warn)
                else:
                    resolved += 1
    return by_table, consumed, warnings, resolved, unresolved


def compute_table_cell_bounds(
    graphic_frame: ET.Element,
) -> Optional[Tuple[List[Tuple[int, int]], List[Tuple[int, int]]]]:
    bbox = extract_bbox_emu(graphic_frame)
    tbl = graphic_frame.find(".//a:tbl", NS)
    if bbox is None or tbl is None:
        return None

    col_elems = tbl.findall("./a:tblGrid/a:gridCol", NS)
    row_elems = tbl.findall("./a:tr", NS)
    col_widths = [parse_int(col.attrib.get("w"), 0) for col in col_elems]
    row_heights = [parse_int(row.attrib.get("h"), 0) for row in row_elems]
    total_col = sum(v for v in col_widths if v > 0)
    total_row = sum(v for v in row_heights if v > 0)
    if total_col <= 0 or total_row <= 0:
        return None

    table_width = bbox[2] - bbox[0]
    table_height = bbox[3] - bbox[1]
    if table_width <= 0 or table_height <= 0:
        return None

    col_bounds: List[Tuple[int, int]] = []
    cursor = bbox[0]
    used_width = 0
    for idx, width in enumerate(col_widths):
        if idx == len(col_widths) - 1:
            right = bbox[2]
        else:
            used_width += width
            right = bbox[0] + int(table_width * used_width / total_col)
        col_bounds.append((cursor, right))
        cursor = right

    row_bounds: List[Tuple[int, int]] = []
    cursor = bbox[1]
    used_height = 0
    for idx, height in enumerate(row_heights):
        if idx == len(row_heights) - 1:
            bottom = bbox[3]
        else:
            used_height += height
            bottom = bbox[1] + int(table_height * used_height / total_row)
        row_bounds.append((cursor, bottom))
        cursor = bottom

    return col_bounds, row_bounds


def find_table_cell_origin(
    parsed_table: Dict[str, object],
    row_idx: int,
    col_idx: int,
) -> Tuple[int, int]:
    rows = parsed_table.get("rows")
    if not isinstance(rows, list):
        return row_idx, col_idx
    if row_idx >= len(rows):
        return row_idx, col_idx
    row = rows[row_idx]
    if not isinstance(row, list) or col_idx >= len(row):
        return row_idx, col_idx
    cell = row[col_idx]
    if not isinstance(cell, dict):
        return row_idx, col_idx
    origin = cell.get("origin")
    if cell.get("type") in {"hMerge", "vMerge"} and isinstance(origin, list) and len(origin) == 2:
        if isinstance(origin[0], int) and isinstance(origin[1], int):
            return origin[0], origin[1]
    return row_idx, col_idx


def inject_table_overlay_content(
    parsed_table: Dict[str, object],
    graphic_frame: ET.Element,
    overlays: Sequence[Dict[str, object]],
    output_dir: Optional[Path],
    media_dir: Optional[Path] = None,
    copied_media: Optional[Dict[str, Path]] = None,
    image_vlm_model: Optional[str] = None,
    image_vlm_prompt: str = DEFAULT_IMAGE_VLM_PROMPT,
    image_vlm_max_new_tokens: int = DEFAULT_IMAGE_VLM_MAX_NEW_TOKENS,
) -> Tuple[Dict[str, object], Dict[str, object]]:
    overlay_stats: Dict[str, object] = {
        "image_markdown_blocks": 0,
        "image_markdown_skipped_blocks": 0,
        "image_markdown_failed_blocks": 0,
        "warnings": [],
        "pipeline_unavailable": False,
    }
    if not overlays:
        return parsed_table, overlay_stats

    bounds = compute_table_cell_bounds(graphic_frame)
    if bounds is None:
        return parsed_table, overlay_stats
    col_bounds, row_bounds = bounds

    rows = parsed_table.get("rows")
    if not isinstance(rows, list):
        return parsed_table, overlay_stats

    for overlay in overlays:
        bbox = overlay.get("bbox")
        if not (isinstance(bbox, tuple) and len(bbox) == 4):
            continue
        center = bbox_center(bbox)
        row_idx = next((idx for idx, (top, bottom) in enumerate(row_bounds) if top <= center[1] <= bottom), None)
        col_idx = next((idx for idx, (left, right) in enumerate(col_bounds) if left <= center[0] <= right), None)
        if row_idx is None or col_idx is None:
            continue
        origin_row, origin_col = find_table_cell_origin(parsed_table, row_idx, col_idx)
        if origin_row >= len(rows):
            continue
        row = rows[origin_row]
        if not isinstance(row, list) or origin_col >= len(row):
            continue
        cell = row[origin_col]
        if not isinstance(cell, dict):
            continue
        existing = normalize_text(str(cell.get("text", "")))
        content_text, content_warn, unavailable, converted, skipped_no_markdown = overlay_content_text(
            str(overlay.get("path", "")),
            output_dir,
            media_dir=media_dir,
            copied_media=copied_media,
            image_vlm_model=image_vlm_model,
            image_vlm_prompt=image_vlm_prompt,
            image_vlm_max_new_tokens=image_vlm_max_new_tokens,
        )
        updated = f"{existing}\n{content_text}".strip() if existing else content_text
        cell["text"] = updated
        if converted:
            overlay_stats["image_markdown_blocks"] = int(overlay_stats.get("image_markdown_blocks", 0)) + 1
        elif skipped_no_markdown:
            overlay_stats["image_markdown_skipped_blocks"] = int(
                overlay_stats.get("image_markdown_skipped_blocks", 0)
            ) + 1
        elif image_vlm_model and not str(overlay.get("path", "")).startswith("[unresolved-image"):
            overlay_stats["image_markdown_failed_blocks"] = int(
                overlay_stats.get("image_markdown_failed_blocks", 0)
            ) + 1
        if isinstance(content_warn, str) and content_warn.strip():
            warnings = overlay_stats.get("warnings")
            if isinstance(warnings, list):
                warnings.append(content_warn)
            if unavailable:
                overlay_stats["pipeline_unavailable"] = True

    return parsed_table, overlay_stats


def convert_table_to_markdown(
    graphic_frame: ET.Element,
    overlays: Optional[Sequence[Dict[str, object]]] = None,
    output_dir: Optional[Path] = None,
    media_dir: Optional[Path] = None,
    copied_media: Optional[Dict[str, Path]] = None,
    image_vlm_model: Optional[str] = None,
    image_vlm_prompt: str = DEFAULT_IMAGE_VLM_PROMPT,
    image_vlm_max_new_tokens: int = DEFAULT_IMAGE_VLM_MAX_NEW_TOKENS,
) -> Tuple[Optional[str], Optional[str], Dict[str, object]]:
    tbl = graphic_frame.find(".//a:tbl", NS)
    if tbl is None:
        return None, "graphicFrame without a:tbl", {
            "image_markdown_blocks": 0,
            "image_markdown_skipped_blocks": 0,
            "image_markdown_failed_blocks": 0,
            "warnings": [],
            "pipeline_unavailable": False,
        }

    from table_pipeline import parse as table_parse  # type: ignore
    from table_pipeline import render as table_render  # type: ignore

    parsed = table_parse.parse_table_element(tbl, source="<slide_table>")
    parsed, overlay_stats = inject_table_overlay_content(
        parsed,
        graphic_frame,
        overlays or [],
        output_dir,
        media_dir=media_dir,
        copied_media=copied_media,
        image_vlm_model=image_vlm_model,
        image_vlm_prompt=image_vlm_prompt,
        image_vlm_max_new_tokens=image_vlm_max_new_tokens,
    )
    md = table_render.render_parsed_table_to_markdown(
        parsed_table=parsed,
        header_rows=1,
        fill_merged=table_render.FILL_BOTH,
    )
    return md, None, overlay_stats


def _load_image_markdown_pipeline() -> Tuple[Optional[object], Optional[str]]:
    global _IMAGE_MARKDOWN_PIPELINE_MODULE, _IMAGE_MARKDOWN_PIPELINE_IMPORT_ERROR
    if _IMAGE_MARKDOWN_PIPELINE_MODULE is not None:
        return _IMAGE_MARKDOWN_PIPELINE_MODULE, None
    if _IMAGE_MARKDOWN_PIPELINE_IMPORT_ERROR is not None:
        return None, _IMAGE_MARKDOWN_PIPELINE_IMPORT_ERROR

    import_errors: List[str] = []
    for module_name in ("image_pipeline.service", "image_markdown_pipeline"):
        try:
            _IMAGE_MARKDOWN_PIPELINE_MODULE = importlib.import_module(module_name)
            return _IMAGE_MARKDOWN_PIPELINE_MODULE, None
        except Exception as exc:  # noqa: BLE001
            import_errors.append(f"{module_name}: {type(exc).__name__}: {exc}")

    _IMAGE_MARKDOWN_PIPELINE_IMPORT_ERROR = "; ".join(import_errors)
    return None, _IMAGE_MARKDOWN_PIPELINE_IMPORT_ERROR


def convert_picture_to_markdown(
    image_path: str,
    *,
    model_spec: str,
    prompt: str,
    max_new_tokens: int,
) -> Tuple[Optional[str], Optional[str], bool, Optional[Dict[str, object]]]:
    module, import_error = _load_image_markdown_pipeline()
    if module is None:
        return None, f"image markdown pipeline unavailable: {import_error}", True, None
    try:
        result = module.extract_markdown_from_image(
            Path(image_path),
            model_spec=model_spec,
            prompt=prompt,
            max_new_tokens=max_new_tokens,
        )
    except Exception as exc:  # noqa: BLE001
        return (
            None,
            f"image markdown pipeline failed on {Path(image_path).name}: {type(exc).__name__}: {exc}",
            False,
            None,
        )

    if not isinstance(result, dict):
        return None, f"image markdown pipeline returned invalid payload: {type(result).__name__}", False, None

    _log_image_markdown_result(image_path, result)

    status = str(result.get("status", "error"))
    if status == "markdown":
        markdown = result.get("markdown")
        if isinstance(markdown, str) and markdown.strip():
            return markdown, None, False, result
        return None, f"image markdown pipeline rendered empty markdown: {Path(image_path).name}", False, result
    if status == "no_markdown":
        return None, None, False, result

    error = result.get("error")
    if not isinstance(error, str) or not error.strip():
        error = "unknown image markdown pipeline error"
    return None, f"{Path(image_path).name}: {error}", False, result


def _to_jsonable_copy(value: object) -> object:
    return json.loads(json.dumps(value, ensure_ascii=False))


def _log_image_markdown_result(image_path: str, result: Dict[str, object]) -> None:
    name = Path(image_path).name
    status = str(result.get("status", "error"))
    model_alias = result.get("model_alias")
    model_id = result.get("model_id")
    elapsed_sec = result.get("elapsed_sec")
    error = result.get("error")
    reason = result.get("reason")
    fallback = result.get("fallback")
    markdown = result.get("markdown")
    markdown_chars = len(markdown.strip()) if isinstance(markdown, str) else 0
    elapsed_label = f"{float(elapsed_sec):.3f}s" if isinstance(elapsed_sec, (int, float)) else "n/a"

    summary = (
        f"[image-vlm] {name} "
        f"status={status} "
        f"model={model_alias or model_id or 'unknown'} "
        f"chars={markdown_chars} "
        f"elapsed={elapsed_label}"
    )
    if isinstance(error, str) and error.strip():
        summary = f"{summary} error={error}"
    elif isinstance(reason, str) and reason.strip():
        summary = f"{summary} reason={reason}"
        if isinstance(fallback, str) and fallback.strip():
            summary = f"{summary} fallback={fallback}"
    print(summary, flush=True)

    if _IMAGE_VLM_DEBUG_JSON:
        debug_payload = {
            "kind": "image_vlm_debug",
            "file": str(Path(str(result.get('file', image_path))).expanduser().resolve()),
            "image_name": name,
            "status": status,
            "model_alias": model_alias,
            "model_id": model_id,
            "elapsed_sec": elapsed_sec if isinstance(elapsed_sec, (int, float)) else None,
            "max_new_tokens": result.get("max_new_tokens"),
            "error": error if isinstance(error, str) and error.strip() else None,
            "reason": reason if isinstance(reason, str) and reason.strip() else None,
            "fallback": fallback if isinstance(fallback, str) and fallback.strip() else None,
            "markdown_chars": markdown_chars,
        }
        print(json.dumps(_to_jsonable_copy(debug_payload), ensure_ascii=False, separators=(",", ":")))


def convert_one_slide(
    slide_xml: Path,
    page_no: int,
    source_slide_xml: Optional[Path] = None,
    output_dir: Optional[Path] = None,
    media_dir: Optional[Path] = None,
    copied_media: Optional[Dict[str, Path]] = None,
    image_vlm_model: Optional[str] = None,
    image_vlm_prompt: str = DEFAULT_IMAGE_VLM_PROMPT,
    image_vlm_max_new_tokens: int = DEFAULT_IMAGE_VLM_MAX_NEW_TOKENS,
    strict_headings: bool = False,
) -> Tuple[str, Dict[str, object]]:
    heading_policy = HeadingPolicy(strict=strict_headings)
    root = ET.parse(slide_xml).getroot()
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        raise ValueError("missing p:cSld/p:spTree")

    rels_path = choose_rels_in_package(slide_xml, source_slide_xml=source_slide_xml)
    rels_map = build_rels_map(rels_path)
    heading_hints = load_heading_hints(slide_xml)
    table_overlay_map, consumed_picture_ids, overlay_warnings, overlay_resolved, overlay_unresolved = (
        collect_table_overlay_pictures(sp_tree, slide_xml, rels_map, rels_path)
    )

    lines: List[str] = [f"[Page_{page_no}]", ""]
    used_headings: set = set()
    text_block_index = 0

    stats = {
        "blocks_total": 0,
        "text_blocks": 0,
        "image_blocks": 0,
        "image_markdown_blocks": 0,
        "image_markdown_skipped_blocks": 0,
        "image_markdown_failed_blocks": 0,
        "table_blocks": 0,
        "unsupported_blocks": 0,
        "skipped_blocks": 0,
        "resolved_images": 0,
        "unresolved_images": 0,
        "warnings": list(overlay_warnings),
        "rels_path": str(rels_path) if rels_path else None,
    }
    stats["resolved_images"] += overlay_resolved
    stats["unresolved_images"] += overlay_unresolved
    image_pipeline_unavailable_reported = False

    for child in list(sp_tree):
        tag = local_name(child.tag)
        if tag not in {"sp", "pic", "graphicFrame", "grpSp", "cxnSp"}:
            continue
        stats["blocks_total"] += 1

        if tag == "cxnSp":
            stats["skipped_blocks"] += 1
            continue

        if tag in {"sp", "grpSp"}:
            ph = child.find(".//p:ph", NS)
            ph_type = ph.attrib.get("type") if ph is not None else None
            if ph_type in {"sldNum", "ftr", "dt"}:
                stats["skipped_blocks"] += 1
                continue
            shape_blocks = extract_shape_blocks(child)
            has_list_semantics = any(kind in {"list_ul", "list_ol"} for kind, _, _ in shape_blocks)
            text = render_shape_blocks(shape_blocks)
            if not text:
                stats["skipped_blocks"] += 1
                continue
            if re.fullmatch(r"\d+", text):
                stats["skipped_blocks"] += 1
                continue
            sid = shape_id_of(child)
            hint = heading_hints.get(sid, {})
            depth = hint.get("heading_depth_hint")
            score = float(hint.get("heading_score", 0.0))
            is_candidate = bool(hint.get("is_heading_candidate", False))
            raw_font_pt = hint.get("font_pt")
            try:
                font_pt = float(raw_font_pt) if raw_font_pt is not None else None
            except (TypeError, ValueError):
                font_pt = None

            if strict_headings:
                strict_ph_type = ph_type if ph_type is not None else hint.get("ph_type")
                strict_depth = strict_heading_depth_from_placeholder(strict_ph_type)
                is_candidate = strict_depth is not None
                depth = strict_depth
                score = 1.0 if is_candidate else 0.0
            else:
                non_strict_ph_type = ph_type if ph_type is not None else hint.get("ph_type")
                non_strict_depth = strict_heading_depth_from_placeholder(non_strict_ph_type)
                if non_strict_depth is not None:
                    depth = non_strict_depth
                    score = max(score, 0.9)
                    is_candidate = True
                if not is_candidate:
                    fb_depth = infer_heading_depth_fallback(text, text_block_index, font_pt=font_pt)
                    if fb_depth is not None:
                        depth = fb_depth
                        score = 0.8
                        is_candidate = True

            rendered = text
            if not strict_headings:
                if has_list_semantics:
                    is_candidate = False
                if looks_like_multi_numbered_items(rendered):
                    is_candidate = False
                if is_body_like_long_sentence(rendered):
                    is_candidate = False
            # Final markdown heading level rendering is converter responsibility.
            heading_threshold = heading_policy.threshold
            if is_candidate and isinstance(depth, int) and 1 <= depth <= 6 and score >= heading_threshold:
                heading_text = text if strict_headings else clean_heading_text_for_render(text)
                key = normalize_text(heading_text)
                if key not in used_headings:
                    rendered = f"{'#' * depth} {heading_text}"
                    used_headings.add(key)
                else:
                    # Deduplicate repeated heading text.
                    stats["skipped_blocks"] += 1
                    continue

            if rendered.startswith("#"):
                lines.append(rendered)
                lines.append("")
            else:
                rendered_lines = split_triangle_bullets(rendered)
                if rendered_lines:
                    lines.extend(rendered_lines)
                    lines.append("")
                else:
                    lines.append(rendered)
                    lines.append("")
            stats["text_blocks"] += 1
            text_block_index += 1
            continue

        if tag == "pic":
            sid = shape_id_of(child)
            if sid and sid in consumed_picture_ids:
                stats["skipped_blocks"] += 1
                continue
            blip = child.find(".//a:blip", NS)
            embed = blip.attrib.get(f"{{{NS['r']}}}embed") if blip is not None else None
            img_path, warn = resolve_image_path(rels_map, rels_path, embed)
            if warn:
                stats["unresolved_images"] += 1
                stats["warnings"].append(warn)
            else:
                stats["resolved_images"] += 1

            if image_vlm_model and not warn and not img_path.startswith("[unresolved-image"):
                image_md, image_warn, unavailable, image_result = convert_picture_to_markdown(
                    img_path,
                    model_spec=image_vlm_model,
                    prompt=image_vlm_prompt,
                    max_new_tokens=image_vlm_max_new_tokens,
                )
                if image_md is not None:
                    lines.append(annotate_generated_image_markdown(image_md, img_path))
                    lines.append("")
                    stats["image_markdown_blocks"] += 1
                    continue
                if isinstance(image_result, dict) and str(image_result.get("status", "")) == "no_markdown":
                    stats["image_markdown_skipped_blocks"] += 1
                else:
                    stats["image_markdown_failed_blocks"] += 1
                if image_warn:
                    if unavailable:
                        if not image_pipeline_unavailable_reported:
                            stats["warnings"].append(image_warn)
                            image_pipeline_unavailable_reported = True
                    else:
                        stats["warnings"].append(image_warn)

            lines.append(
                format_markdown_image(
                    img_path,
                    output_dir=output_dir,
                    media_dir=media_dir,
                    copied_media=copied_media,
                )
            )
            lines.append("")
            stats["image_blocks"] += 1
            continue

        if tag == "graphicFrame":
            table_md, err, overlay_stats = convert_table_to_markdown(
                child,
                overlays=table_overlay_map.get(shape_id_of(child), []),
                output_dir=output_dir,
                media_dir=media_dir,
                copied_media=copied_media,
                image_vlm_model=image_vlm_model,
                image_vlm_prompt=image_vlm_prompt,
                image_vlm_max_new_tokens=image_vlm_max_new_tokens,
            )
            stats["image_markdown_blocks"] += int(overlay_stats.get("image_markdown_blocks", 0))
            stats["image_markdown_skipped_blocks"] += int(overlay_stats.get("image_markdown_skipped_blocks", 0))
            stats["image_markdown_failed_blocks"] += int(overlay_stats.get("image_markdown_failed_blocks", 0))
            overlay_pipeline_unavailable = bool(overlay_stats.get("pipeline_unavailable"))
            overlay_warnings_list = overlay_stats.get("warnings")
            if isinstance(overlay_warnings_list, list):
                for overlay_warn in overlay_warnings_list:
                    if not isinstance(overlay_warn, str) or not overlay_warn.strip():
                        continue
                    if overlay_pipeline_unavailable:
                        if not image_pipeline_unavailable_reported:
                            stats["warnings"].append(overlay_warn)
                            image_pipeline_unavailable_reported = True
                    else:
                        stats["warnings"].append(overlay_warn)
            if table_md is not None:
                lines.append(table_md.strip())
                lines.append("")
                stats["table_blocks"] += 1
            else:
                gf_kind = graphic_frame_kind(child)
                if gf_kind == "diagram":
                    diagram_path = diagram_data_path(child, rels_path, rels_map)
                    diagram_text = format_diagram_as_markdown(
                        extract_diagram_texts(diagram_path) if diagram_path else []
                    )
                    if diagram_text:
                        lines.append(diagram_text)
                        lines.append("")
                        stats["text_blocks"] += 1
                    else:
                        lines.append("[unsupported: graphicFrame(non-table)]")
                        lines.append("")
                        stats["unsupported_blocks"] += 1
                        stats["warnings"].append("diagram text extraction failed")
                else:
                    lines.append("[unsupported: graphicFrame(non-table)]")
                    lines.append("")
                    stats["unsupported_blocks"] += 1
                if err:
                    stats["warnings"].append(err)
            continue

    lines = normalize_single_heading_to_h1(lines)

    md_text = "\n".join(lines).rstrip() + "\n"
    return md_text, stats


def resolve_surya_structure_dir(surya_root: Path, package_name: str) -> Path:
    def has_reordered_xmls(root: Path) -> bool:
        return root.exists() and root.is_dir() and any(root.glob("slide*.reordered.xml"))

    manifest_here = surya_root / "structure_analysis_manifest.json"
    if manifest_here.exists():
        return surya_root

    if has_reordered_xmls(surya_root):
        return surya_root

    candidate = surya_root / package_name
    if has_reordered_xmls(candidate):
        return candidate

    manifest_there = candidate / "structure_analysis_manifest.json"
    if manifest_there.exists():
        return candidate

    raise FileNotFoundError(
        f"surya structure-ready output not found for package '{package_name}' under {surya_root}"
    )


def run_surya_pipeline_stage(
    surya_root: Path,
    force: bool = False,
    targets: Optional[Sequence[str]] = None,
    target_pptx_dir: Optional[Path] = None,
    target_slides_dir: Optional[Path] = None,
) -> Path:
    def choose_python_for_surya() -> str:
        # Prefer project venv when available so `surya` import resolves reliably.
        repo_root = Path(__file__).resolve().parent.parent
        venv_py = repo_root / ".venv" / "bin" / "python"
        if venv_py.exists() and venv_py.is_file():
            return str(venv_py)
        return sys.executable

    run_script = surya_root / "run_surya_pipeline.py"
    if not run_script.exists():
        raise FileNotFoundError(f"surya pipeline script not found: {run_script}")

    py_exe = choose_python_for_surya()
    cmd = [
        py_exe,
        str(run_script),
    ]
    if target_pptx_dir is not None:
        cmd.extend(["--target-pptx-dir", str(target_pptx_dir)])
    if target_slides_dir is not None:
        cmd.extend(["--target-slides-dir", str(target_slides_dir)])
        cmd.append("--prefer-existing-target-slides")
    if force:
        cmd.append("--force")
    if targets:
        cmd.extend(str(t) for t in targets if str(t).strip())

    print(f"[surya] Running pipeline: {' '.join(cmd)}")
    proc = subprocess.run(cmd, text=True, cwd=str(surya_root))
    if proc.returncode != 0:
        raise RuntimeError(
            "surya pipeline failed\n"
            f"cmd: {' '.join(cmd)}\n"
        )
    structure_root = surya_root / "output" / "structure_ready"
    if not structure_root.exists() or not structure_root.is_dir():
        raise FileNotFoundError(f"surya structure-ready output not found after pipeline run: {structure_root}")
    return structure_root


def prepare_surya_structure_root(
    force: bool = False,
    reuse_existing_output: bool = False,
    targets: Optional[Sequence[str]] = None,
    target_pptx_dir: Optional[Path] = None,
    target_slides_dir: Optional[Path] = None,
) -> Path:
    repo_root = Path(__file__).resolve().parent.parent
    candidate = repo_root / "surya_pipeline"

    if reuse_existing_output:
        if (candidate / "structure_analysis_manifest.json").exists():
            return candidate
        structure_root = candidate / "output" / "structure_ready"
        if structure_root.exists() and structure_root.is_dir():
            return structure_root

    # Surya pipeline root passed in or inferred.
    if (candidate / "run_surya_pipeline.py").exists():
        return run_surya_pipeline_stage(
            surya_root=candidate,
            force=force,
            targets=targets,
            target_pptx_dir=target_pptx_dir,
            target_slides_dir=target_slides_dir,
        )

    if reuse_existing_output:
        structure_root = candidate / "output" / "structure_ready"
        if structure_root.exists() and structure_root.is_dir():
            return structure_root

    raise FileNotFoundError(
        f"valid surya input not found under fixed surya_pipeline path: {candidate}"
    )


class ConverterConfig(BaseModel):
    model_config = ConfigDict(arbitrary_types_allowed=True)

    cwd: Path
    repo_root: Path
    output_dir: Path
    debug_output_dir: Path
    inputs: List[str] = Field(default_factory=list)
    per_slide: bool = False
    reading_order: str = "xml"
    strict: bool = False
    reuse_surya_cache: bool = False
    image_vlm_model: Optional[str] = None
    image_vlm_prompt: str = DEFAULT_IMAGE_VLM_PROMPT
    image_vlm_max_new_tokens: int = DEFAULT_IMAGE_VLM_MAX_NEW_TOKENS


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert extracted PPTX package(s) to markdown."
    )
    parser.add_argument(
        "inputs",
        nargs="*",
        help=(
            "Optional .pptx selections (e.g., sample3.pptx sample4.pptx). "
            "If omitted, all .pptx files under main_converter/target_pptx are extracted/processed."
        ),
    )
    parser.add_argument(
        "--per-slide",
        action="store_true",
        help="Also write per-slide markdown files under main_converter/output/.../<package>/per_slide.",
    )
    parser.add_argument(
        "--reading-order",
        choices=("xml", "surya"),
        default="xml",
        help="Reading-order strategy. Default uses legacy XML-only ordering.",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Use strict heading detection in xml reading-order mode only.",
    )
    parser.add_argument(
        "--reuse-surya-cache",
        action="store_true",
        help="Reuse existing Surya structure_ready outputs instead of re-running the Surya pipeline.",
    )
    parser.add_argument(
        "--image-vlm-model",
        choices=("3b", "7b"),
        help="Convert every image block with a local Qwen2.5-VL model instead of leaving raw image links.",
    )
    parser.add_argument(
        "--image-vlm-prompt",
        default=DEFAULT_IMAGE_VLM_PROMPT,
        help="Prompt passed to the local image VLM when --image-vlm-model is enabled.",
    )
    parser.add_argument(
        "--image-vlm-max-new-tokens",
        type=int,
        default=DEFAULT_IMAGE_VLM_MAX_NEW_TOKENS,
        help="Maximum number of tokens to generate per image when --image-vlm-model is enabled.",
    )
    return parser.parse_args()


def _build_config(args: argparse.Namespace) -> ConverterConfig:
    repo_root = Path(__file__).resolve().parent.parent
    main_converter_root = repo_root / "main_converter"
    return ConverterConfig(
        cwd=main_converter_root,
        repo_root=repo_root,
        output_dir=main_converter_root / "output" / args.reading_order,
        debug_output_dir=main_converter_root / "output" / args.reading_order,
        inputs=list(args.inputs),
        per_slide=bool(args.per_slide),
        reading_order=str(args.reading_order),
        strict=bool(args.strict),
        reuse_surya_cache=bool(args.reuse_surya_cache),
        image_vlm_model=(str(args.image_vlm_model).strip().lower() if args.image_vlm_model else None),
        image_vlm_prompt=str(args.image_vlm_prompt),
        image_vlm_max_new_tokens=max(1, int(args.image_vlm_max_new_tokens)),
    )


def _resolve_prepared_inputs(config: ConverterConfig) -> Tuple[Optional[List[str]], Optional[int]]:
    if config.inputs:
        non_pptx_inputs = [x for x in config.inputs if Path(x).suffix.lower() != ".pptx"]
        if non_pptx_inputs:
            print("Only .pptx inputs are allowed.")
            print("Provide files like: sample1.pptx sample2.pptx")
            for item in non_pptx_inputs:
                print(f"- {item}")
            return None, 1
        prepared_inputs, missing_inputs = prepare_package_inputs(config.cwd, config.inputs, force_extract=True)
        if missing_inputs:
            print("Input .pptx file not found.")
            for row in missing_inputs:
                if not isinstance(row, dict):
                    continue
                item = str(row.get("input", "")).strip()
                checked = row.get("checked")
                print(f"- requested: {item or '(unknown)'}")
                if isinstance(checked, list):
                    for cand in checked:
                        print(f"  checked: {cand}")
            available_inputs = collect_target_pptx_inputs(config.cwd)
            if available_inputs:
                print("Available .pptx files under main_converter/target_pptx:")
                for path in available_inputs:
                    print(f"- {Path(path).name}")
            else:
                print(f"No .pptx files found in: {preferred_pptx_input_dir(config.cwd).resolve()}")
            return None, 1
        return prepared_inputs, None

    auto_pptx_inputs = collect_target_pptx_inputs(config.cwd)
    if not auto_pptx_inputs:
        print(f"No .pptx files found in: {preferred_pptx_input_dir(config.cwd).resolve()}")
        return None, 0
    prepared_inputs, _ = prepare_package_inputs(config.cwd, auto_pptx_inputs, force_extract=True)
    return prepared_inputs, None


def _resolve_packages(config: ConverterConfig, prepared_inputs: Sequence[str]) -> List[Path]:
    target_dirs = default_target_dirs(config.cwd)
    packages = pick_packages(target_dirs, prepared_inputs)
    if not packages:
        print("No valid PPTX package directories found.")
        if target_dirs:
            print("Checked default directories:")
            for d in target_dirs:
                print(f"- {d.resolve()}")
        else:
            print("Checked default directories: none found (expected ./target_slides).")
    return packages


def _prepare_surya_context(config: ConverterConfig, packages: Sequence[Path]) -> Optional[Path]:
    if config.reading_order != "surya":
        return None
    if config.strict:
        print("[info] --strict is ignored for surya mode. surya heading logic remains separate.")
    shared_target_slides_dir = preferred_target_dir(config.cwd).resolve()
    shared_target_pptx_dir = preferred_pptx_input_dir(config.cwd).resolve()
    return prepare_surya_structure_root(
        force=not config.reuse_surya_cache,
        reuse_existing_output=config.reuse_surya_cache,
        targets=[pkg.name for pkg in packages],
        target_pptx_dir=shared_target_pptx_dir,
        target_slides_dir=shared_target_slides_dir,
    )


def _new_manifest() -> Dict[str, object]:
    return {
        "started_at": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "packages": [],
        "summary": {
            "processed_packages": 0,
            "processed_slides": 0,
            "failed": 0,
            "resolved_images": 0,
            "unresolved_images": 0,
            "image_markdown_blocks": 0,
            "image_markdown_skipped_blocks": 0,
            "image_markdown_failed_blocks": 0,
            "table_blocks": 0,
            "elapsed_sec": 0.0,
        },
    }


def _append_package_stage_failure(
    pkg_row: Dict[str, object],
    slide_xmls: Sequence[Path],
    error_message: str,
    manifest: Dict[str, object],
    package_name: str,
    stage_label: str,
) -> None:
    summary = manifest.get("summary")
    if not isinstance(summary, dict):
        return
    for slide_xml in slide_xmls:
        row = {
            "page": parse_slide_number(slide_xml.name, 0),
            "source_xml": str(slide_xml),
            "status": "failed",
            "error": error_message,
            "warnings": [],
        }
        slides = pkg_row.get("slides")
        if isinstance(slides, list):
            slides.append(row)
        summary["failed"] = int(summary.get("failed", 0)) + 1
    packages = manifest.get("packages")
    if isinstance(packages, list):
        packages.append(pkg_row)
    print(f"[{package_name}] {stage_label} failed: {error_message}")


def _convert_package(
    config: ConverterConfig,
    pkg: Path,
    surya_structure_root: Optional[Path],
    manifest: Dict[str, object],
) -> None:
    package_started_at = time.perf_counter()
    summary = manifest.get("summary")
    packages = manifest.get("packages")
    if not isinstance(summary, dict) or not isinstance(packages, list):
        raise ValueError("invalid manifest payload shape")

    pkg_name = pkg.name
    slides_dir = pkg / "ppt" / "slides"
    slide_xmls = sorted(
        [p for p in slides_dir.glob("slide*.xml") if p.is_file()],
        key=lambda p: natural_key(p.name),
    )
    pkg_out = config.output_dir / pkg_name
    per_slide_dir = pkg_out / "per_slide"
    media_dir = pkg_out / "media"
    copied_media: Dict[str, Path] = {}
    pkg_out.mkdir(parents=True, exist_ok=True)
    if config.per_slide:
        per_slide_dir.mkdir(parents=True, exist_ok=True)
    elif per_slide_dir.exists():
        shutil.rmtree(per_slide_dir)

    pkg_row: Dict[str, object] = {
        "package": str(pkg),
        "name": pkg_name,
        "slides": [],
        "result_md": str(pkg_out / "result.md"),
        "pipeline_mode": config.reading_order,
        "image_vlm_model": config.image_vlm_model,
    }

    all_chunks: List[str] = []
    ro_map: Dict[str, Path] = {}
    structure_output_dir: Optional[Path] = None
    if config.reading_order == "xml":
        try:
            ro_map, ro_output = run_structure_analysis_stage(
                repo_root=config.repo_root,
                package_name=pkg_name,
                slide_xmls=slide_xmls,
                strict=config.strict,
            )
            pkg_row["structure_analysis_output_dir"] = str(ro_output)
        except Exception as e:  # noqa: BLE001
            _append_package_stage_failure(
                pkg_row=pkg_row,
                slide_xmls=slide_xmls,
                error_message=f"structure_analysis stage failed: {e}",
                manifest=manifest,
                package_name=pkg_name,
                stage_label="structure_analysis",
            )
            return
    else:
        try:
            structure_output_dir = resolve_surya_structure_dir(surya_structure_root, pkg_name) if surya_structure_root else None
            pkg_row["structure_analysis_output_dir"] = str(structure_output_dir)
        except Exception as e:  # noqa: BLE001
            _append_package_stage_failure(
                pkg_row=pkg_row,
                slide_xmls=slide_xmls,
                error_message=f"surya structure-ready stage failed: {e}",
                manifest=manifest,
                package_name=pkg_name,
                stage_label="surya structure-ready",
            )
            return

    for i, slide_xml in enumerate(slide_xmls, 1):
        slide_started_at = time.perf_counter()
        page_no = parse_slide_number(slide_xml.name, i)
        ordered_slide_xml = ro_map.get(str(slide_xml.resolve()), slide_xml)
        if config.reading_order == "surya" and structure_output_dir is not None:
            ordered_slide_xml = structure_output_dir / f"{slide_xml.stem}.reordered.xml"
        row: Dict[str, object] = {
            "page": page_no,
            "source_xml": str(slide_xml),
            "structure_analysis_xml": str(ordered_slide_xml),
            "status": "ok",
            "warnings": [],
        }
        try:
            if not ordered_slide_xml.exists():
                raise FileNotFoundError(f"reordered slide xml not found: {ordered_slide_xml}")
            merged_md_text, stats = convert_one_slide(
                ordered_slide_xml,
                page_no,
                source_slide_xml=(slide_xml if config.reading_order == "surya" else None),
                output_dir=pkg_out,
                media_dir=media_dir,
                copied_media=copied_media,
                image_vlm_model=config.image_vlm_model,
                image_vlm_prompt=config.image_vlm_prompt,
                image_vlm_max_new_tokens=config.image_vlm_max_new_tokens,
                strict_headings=(config.reading_order == "xml" and config.strict),
            )
            if config.reading_order == "surya":
                row["surya_source"] = str(structure_output_dir)
            if config.per_slide:
                out_md = per_slide_dir / f"{slide_xml.stem}.md"
                out_md.write_text(_rewrite_markdown_for_per_slide(merged_md_text), encoding="utf-8")
                row["output_md"] = str(out_md)
            all_chunks.append(merged_md_text.rstrip())

            row.update(
                {
                    "status": "ok",
                    "elapsed_sec": round(time.perf_counter() - slide_started_at, 3),
                    "blocks_total": stats["blocks_total"],
                    "text_blocks": stats["text_blocks"],
                    "image_blocks": stats["image_blocks"],
                    "image_markdown_blocks": stats["image_markdown_blocks"],
                    "image_markdown_skipped_blocks": stats["image_markdown_skipped_blocks"],
                    "image_markdown_failed_blocks": stats["image_markdown_failed_blocks"],
                    "table_blocks": stats["table_blocks"],
                    "unsupported_blocks": stats["unsupported_blocks"],
                    "skipped_blocks": stats["skipped_blocks"],
                    "rels_path": stats["rels_path"],
                    "warnings": stats["warnings"],
                }
            )
            summary["processed_slides"] = int(summary.get("processed_slides", 0)) + 1
            summary["resolved_images"] = int(summary.get("resolved_images", 0)) + int(stats["resolved_images"])
            summary["unresolved_images"] = int(summary.get("unresolved_images", 0)) + int(stats["unresolved_images"])
            summary["image_markdown_blocks"] = int(summary.get("image_markdown_blocks", 0)) + int(
                stats["image_markdown_blocks"]
            )
            summary["image_markdown_skipped_blocks"] = int(summary.get("image_markdown_skipped_blocks", 0)) + int(
                stats["image_markdown_skipped_blocks"]
            )
            summary["image_markdown_failed_blocks"] = int(summary.get("image_markdown_failed_blocks", 0)) + int(
                stats["image_markdown_failed_blocks"]
            )
            summary["table_blocks"] = int(summary.get("table_blocks", 0)) + int(stats["table_blocks"])
            print(
                f"[{pkg_name}] Processed: {slide_xml.name} "
                f"elapsed={row['elapsed_sec']:.3f}s "
                f"image_md={stats['image_markdown_blocks']} "
                f"image_md_skipped={stats['image_markdown_skipped_blocks']} "
                f"image_md_failed={stats['image_markdown_failed_blocks']} "
                f"tables={stats['table_blocks']}",
                flush=True,
            )
        except Exception as e:  # noqa: BLE001
            row.update({"status": "failed", "error": str(e), "elapsed_sec": round(time.perf_counter() - slide_started_at, 3)})
            summary["failed"] = int(summary.get("failed", 0)) + 1
            print(
                f"[{pkg_name}] Failed: {slide_xml.name} elapsed={row['elapsed_sec']:.3f}s -> {e}",
                flush=True,
            )
        slides = pkg_row.get("slides")
        if isinstance(slides, list):
            slides.append(row)

    merged = "\n\n".join(all_chunks).strip()
    if merged:
        merged += "\n"
    merged_path = pkg_out / "result.md"
    merged_path.write_text(merged, encoding="utf-8")
    pkg_row["elapsed_sec"] = round(time.perf_counter() - package_started_at, 3)
    summary["processed_packages"] = int(summary.get("processed_packages", 0)) + 1
    summary["elapsed_sec"] = round(float(summary.get("elapsed_sec", 0.0)) + float(pkg_row["elapsed_sec"]), 3)
    packages.append(pkg_row)
    print(
        f"[{pkg_name}] Package done: slides={len(slide_xmls)} elapsed={pkg_row['elapsed_sec']:.3f}s",
        flush=True,
    )


def _write_manifest(output_dir: Path, manifest: Dict[str, object]) -> Path:
    manifest["finished_at"] = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
    manifest_path = output_dir / "convert_manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest_path


def main() -> int:
    args = _parse_args()
    config = _build_config(args)
    config.output_dir.mkdir(parents=True, exist_ok=True)
    config.debug_output_dir.mkdir(parents=True, exist_ok=True)

    ensure_imports(config.repo_root)
    prepared_inputs, early_exit_code = _resolve_prepared_inputs(config)
    if early_exit_code is not None:
        return early_exit_code
    if prepared_inputs is None:
        return 0

    packages = _resolve_packages(config, prepared_inputs)
    if not packages:
        return 0

    surya_structure_root = _prepare_surya_context(config, packages)
    manifest = _new_manifest()
    for pkg in packages:
        _convert_package(config, pkg, surya_structure_root, manifest)

    manifest_path = _write_manifest(config.output_dir, manifest)

    print(f"Wrote package outputs under: {config.output_dir.resolve()}")
    print(f"Wrote: {manifest_path.resolve()}")
    print(
        "Summary: "
        f"packages={manifest['summary']['processed_packages']} "
        f"slides={manifest['summary']['processed_slides']} "
        f"failed={manifest['summary']['failed']} "
        f"image_md={manifest['summary']['image_markdown_blocks']} "
        f"image_md_skipped={manifest['summary']['image_markdown_skipped_blocks']} "
        f"image_md_failed={manifest['summary']['image_markdown_failed_blocks']} "
        f"tables={manifest['summary']['table_blocks']} "
        f"images_resolved={manifest['summary']['resolved_images']} "
        f"images_unresolved={manifest['summary']['unresolved_images']} "
        f"elapsed={float(manifest['summary']['elapsed_sec']):.3f}s"
    )
    return 1 if manifest["summary"]["failed"] > 0 else 0


if __name__ == "__main__":
    raise SystemExit(main())
