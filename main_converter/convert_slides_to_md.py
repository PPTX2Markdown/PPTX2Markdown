#!/usr/bin/env python3
"""
Convert extracted PPTX package(s) to markdown.

Usage:
  python convert_slides_to_md.py [ppt_root_dir|file.pptx ...]
  python convert_slides_to_md.py --raw
  python convert_slides_to_md.py --raw [sample1|sample1.pptx|raw_pptx/sample1.pptx ...]

Rules:
  - If no positional args are provided, process all package roots in ./target_pptx.
  - Package root example: ./target_pptx/sample1
  - If a .pptx file is provided, it is extracted automatically into ./target_pptx/<stem>/.
  - If --raw is provided, process all .pptx files in ./raw_pptx.
    If positional args are also provided, only those raw .pptx files are processed.
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
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple
import xml.etree.ElementTree as ET

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

_IMAGE_TABLE_PIPELINE_MODULE: Optional[object] = None
_IMAGE_TABLE_PIPELINE_IMPORT_ERROR: Optional[str] = None
_IMAGE_TABLE_DEBUG_JSON = os.getenv("IMAGE_TABLE_DEBUG_JSON", "").strip().lower() in {"1", "true", "yes", "on"}


def run_structure_analysis_stage(
    repo_root: Path,
    slide_xmls: Sequence[Path],
    strict: bool = False,
) -> Tuple[Dict[str, Path], Path]:
    ro_script = repo_root / "structure_analyzer" / "extract_structure_analysis.py"
    if not ro_script.exists():
        raise FileNotFoundError(f"structure_analyzer script not found: {ro_script}")

    ro_output = Path(tempfile.mkdtemp(prefix="struct_analysis_xml_", dir="/tmp"))
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
    # Support both new and legacy repository layouts.
    candidates = [
        repo_root,
        repo_root / "table_parser",
        repo_root / "pptx_table_parser" / "table_parser",
        repo_root / "pptx_table_parser",
    ]
    for cand in candidates:
        if cand.exists() and cand.is_dir() and str(cand) not in sys.path:
            sys.path.insert(0, str(cand))


def default_target_dirs(cwd: Path) -> List[Path]:
    # Prefer local main_converter/target_pptx. Create it when absent.
    local_target = cwd / "target_pptx"
    if local_target.exists() and not local_target.is_dir():
        raise NotADirectoryError(f"target_pptx path exists but is not a directory: {local_target}")
    local_target.mkdir(parents=True, exist_ok=True)

    out: List[Path] = [local_target]
    parent_target = cwd.parent / "target_pptx"
    if parent_target.exists() and parent_target.is_dir() and parent_target != local_target:
        out.append(parent_target)
    return out


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


def preferred_target_dir(cwd: Path) -> Path:
    cands = default_target_dirs(cwd)
    if cands:
        return cands[0]
    return cwd / "target_pptx"


def raw_pptx_dir(cwd: Path) -> Path:
    return cwd / "raw_pptx"


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
    return pkg_dir.resolve()


def is_ignored_pptx_file(path: Path) -> bool:
    # Skip Office lock/temp files like "~$sample1.pptx".
    return path.name.startswith("~$")


def prepare_package_inputs(
    cwd: Path,
    raw_inputs: Sequence[str],
    force_extract: bool = False,
) -> List[str]:
    if not raw_inputs:
        return list(raw_inputs)

    prepared: List[str] = []
    extraction_root = preferred_target_dir(cwd)
    search_roots = [cwd]
    for root in default_target_dirs(cwd):
        if root not in search_roots:
            search_roots.append(root)

    for item in raw_inputs:
        p = Path(item)
        candidates = [p]
        for root in search_roots:
            candidates.append(root / item)
        picked_file: Optional[Path] = None
        for cand in candidates:
            if cand.exists() and cand.is_file() and cand.suffix.lower() == ".pptx":
                if is_ignored_pptx_file(cand):
                    continue
                picked_file = cand.resolve()
                break
        if picked_file is None:
            prepared.append(item)
            continue
        pkg_dir = extract_pptx_to_target(
            picked_file,
            extraction_root,
            allow_replace_unmanaged=force_extract,
        )
        prepared.append(str(pkg_dir))
    return prepared


def collect_raw_pptx_inputs(cwd: Path) -> List[str]:
    raw_dir = raw_pptx_dir(cwd)
    raw_dir.mkdir(parents=True, exist_ok=True)
    files = sorted(
        [
            path.resolve()
            for path in raw_dir.glob("*.pptx")
            if path.is_file() and not is_ignored_pptx_file(path)
        ],
        key=lambda path: natural_key(path.name),
    )
    return [str(path) for path in files]


def resolve_selected_raw_inputs(cwd: Path, selections: Sequence[str]) -> List[str]:
    raw_dir = raw_pptx_dir(cwd)
    resolved: List[str] = []
    missing: List[str] = []

    for item in selections:
        token = item.strip()
        if not token:
            continue
        as_path = Path(token)
        candidates: List[Path] = [as_path, cwd / as_path, raw_dir / as_path]
        if as_path.suffix.lower() != ".pptx":
            candidates.append(raw_dir / f"{token}.pptx")

        picked: Optional[Path] = None
        for cand in candidates:
            if not cand.exists() or not cand.is_file():
                continue
            if cand.suffix.lower() != ".pptx" or is_ignored_pptx_file(cand):
                continue
            picked = cand.resolve()
            break

        if picked is None:
            missing.append(item)
            continue
        resolved.append(str(picked))

    if missing:
        msg = ", ".join(missing)
        raise FileNotFoundError(
            "Raw selection did not match any .pptx file in ./raw_pptx (or provided path): "
            f"{msg}"
        )
    return resolved


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
    slide_xml: Path,
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
        remapped = (repo_root / "main_converter" / "target_pptx" / pkg_name / "ppt" / "media" / target_name)
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


def extract_shape_text(shape_elem: ET.Element) -> str:
    return render_shape_blocks(extract_shape_blocks(shape_elem))


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


def infer_heading_depth_fallback_strict(text: str, text_block_index: int) -> Optional[int]:
    # Strict mode keeps fallback conservative; rely on structure_analyzer hints first.
    # TODO: Add stricter lexical and position-aware fallback heuristics for xml-only mode.
    return None


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
            path, warn = resolve_image_path(slide_xml, rels_map, rels_path, embed)
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


def inject_table_overlay_links(
    parsed_table: Dict[str, object],
    graphic_frame: ET.Element,
    overlays: Sequence[Dict[str, object]],
    output_dir: Optional[Path],
    media_dir: Optional[Path] = None,
    copied_media: Optional[Dict[str, Path]] = None,
) -> Dict[str, object]:
    if not overlays:
        return parsed_table

    bounds = compute_table_cell_bounds(graphic_frame)
    if bounds is None:
        return parsed_table
    col_bounds, row_bounds = bounds

    rows = parsed_table.get("rows")
    if not isinstance(rows, list):
        return parsed_table

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
        link = overlay_link_text(
            str(overlay.get("path", "")),
            output_dir,
            media_dir=media_dir,
            copied_media=copied_media,
        )
        updated = f"{existing}\n{link}".strip() if existing else link
        cell["text"] = updated

    return parsed_table


def table_overlay_picture_ids(sp_tree: ET.Element) -> set:
    return set()


def convert_table_to_markdown(
    graphic_frame: ET.Element,
    overlays: Optional[Sequence[Dict[str, object]]] = None,
    output_dir: Optional[Path] = None,
    media_dir: Optional[Path] = None,
    copied_media: Optional[Dict[str, Path]] = None,
) -> Tuple[Optional[str], Optional[str]]:
    tbl = graphic_frame.find(".//a:tbl", NS)
    if tbl is None:
        return None, "graphicFrame without a:tbl"

    import parse_table  # type: ignore
    import tableMaker  # type: ignore

    with tempfile.NamedTemporaryFile("wb", suffix=".xml", delete=True) as tmp:
        tmp.write(ET.tostring(tbl, encoding="utf-8"))
        tmp.flush()
        parsed = parse_table.parse_table_xml(Path(tmp.name))
        parsed = inject_table_overlay_links(
            parsed,
            graphic_frame,
            overlays or [],
            output_dir,
            media_dir=media_dir,
            copied_media=copied_media,
        )
        dense = tableMaker._dense_grid_from_parsed_table(parsed, fill_merged=tableMaker.FILL_BOTH)
        md = tableMaker._render_markdown_flat(dense=dense, header_rows=1, use_header_rows=True)
    return md, None


def _load_image_table_pipeline() -> Tuple[Optional[object], Optional[str]]:
    global _IMAGE_TABLE_PIPELINE_MODULE, _IMAGE_TABLE_PIPELINE_IMPORT_ERROR
    if _IMAGE_TABLE_PIPELINE_MODULE is not None:
        return _IMAGE_TABLE_PIPELINE_MODULE, None
    if _IMAGE_TABLE_PIPELINE_IMPORT_ERROR is not None:
        return None, _IMAGE_TABLE_PIPELINE_IMPORT_ERROR

    import_errors: List[str] = []
    for module_name in ("image_pipeline.service", "image_table_pipeline"):
        try:
            _IMAGE_TABLE_PIPELINE_MODULE = importlib.import_module(module_name)
            return _IMAGE_TABLE_PIPELINE_MODULE, None
        except Exception as exc:  # noqa: BLE001
            import_errors.append(f"{module_name}: {type(exc).__name__}: {exc}")

    _IMAGE_TABLE_PIPELINE_IMPORT_ERROR = "; ".join(import_errors)
    return None, _IMAGE_TABLE_PIPELINE_IMPORT_ERROR


def convert_picture_to_table_markdown(
    image_path: str,
) -> Tuple[Optional[str], Optional[str], bool, Optional[Dict[str, object]]]:
    module, import_error = _load_image_table_pipeline()
    if module is None:
        return None, f"image-table pipeline unavailable: {import_error}", True, None
    try:
        result = module.extract_table_markdown_from_image(Path(image_path), header_rows=1)
    except Exception as exc:  # noqa: BLE001
        return (
            None,
            f"image-table pipeline failed on {Path(image_path).name}: {type(exc).__name__}: {exc}",
            False,
            None,
        )

    if not isinstance(result, dict):
        return None, f"image-table pipeline returned invalid payload: {type(result).__name__}", False, None

    _log_image_table_evaluation(image_path, result)

    status = str(result.get("status", "error"))
    if status == "table":
        markdown = result.get("markdown")
        if isinstance(markdown, str) and markdown.strip():
            return markdown, None, False, result
        return None, f"image-table pipeline rendered empty markdown: {Path(image_path).name}", False, result
    if status == "table_skipped":
        return None, None, False, result
    if status == "not_table":
        return None, None, False, result

    error = result.get("error")
    if not isinstance(error, str) or not error.strip():
        error = "unknown image-table pipeline error"
    return None, f"{Path(image_path).name}: {error}", False, result


def _fmt_eval_value(value: object, digits: int = 3) -> str:
    if isinstance(value, (int, float)):
        return f"{float(value):.{digits}f}"
    return "n/a"


def _to_jsonable_copy(value: object) -> object:
    return json.loads(json.dumps(value, ensure_ascii=False))


def _log_image_table_evaluation(image_path: str, result: Dict[str, object]) -> None:
    name = Path(image_path).name
    status = str(result.get("status", "error"))
    predicted = str(result.get("predicted_class", "unknown"))

    classification = result.get("classification")
    if not isinstance(classification, dict):
        classification = {}
    feature_values = classification.get("feature_values")
    if not isinstance(feature_values, dict):
        feature_values = None
    score_breakdown = classification.get("score_breakdown")
    if not isinstance(score_breakdown, dict):
        score_breakdown = None
    decision_flags = classification.get("decision_flags")
    if not isinstance(decision_flags, dict):
        decision_flags = None
    feature_details = classification.get("feature_details")
    if not isinstance(feature_details, dict):
        feature_details = None
    thresholds = classification.get("thresholds")
    if not isinstance(thresholds, dict):
        thresholds = None
    image_size = classification.get("image_size")
    if not isinstance(image_size, dict):
        image_size = None
    visible_image = classification.get("visible_image")
    if not isinstance(visible_image, dict):
        visible_image = None

    score = result.get("score", classification.get("score", score_breakdown.get("final_score") if isinstance(score_breakdown, dict) else None))
    table_count = result.get("table_count")
    reason = result.get("reason")
    error = result.get("error")
    surya_attempted = bool(result.get("surya_attempted"))
    surya_quality_ok = result.get("surya_quality_ok")

    geometry_score = score_breakdown.get("geometry_score") if isinstance(score_breakdown, dict) else None
    structure_score = score_breakdown.get("structure_score") if isinstance(score_breakdown, dict) else None
    area_score = score_breakdown.get("area_score") if isinstance(score_breakdown, dict) else None
    iou_score = score_breakdown.get("iou_score") if isinstance(score_breakdown, dict) else None
    center_score = score_breakdown.get("center_score") if isinstance(score_breakdown, dict) else None
    box_count_score = score_breakdown.get("box_count_score") if isinstance(score_breakdown, dict) else None

    summary = (
        f"[image-table] {name} "
        f"status={status} "
        f"pred={predicted} "
        f"score={_fmt_eval_value(score)} "
        f"det(geom={_fmt_eval_value(geometry_score)},"
        f"struct={_fmt_eval_value(structure_score)},"
        f"area={_fmt_eval_value(area_score)},"
        f"iou={_fmt_eval_value(iou_score)},"
        f"center={_fmt_eval_value(center_score)},"
        f"boxes={_fmt_eval_value(box_count_score)}) "
        f"bbox(area_ratio={_fmt_eval_value(feature_values.get('table_union_area_ratio') if isinstance(feature_values, dict) else None)},"
        f"bbox_ratio={_fmt_eval_value(feature_values.get('table_union_bbox_area_ratio') if isinstance(feature_values, dict) else None)},"
        f"iou={_fmt_eval_value(feature_values.get('table_union_bbox_iou') if isinstance(feature_values, dict) else None)},"
        f"offset={_fmt_eval_value(feature_values.get('table_union_center_offset') if isinstance(feature_values, dict) else None)}) "
        f"count(tbl={_fmt_eval_value(feature_values.get('detected_table_count') if isinstance(feature_values, dict) else None, digits=0)},"
        f"rows={_fmt_eval_value(feature_values.get('detected_row_count') if isinstance(feature_values, dict) else None, digits=0)},"
        f"cols={_fmt_eval_value(feature_values.get('detected_col_count') if isinstance(feature_values, dict) else None, digits=0)},"
        f"cells={_fmt_eval_value(feature_values.get('detected_cell_count') if isinstance(feature_values, dict) else None, digits=0)},"
        f"box={_fmt_eval_value(feature_values.get('detected_box_count') if isinstance(feature_values, dict) else None, digits=0)})"
    )

    extras: List[str] = []
    extras.append(f"surya={'run' if surya_attempted else 'skip'}")
    if isinstance(table_count, int):
        extras.append(f"tables={table_count}")
    if isinstance(surya_quality_ok, bool):
        extras.append(f"surya_quality={'ok' if surya_quality_ok else 'low'}")
    if isinstance(reason, str) and reason.strip():
        extras.append(f"reason={reason}")
    if isinstance(error, str) and error.strip():
        extras.append(f"error={error}")
    active_flags = [name for name, active in decision_flags.items() if active] if isinstance(decision_flags, dict) else []
    if active_flags:
        extras.append("flags=" + ",".join(sorted(active_flags)))
    if extras:
        summary = f"{summary} " + " ".join(extras)

    print(summary)
    if _IMAGE_TABLE_DEBUG_JSON:
        feature_details_payload = {
            "visible_image_bbox": feature_details.get("visible_image_bbox") if isinstance(feature_details, dict) else None,
            "detected_table_union_bbox": (
                feature_details.get("detected_table_union_bbox") if isinstance(feature_details, dict) else None
            ),
            "detected_tables": feature_details.get("detected_tables") if isinstance(feature_details, dict) else None,
        }
        if isinstance(feature_details, dict):
            feature_details_payload.update(feature_details)

        debug_payload = {
            "kind": "image_table_debug",
            "file": str(Path(str(result.get("file", image_path))).expanduser().resolve()),
            "image_name": name,
            "status": status,
            "predicted_class": result.get("predicted_class"),
            "score": score,
            "low_confidence": result.get("low_confidence"),
            "surya_attempted": surya_attempted,
            "surya_quality_ok": surya_quality_ok if isinstance(surya_quality_ok, bool) else None,
            "table_count": table_count if isinstance(table_count, int) else None,
            "reason": reason if isinstance(reason, str) and reason.strip() else None,
            "error": error if isinstance(error, str) and error.strip() else None,
            "thresholds": thresholds,
            "image_size": image_size,
            "visible_image": visible_image,
            "feature_values": feature_values,
            "score_breakdown": score_breakdown,
            "decision_flags": decision_flags,
            "feature_details": feature_details_payload,
        }
        print(json.dumps(_to_jsonable_copy(debug_payload), ensure_ascii=False, separators=(",", ":")))


def convert_one_slide(
    slide_xml: Path,
    page_no: int,
    source_slide_xml: Optional[Path] = None,
    output_dir: Optional[Path] = None,
    media_dir: Optional[Path] = None,
    copied_media: Optional[Dict[str, Path]] = None,
    surya_debug_dir: Optional[Path] = None,
    copied_surya_debug_images: Optional[Dict[str, Path]] = None,
    enable_image_table_pipeline: bool = False,
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
        "table_blocks": 0,
        "table_skipped_blocks": 0,
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
            img_path, warn = resolve_image_path(slide_xml, rels_map, rels_path, embed)
            if warn:
                stats["unresolved_images"] += 1
                stats["warnings"].append(warn)
            else:
                stats["resolved_images"] += 1

            if enable_image_table_pipeline and not warn and not img_path.startswith("[unresolved-image"):
                table_md, table_warn, unavailable, table_result = convert_picture_to_table_markdown(img_path)
                if isinstance(table_result, dict) and (
                    bool(table_result.get("surya_attempted"))
                    or str(table_result.get("status", "")) == "table_skipped"
                ):
                    copy_debug_image_asset(
                        img_path,
                        debug_dir=surya_debug_dir,
                        copied_debug_images=copied_surya_debug_images,
                    )
                if table_md is not None:
                    lines.append(table_md.strip())
                    lines.append("")
                    stats["table_blocks"] += 1
                    continue
                if isinstance(table_result, dict) and str(table_result.get("status", "")) == "table_skipped":
                    stats["table_skipped_blocks"] += 1
                if table_warn:
                    if unavailable:
                        if not image_pipeline_unavailable_reported:
                            stats["warnings"].append(table_warn)
                            image_pipeline_unavailable_reported = True
                    else:
                        stats["warnings"].append(table_warn)

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
            table_md, err = convert_table_to_markdown(
                child,
                overlays=table_overlay_map.get(shape_id_of(child), []),
                output_dir=output_dir,
                media_dir=media_dir,
                copied_media=copied_media,
            )
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
    manifest_here = surya_root / "structure_analysis_manifest.json"
    if manifest_here.exists():
        return surya_root
    candidate = surya_root / package_name
    manifest_there = candidate / "structure_analysis_manifest.json"
    if manifest_there.exists():
        return candidate
    raise FileNotFoundError(
        f"surya structure-ready output not found for package '{package_name}' under {surya_root}"
    )


def run_surya_pipeline_stage(
    repo_root: Path,
    surya_root: Path,
    force: bool = False,
) -> Path:
    run_script = repo_root / "surya_pipeline" / "run_surya_pipeline.py"
    if not run_script.exists():
        raise FileNotFoundError(f"surya pipeline script not found: {run_script}")

    cmd = [
        sys.executable,
        str(run_script),
        "--surya-dir",
        str(surya_root),
    ]
    if force:
        cmd.append("--force")

    print(f"[surya] Running pipeline: {' '.join(cmd)}")
    proc = subprocess.run(cmd, text=True, cwd=str(repo_root))
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
    repo_root: Path,
    raw_surya_dir: Optional[Path],
    force: bool = False,
    use_existing_output: bool = False,
) -> Path:
    if raw_surya_dir is None:
        candidate = repo_root / "surya_pipeline"
    else:
        candidate = raw_surya_dir

    if use_existing_output:
        # Structure-ready dir passed directly.
        if (candidate / "structure_analysis_manifest.json").exists():
            return candidate
        # Package subdirs under structure_ready root.
        if candidate.name == "structure_ready" and candidate.exists() and candidate.is_dir():
            return candidate
        structure_root = candidate / "output" / "structure_ready"
        if structure_root.exists() and structure_root.is_dir():
            return structure_root

    # Surya pipeline root passed in or inferred.
    if (candidate / "run_surya_pipeline.py").exists():
        return run_surya_pipeline_stage(repo_root=repo_root, surya_root=candidate, force=force)

    if use_existing_output:
        structure_root = candidate / "output" / "structure_ready"
        if structure_root.exists() and structure_root.is_dir():
            return structure_root

    raise FileNotFoundError(
        "valid surya input not found. Pass surya_pipeline root, or use --use-existing-surya-output "
        "with output/structure_ready: "
        f"{candidate}"
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Convert extracted PPTX package(s) to markdown."
    )
    parser.add_argument(
        "inputs",
        nargs="*",
        help=(
            "Package root dir(s) or .pptx file(s). "
            "If --raw is used, these are treated as raw selections."
        ),
    )
    parser.add_argument(
        "--raw",
        action="store_true",
        help=(
            "Process .pptx files in ./raw_pptx by extracting into ./target_pptx first. "
            "With positional args, process only selected raw files."
        ),
    )
    parser.add_argument(
        "--per-slide",
        action="store_true",
        help="Also write per-slide markdown files under ./output/.../<package>/per_slide.",
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
        "--surya-dir",
        default=None,
        help="Path to surya_pipeline root. If omitted, defaults to ../surya_pipeline.",
    )
    parser.add_argument(
        "--force-surya-pipeline",
        action="store_true",
        help="Deprecated alias. Surya mode now re-runs the pipeline by default.",
    )
    parser.add_argument(
        "--reuse-surya-cache",
        action="store_true",
        help="Reuse existing Surya outputs instead of re-running the pipeline.",
    )
    parser.add_argument(
        "--use-existing-surya-output",
        action="store_true",
        help="Use existing output/structure_ready instead of running the Surya pipeline.",
    )
    parser.add_argument(
        "--image-table-pipeline",
        action="store_true",
        help="Classify image blocks with Surya bbox/box-count signals and parse detected table images with Surya.",
    )
    parser.add_argument(
        "--output-file",
        default=None,
        help="Optional single markdown output path merged from processed package result(s).",
    )
    args = parser.parse_args()
    cwd = Path.cwd()
    output_root = cwd / "output"
    output_dir = output_root / args.reading_order
    output_dir.mkdir(parents=True, exist_ok=True)

    repo_root = Path(__file__).resolve().parent.parent
    debug_output_dir = repo_root / "main_converter" / "output" / args.reading_order
    debug_output_dir.mkdir(parents=True, exist_ok=True)
    ensure_imports(repo_root)
    if args.raw:
        raw_inputs = (
            resolve_selected_raw_inputs(cwd, args.inputs) if args.inputs else collect_raw_pptx_inputs(cwd)
        )
        if not raw_inputs:
            print(f"No .pptx files found in: {raw_pptx_dir(cwd).resolve()}")
            return 0
        prepared_inputs = prepare_package_inputs(cwd, raw_inputs, force_extract=True)
    else:
        prepared_inputs = prepare_package_inputs(cwd, args.inputs)
    surya_dir = Path(args.surya_dir).resolve() if args.surya_dir else None
    surya_structure_root: Optional[Path] = None
    if args.reading_order == "surya":
        if args.strict:
            print("[info] --strict is ignored for surya mode. surya heading logic remains separate.")
        surya_structure_root = prepare_surya_structure_root(
            repo_root=repo_root,
            raw_surya_dir=surya_dir,
            force=(not args.reuse_surya_cache) or args.force_surya_pipeline,
            use_existing_output=args.use_existing_surya_output,
        )

    target_dirs = default_target_dirs(cwd)
    packages = pick_packages(target_dirs, prepared_inputs)
    if not packages:
        print("No valid PPTX package directories found.")
        if target_dirs:
            print("Checked default directories:")
            for d in target_dirs:
                print(f"- {d.resolve()}")
        else:
            print("Checked default directories: none found (expected ./target_pptx).")
        return 0

    manifest = {
        "started_at": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "packages": [],
        "summary": {
            "processed_packages": 0,
            "processed_slides": 0,
            "failed": 0,
            "resolved_images": 0,
            "unresolved_images": 0,
            "table_blocks": 0,
            "table_skipped_blocks": 0,
        },
    }
    merged_packages: List[Tuple[str, str]] = []

    for pkg in packages:
        pkg_name = pkg.name
        slides_dir = pkg / "ppt" / "slides"
        slide_xmls = sorted(
            [p for p in slides_dir.glob("slide*.xml") if p.is_file()],
            key=lambda p: natural_key(p.name),
        )
        pkg_out = output_dir / pkg_name
        per_slide_dir = pkg_out / "per_slide"
        media_dir = pkg_out / "media"
        surya_debug_dir = debug_output_dir / pkg_name / "surya_run_images"
        copied_media: Dict[str, Path] = {}
        copied_surya_debug_images: Dict[str, Path] = {}
        pkg_out.mkdir(parents=True, exist_ok=True)
        if args.per_slide:
            per_slide_dir.mkdir(parents=True, exist_ok=True)
        elif per_slide_dir.exists():
            shutil.rmtree(per_slide_dir)

        pkg_row = {
            "package": str(pkg),
            "name": pkg_name,
            "slides": [],
            "result_md": str(pkg_out / "result.md"),
            "pipeline_mode": args.reading_order,
        }

        all_chunks: List[str] = []
        ro_map: Dict[str, Path] = {}
        structure_output_dir: Optional[Path] = None
        if args.reading_order == "xml":
            try:
                ro_map, ro_output = run_structure_analysis_stage(
                    repo_root=repo_root,
                    slide_xmls=slide_xmls,
                    strict=args.strict,
                )
                pkg_row["structure_analysis_output_dir"] = str(ro_output)
            except Exception as e:  # noqa: BLE001
                for slide_xml in slide_xmls:
                    row = {
                        "page": parse_slide_number(slide_xml.name, 0),
                        "source_xml": str(slide_xml),
                        "status": "failed",
                        "error": f"structure_analysis stage failed: {e}",
                        "warnings": [],
                    }
                    pkg_row["slides"].append(row)
                    manifest["summary"]["failed"] += 1
                manifest["packages"].append(pkg_row)
                print(f"[{pkg_name}] structure_analysis failed: {e}")
                continue
        else:
            try:
                structure_output_dir = (
                    resolve_surya_structure_dir(surya_structure_root, pkg_name) if surya_structure_root else None
                )
                pkg_row["structure_analysis_output_dir"] = str(structure_output_dir)
            except Exception as e:  # noqa: BLE001
                for slide_xml in slide_xmls:
                    row = {
                        "page": parse_slide_number(slide_xml.name, 0),
                        "source_xml": str(slide_xml),
                        "status": "failed",
                        "error": f"surya structure-ready stage failed: {e}",
                        "warnings": [],
                    }
                    pkg_row["slides"].append(row)
                    manifest["summary"]["failed"] += 1
                manifest["packages"].append(pkg_row)
                print(f"[{pkg_name}] surya structure-ready failed: {e}")
                continue

        for i, slide_xml in enumerate(slide_xmls, 1):
            page_no = parse_slide_number(slide_xml.name, i)
            ordered_slide_xml = ro_map.get(str(slide_xml.resolve()), slide_xml)
            if args.reading_order == "surya" and structure_output_dir is not None:
                ordered_slide_xml = structure_output_dir / f"{slide_xml.stem}.reordered.xml"
            row = {
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
                    source_slide_xml=(slide_xml if args.reading_order == "surya" else None),
                    output_dir=pkg_out,
                    media_dir=media_dir,
                    copied_media=copied_media,
                    surya_debug_dir=surya_debug_dir,
                    copied_surya_debug_images=copied_surya_debug_images,
                    enable_image_table_pipeline=args.image_table_pipeline,
                    strict_headings=(args.reading_order == "xml" and args.strict),
                )
                if args.reading_order == "surya":
                    row["surya_source"] = str(structure_output_dir)
                if args.per_slide:
                    md_text, _ = convert_one_slide(
                        ordered_slide_xml,
                        page_no,
                        source_slide_xml=(slide_xml if args.reading_order == "surya" else None),
                        output_dir=per_slide_dir,
                        media_dir=media_dir,
                        copied_media=copied_media,
                        surya_debug_dir=surya_debug_dir,
                        copied_surya_debug_images=copied_surya_debug_images,
                        enable_image_table_pipeline=args.image_table_pipeline,
                        strict_headings=(args.reading_order == "xml" and args.strict),
                    )
                    out_md = per_slide_dir / f"{slide_xml.stem}.md"
                    out_md.write_text(md_text, encoding="utf-8")
                all_chunks.append(merged_md_text.rstrip())

                row.update(
                    {
                        "status": "ok",
                        "blocks_total": stats["blocks_total"],
                        "text_blocks": stats["text_blocks"],
                        "image_blocks": stats["image_blocks"],
                        "table_blocks": stats["table_blocks"],
                        "table_skipped_blocks": stats["table_skipped_blocks"],
                        "unsupported_blocks": stats["unsupported_blocks"],
                        "skipped_blocks": stats["skipped_blocks"],
                        "rels_path": stats["rels_path"],
                        "warnings": stats["warnings"],
                    }
                )
                if args.per_slide:
                    row["output_md"] = str(out_md)
                manifest["summary"]["processed_slides"] += 1
                manifest["summary"]["resolved_images"] += stats["resolved_images"]
                manifest["summary"]["unresolved_images"] += stats["unresolved_images"]
                manifest["summary"]["table_blocks"] += stats["table_blocks"]
                manifest["summary"]["table_skipped_blocks"] += stats["table_skipped_blocks"]
                print(f"[{pkg_name}] Processed: {slide_xml.name}")
            except Exception as e:  # noqa: BLE001
                row.update({"status": "failed", "error": str(e)})
                manifest["summary"]["failed"] += 1
                print(f"[{pkg_name}] Failed: {slide_xml.name} -> {e}")
            pkg_row["slides"].append(row)

        merged = "\n\n".join(all_chunks).strip()
        if merged:
            merged += "\n"
        merged_path = pkg_out / "result.md"
        merged_path.write_text(merged, encoding="utf-8")
        merged_packages.append((pkg_name, merged))
        manifest["summary"]["processed_packages"] += 1
        manifest["packages"].append(pkg_row)

    manifest["finished_at"] = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
    manifest_path = output_dir / "convert_manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    if args.output_file:
        out_path = Path(args.output_file).expanduser()
        if not out_path.is_absolute():
            out_path = (cwd / out_path).resolve()
        out_path.parent.mkdir(parents=True, exist_ok=True)

        if len(merged_packages) <= 1:
            merged_text = merged_packages[0][1] if merged_packages else ""
        else:
            blocks: List[str] = []
            for pkg_name, pkg_md in merged_packages:
                block = f"## Package: {pkg_name}\n\n{pkg_md.strip()}".strip()
                if block:
                    blocks.append(block)
            merged_text = "\n\n".join(blocks)
            if merged_text:
                merged_text += "\n"

        out_path.write_text(merged_text, encoding="utf-8")
        print(f"Wrote combined markdown: {out_path.resolve()}")

    print(f"Wrote package outputs under: {output_dir.resolve()}")
    print(f"Wrote: {manifest_path.resolve()}")
    print(
        "Summary: "
        f"packages={manifest['summary']['processed_packages']} "
        f"slides={manifest['summary']['processed_slides']} "
        f"failed={manifest['summary']['failed']} "
        f"tables={manifest['summary']['table_blocks']} "
        f"table_skipped={manifest['summary']['table_skipped_blocks']} "
        f"images_resolved={manifest['summary']['resolved_images']} "
        f"images_unresolved={manifest['summary']['unresolved_images']}"
    )
    return 1 if manifest["summary"]["failed"] > 0 else 0


if __name__ == "__main__":
    raise SystemExit(main())
