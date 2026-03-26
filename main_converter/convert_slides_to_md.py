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
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import time
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple
import xml.etree.ElementTree as ET

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from asset_utils import copy_debug_image_asset, copy_media_asset
from converter_models import ConversionManifest, ConverterConfig, SlideStats
from reading_order_pipeline import (
    prepare_surya_structure_root,
    resolve_surya_structure_dir,
    run_structure_analysis_stage,
)
from heading_rules import normalize_single_heading_to_h1
from image_pipeline.service import (
    DEFAULT_GEMINI_API_KEY_ENV,
    DEFAULT_GEMINI_MODEL,
    DEFAULT_MAX_NEW_TOKENS as DEFAULT_IMAGE_VLM_MAX_NEW_TOKENS,
    DEFAULT_PROMPT as DEFAULT_IMAGE_VLM_PROMPT,
    DEFAULT_PROVIDER as DEFAULT_IMAGE_VLM_PROVIDER,
    extract_markdown_from_image,
    normalize_provider,
)
from image_table_pipeline_adapter import convert_picture_to_table_markdown
from slide_converter import SlideConversionDeps, convert_one_slide as convert_one_slide_core
from table_overlay import (
    collect_table_overlay_pictures as collect_table_overlay_pictures_core,
    convert_table_to_markdown as convert_table_to_markdown_core,
    shape_id_of as shape_id_of_core,
)


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
REL_NS = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}

logger = logging.getLogger(__name__)


def _configure_logging(verbose: bool = False) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(level=level, format="%(message)s")


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


def convert_picture_to_markdown(
    image_path: str,
    *,
    provider: str = DEFAULT_IMAGE_VLM_PROVIDER,
    model_spec: Optional[str] = None,
    prompt: str = DEFAULT_IMAGE_VLM_PROMPT,
    max_new_tokens: int = DEFAULT_IMAGE_VLM_MAX_NEW_TOKENS,
    gemini_api_key_env: str = DEFAULT_GEMINI_API_KEY_ENV,
) -> Tuple[Optional[str], Optional[str], bool, Optional[Dict[str, object]]]:
    def is_pipeline_unavailable(message: str) -> bool:
        normalized = str(message or "").strip().lower()
        return (
            "api key not found" in normalized
            or "modulenotfounderror" in normalized
            or "importerror" in normalized
        )

    normalized_provider = normalize_provider(provider)
    effective_model = model_spec
    if normalized_provider == "gemini" and not effective_model:
        effective_model = DEFAULT_GEMINI_MODEL
    if not effective_model:
        return None, None, False, None

    try:
        result = extract_markdown_from_image(
            Path(image_path),
            provider=normalized_provider,
            model_spec=effective_model,
            prompt=prompt,
            max_new_tokens=max(1, int(max_new_tokens)),
            gemini_api_key_env=gemini_api_key_env,
        )
    except Exception as exc:  # noqa: BLE001
        error_message = f"image pipeline failed on {Path(image_path).name}: {type(exc).__name__}: {exc}"
        return (
            None,
            error_message,
            is_pipeline_unavailable(error_message),
            None,
        )

    if not isinstance(result, dict):
        return None, f"image pipeline returned invalid payload: {type(result).__name__}", False, None

    status = str(result.get("status", "")).strip().lower()
    if status == "markdown":
        markdown = str(result.get("markdown", "")).strip()
        if markdown:
            return markdown, None, False, result
        return None, f"image pipeline rendered empty markdown: {Path(image_path).name}", False, result
    if status == "no_markdown":
        return None, None, False, result

    error = str(result.get("error", "")).strip() or "unknown image pipeline error"
    return None, f"{Path(image_path).name}: {error}", is_pipeline_unavailable(error), result


def format_markdown_image(
    path: str,
    output_dir: Optional[Path],
    alt_text: str = "image",
    media_dir: Optional[Path] = None,
    copied_media: Optional[Dict[str, Path]] = None,
    image_vlm_provider: str = DEFAULT_IMAGE_VLM_PROVIDER,
    image_vlm_model: Optional[str] = None,
    image_vlm_prompt: str = DEFAULT_IMAGE_VLM_PROMPT,
    image_vlm_max_new_tokens: int = DEFAULT_IMAGE_VLM_MAX_NEW_TOKENS,
    image_vlm_api_key_env: str = DEFAULT_GEMINI_API_KEY_ENV,
) -> Tuple[str, Optional[str], bool, bool, bool]:
    if path.startswith("[unresolved-image"):
        return path, None, False, False, False

    normalized_provider = normalize_provider(image_vlm_provider)
    image_vlm_enabled = bool(image_vlm_model) or normalized_provider == "gemini"
    if image_vlm_enabled:
        image_md, image_warn, unavailable, result = convert_picture_to_markdown(
            path,
            provider=normalized_provider,
            model_spec=image_vlm_model,
            prompt=image_vlm_prompt,
            max_new_tokens=image_vlm_max_new_tokens,
            gemini_api_key_env=image_vlm_api_key_env,
        )
        if image_md is not None:
            return annotate_generated_image_markdown(image_md, path), None, unavailable, True, False
        skipped_no_markdown = isinstance(result, dict) and str(result.get("status", "")) == "no_markdown"
        copied_path = copy_media_asset(path, media_dir=media_dir, copied_media=copied_media)
        relative_path = relativize_markdown_path(copied_path, output_dir)
        return f"![{alt_text}]({relative_path})", image_warn, unavailable, False, skipped_no_markdown

    copied_path = copy_media_asset(path, media_dir=media_dir, copied_media=copied_media)
    relative_path = relativize_markdown_path(copied_path, output_dir)
    return f"![{alt_text}]({relative_path})", None, False, False, False


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
    return shape_id_of_core(elem, NS)


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
) -> Tuple[Dict[str, List[Dict[str, object]]], set[str], List[str], int, int]:
    _ = slide_xml
    return collect_table_overlay_pictures_core(
        sp_tree,
        rels_map,
        rels_path,
        ns=NS,
        local_name_fn=local_name,
        resolve_image_path_fn=resolve_image_path,
    )


def convert_table_to_markdown(
    graphic_frame: ET.Element,
    overlays: Optional[Sequence[Dict[str, object]]] = None,
    output_dir: Optional[Path] = None,
    media_dir: Optional[Path] = None,
    copied_media: Optional[Dict[str, Path]] = None,
) -> Tuple[Optional[str], Optional[str]]:
    return convert_table_to_markdown_core(
        graphic_frame,
        overlays=overlays,
        output_dir=output_dir,
        media_dir=media_dir,
        copied_media=copied_media,
        ns=NS,
        normalize_text_fn=normalize_text,
        overlay_link_text_fn=overlay_link_text,
    )


def _slide_conversion_deps() -> SlideConversionDeps:
    return SlideConversionDeps(
        local_name=local_name,
        choose_rels_in_package=choose_rels_in_package,
        build_rels_map=build_rels_map,
        load_heading_hints=load_heading_hints,
        collect_table_overlay_pictures=collect_table_overlay_pictures,
        extract_shape_blocks=extract_shape_blocks,
        render_shape_blocks=render_shape_blocks,
        split_triangle_bullets=split_triangle_bullets,
        normalize_text=normalize_text,
        shape_id_of=shape_id_of,
        resolve_image_path=resolve_image_path,
        convert_picture_to_table_markdown=convert_picture_to_table_markdown,
        copy_debug_image_asset=copy_debug_image_asset,
        format_markdown_image=format_markdown_image,
        convert_table_to_markdown=convert_table_to_markdown,
        graphic_frame_kind=graphic_frame_kind,
        diagram_data_path=diagram_data_path,
        extract_diagram_texts=extract_diagram_texts,
        format_diagram_as_markdown=format_diagram_as_markdown,
        normalize_single_heading_to_h1=normalize_single_heading_to_h1,
    )


def convert_one_slide(
    slide_xml: Path,
    page_no: int,
    source_slide_xml: Optional[Path] = None,
    output_dir: Optional[Path] = None,
    media_dir: Optional[Path] = None,
    copied_media: Optional[Dict[str, Path]] = None,
    image_vlm_provider: str = DEFAULT_IMAGE_VLM_PROVIDER,
    image_vlm_model: Optional[str] = None,
    image_vlm_prompt: str = DEFAULT_IMAGE_VLM_PROMPT,
    image_vlm_max_new_tokens: int = DEFAULT_IMAGE_VLM_MAX_NEW_TOKENS,
    image_vlm_api_key_env: str = DEFAULT_GEMINI_API_KEY_ENV,
    strict_headings: bool = False,
) -> Tuple[str, SlideStats]:
    return convert_one_slide_core(
        slide_xml,
        page_no,
        source_slide_xml=source_slide_xml,
        output_dir=output_dir,
        media_dir=media_dir,
        copied_media=copied_media,
        image_vlm_provider=image_vlm_provider,
        image_vlm_model=image_vlm_model,
        image_vlm_prompt=image_vlm_prompt,
        image_vlm_max_new_tokens=image_vlm_max_new_tokens,
        image_vlm_api_key_env=image_vlm_api_key_env,
        surya_debug_dir=None,
        copied_surya_debug_images=None,
        enable_image_table_pipeline=False,
        strict_headings=strict_headings,
        deps=_slide_conversion_deps(),
        ns=NS,
    )


def _parse_args() -> argparse.Namespace:
    # parser 생성 
    parser = argparse.ArgumentParser(
        description="Convert extracted PPTX package(s) to markdown."
    )
    # 필요한 인자들 추가
    parser.add_argument(
        "inputs",
        nargs="*",
        help=(
            "Optional .pptx selections (e.g., sample3.pptx sample4.pptx). "
            "If omitted, all .pptx files under main_converter/target_pptx are extracted/processed."
        ),
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
        "--image-vlm-provider",
        choices=("local", "gemini"),
        default=DEFAULT_IMAGE_VLM_PROVIDER,
        help="Image VLM backend. local uses Qwen2.5-VL, gemini uses the Gemini API.",
    )
    parser.add_argument(
        "--image-vlm-model",
        help=(
            "Image VLM model identifier. "
            "Use 3b/7b (or a Hugging Face model id) for --image-vlm-provider local, "
            "or a Gemini model id such as gemini-2.5-flash for --image-vlm-provider gemini."
        ),
    )
    parser.add_argument(
        "--image-vlm-prompt",
        default=DEFAULT_IMAGE_VLM_PROMPT,
        help="Prompt passed to the image VLM when image conversion is enabled.",
    )
    parser.add_argument(
        "--image-vlm-max-new-tokens",
        type=int,
        default=DEFAULT_IMAGE_VLM_MAX_NEW_TOKENS,
        help="Maximum number of tokens to generate per image when image conversion is enabled.",
    )
    parser.add_argument(
        "--image-vlm-api-key-env",
        default=DEFAULT_GEMINI_API_KEY_ENV,
        help="Environment variable name containing the Gemini API key when --image-vlm-provider gemini is used.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable verbose debug logging.",
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
        reading_order=str(args.reading_order),
        strict=bool(args.strict),
        reuse_surya_cache=bool(args.reuse_surya_cache),
        image_vlm_provider=normalize_provider(args.image_vlm_provider),
        image_vlm_model=(str(args.image_vlm_model).strip() if args.image_vlm_model else None),
        image_vlm_prompt=str(args.image_vlm_prompt),
        image_vlm_max_new_tokens=max(1, int(args.image_vlm_max_new_tokens)),
        image_vlm_api_key_env=str(args.image_vlm_api_key_env).strip() or DEFAULT_GEMINI_API_KEY_ENV,
    )


def _resolve_prepared_inputs(config: ConverterConfig) -> List[str]:
    if config.inputs:
        non_pptx_inputs = [x for x in config.inputs if Path(x).suffix.lower() != ".pptx"]
        if non_pptx_inputs:
            logger.error("Only .pptx inputs are allowed.")
            logger.error("Provide files like: sample1.pptx sample2.pptx")
            for item in non_pptx_inputs:
                logger.error("- %s", item)
            raise ValueError("invalid non-pptx inputs")
        return prepare_package_inputs(config.cwd, config.inputs, force_extract=True)

    auto_pptx_inputs = collect_target_pptx_inputs(config.cwd)
    if not auto_pptx_inputs:
        logger.info("No .pptx files found in: %s", preferred_pptx_input_dir(config.cwd).resolve())
        return []
    return prepare_package_inputs(config.cwd, auto_pptx_inputs, force_extract=True)


def _resolve_packages(config: ConverterConfig, prepared_inputs: Sequence[str]) -> List[Path]:
    target_dirs = default_target_dirs(config.cwd)
    packages = pick_packages(target_dirs, prepared_inputs)
    if not packages:
        logger.info("No valid PPTX package directories found.")
        if target_dirs:
            logger.info("Checked default directories:")
            for d in target_dirs:
                logger.info("- %s", d.resolve())
        else:
            logger.info("Checked default directories: none found (expected ./target_slides).")
    return packages


def _prepare_surya_context(config: ConverterConfig, packages: Sequence[Path]) -> Optional[Path]:
    if config.reading_order != "surya":
        return None
    if config.strict:
        logger.info("[info] --strict is ignored for surya mode. surya heading logic remains separate.")
    shared_target_slides_dir = preferred_target_dir(config.cwd).resolve()
    shared_target_pptx_dir = preferred_pptx_input_dir(config.cwd).resolve()
    return prepare_surya_structure_root(
        force=not config.reuse_surya_cache,
        reuse_existing_output=config.reuse_surya_cache,
        targets=[pkg.name for pkg in packages],
        target_pptx_dir=shared_target_pptx_dir,
        target_slides_dir=shared_target_slides_dir,
    )


def _new_manifest() -> ConversionManifest:
    return ConversionManifest()


def _append_package_stage_failure(
    pkg_row: Dict[str, object],
    slide_xmls: Sequence[Path],
    error_message: str,
    manifest: ConversionManifest,
    package_name: str,
    stage_label: str,
) -> None:
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
        manifest.summary.failed += 1
    manifest.packages.append(pkg_row)
    logger.error("[%s] %s failed: %s", package_name, stage_label, error_message)


def _convert_package(
    config: ConverterConfig,
    pkg: Path,
    surya_structure_root: Optional[Path],
    manifest: ConversionManifest,
) -> None:
    pkg_name = pkg.name
    slides_dir = pkg / "ppt" / "slides"
    slide_xmls = sorted(
        [p for p in slides_dir.glob("slide*.xml") if p.is_file()],
        key=lambda p: natural_key(p.name),
    )
    pkg_out = config.output_dir / pkg_name
    media_dir = pkg_out / "media"
    copied_media: Dict[str, Path] = {}
    pkg_out.mkdir(parents=True, exist_ok=True)

    pkg_row: Dict[str, object] = {
        "package": str(pkg),
        "name": pkg_name,
        "slides": [],
        "result_md": str(pkg_out / "result.md"),
        "pipeline_mode": config.reading_order,
        "image_vlm_provider": config.image_vlm_provider,
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
                image_vlm_provider=config.image_vlm_provider,
                image_vlm_model=config.image_vlm_model,
                image_vlm_prompt=config.image_vlm_prompt,
                image_vlm_max_new_tokens=config.image_vlm_max_new_tokens,
                image_vlm_api_key_env=config.image_vlm_api_key_env,
                strict_headings=(config.reading_order == "xml" and config.strict),
            )
            if config.reading_order == "surya":
                row["surya_source"] = str(structure_output_dir)
            all_chunks.append(merged_md_text.rstrip())

            row.update(
                {"status": "ok", **stats.to_slide_row_fields()}
            )
            manifest.summary.add_slide(stats)
            logger.info("[%s] Processed: %s", pkg_name, slide_xml.name)
        except Exception as e:  # noqa: BLE001
            row.update({"status": "failed", "error": str(e)})
            manifest.summary.failed += 1
            logger.error("[%s] Failed: %s -> %s", pkg_name, slide_xml.name, e)
        slides = pkg_row.get("slides")
        if isinstance(slides, list):
            slides.append(row)

    merged = "\n\n".join(all_chunks).strip()
    if merged:
        merged += "\n"
    merged_path = pkg_out / "result.md"
    merged_path.write_text(merged, encoding="utf-8")
    manifest.summary.processed_packages += 1
    manifest.packages.append(pkg_row)


def _write_manifest(output_dir: Path, manifest: ConversionManifest) -> Path:
    manifest.mark_finished()
    manifest_path = output_dir / "convert_manifest.json"
    manifest_path.write_text(
        json.dumps(manifest.model_dump(mode="python"), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return manifest_path


def main() -> int:
    args = _parse_args() 
    _configure_logging(verbose=bool(getattr(args, "verbose", False)))
    config = _build_config(args)
    config.output_dir.mkdir(parents=True, exist_ok=True)
    config.debug_output_dir.mkdir(parents=True, exist_ok=True)

    ensure_imports(config.repo_root)
    try:
        prepared_inputs = _resolve_prepared_inputs(config)
    except ValueError:
        return 1
    if not prepared_inputs:
        return 0

    packages = _resolve_packages(config, prepared_inputs)
    if not packages:
        return 0

    surya_structure_root = _prepare_surya_context(config, packages)
    manifest = _new_manifest()
    for pkg in packages:
        _convert_package(config, pkg, surya_structure_root, manifest)

    manifest_path = _write_manifest(config.output_dir, manifest)

    logger.info("Wrote package outputs under: %s", config.output_dir.resolve())
    logger.info("Wrote: %s", manifest_path.resolve())
    logger.info(
        "Summary: "
        "packages=%s "
        "slides=%s "
        "failed=%s "
        "tables=%s "
        "table_skipped=%s "
        "images_resolved=%s "
        "images_unresolved=%s",
        manifest.summary.processed_packages,
        manifest.summary.processed_slides,
        manifest.summary.failed,
        manifest.summary.table_blocks,
        manifest.summary.table_skipped_blocks,
        manifest.summary.resolved_images,
        manifest.summary.unresolved_images,
    )
    return 1 if manifest.summary.failed > 0 else 0


if __name__ == "__main__":
    raise SystemExit(main())
