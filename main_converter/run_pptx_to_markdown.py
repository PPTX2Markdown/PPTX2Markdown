#!/usr/bin/env python3
"""
Convert extracted PPTX package(s) to markdown.

Usage:
  python run_pptx_to_markdown.py [package_name|file.pptx ...]

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
import posixpath
import re
import shutil
import subprocess
import sys
import time
import zipfile
from itertools import count
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple
import xml.etree.ElementTree as ET

from chart2md import ZipContext, convert_chart
from omml2latex import convert_omml
from smartart2md import ZipContext as SmartArtZipContext, convert_smartart

# REPO Root dir - 현재 /main_converter/* 위치이니 root는 .parent.parent가 된다.
REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from asset_utils import copy_media_asset
from converter_models import (
    ConversionManifest,
    ConverterConfig,
    ParagraphSegment,
    PreparedPackage,
    ShapeBlock,
    SlideStats,
)
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
    DEFAULT_OPENAI_API_KEY_ENV,
    DEFAULT_OPENAI_MODEL,
    DEFAULT_PROMPT as DEFAULT_IMAGE_VLM_PROMPT,
    DEFAULT_PROVIDER as DEFAULT_IMAGE_VLM_PROVIDER,
    extract_markdown_from_image,
    normalize_provider,
)
from slide_converter import (
    SlideConversionContext,
    SlideConversionDeps,
    SlideRenderAssets,
    convert_one_slide as convert_one_slide_core,
)
from table_overlay import (
    collect_table_overlay_pictures as collect_table_overlay_pictures_core,
    convert_table_to_markdown as convert_table_to_markdown_core,
    shape_id_of as shape_id_of_core,
)


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
REL_NS = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}

logger = logging.getLogger(__name__)
_SMARTART_ASSET_COUNTER = count()


# 로거 출력 레벨과 포맷을 한 번에 설정한다.
# verbose 여부에 따라 DEBUG/INFO를 전환하고, 이후 전체 변환 파이프라인의 로그 형식을 통일한다.
def _configure_logging(verbose: bool = False) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(level=level, format="%(message)s")


# 추출된 PPTX 패키지를 저장할 기본 디렉터리를 반환한다.
# 항상 main_converter/target_slides 아래를 사용하며, 디렉터리가 없으면 생성하고
# 같은 이름의 일반 파일이 있으면 잘못된 상태로 보고 예외를 발생시킨다.
def default_target_dir(base_dir: Path) -> Path:
    # Always anchor under main_converter/.
    local_target = base_dir / "target_slides"
    if local_target.exists() and not local_target.is_dir():
        raise NotADirectoryError(f"target_slides path exists but is not a directory: {local_target}")
    local_target.mkdir(parents=True, exist_ok=True)
    return local_target


# 원본 PPTX 입력 파일을 모아둘 기본 디렉터리를 반환한다.
# 항상 main_converter/target_pptx 아래를 사용하며, 디렉터리 존재를 보장해
# 이후 자동 탐색 로직이 별도 분기 없이 동작하게 만든다.
def default_pptx_input_dir(base_dir: Path) -> Path:
    # Always anchor under main_converter/.
    local_target = base_dir / "target_pptx"
    if local_target.exists() and not local_target.is_dir():
        raise NotADirectoryError(f"target_pptx path exists but is not a directory: {local_target}")
    local_target.mkdir(parents=True, exist_ok=True)
    return local_target


# 파일명 안의 숫자를 자연 정렬 기준으로 바꿔준다.
# 예를 들어 slide2, slide10 같은 이름을 문자열 순서가 아니라 사람이 기대하는 순서대로 정렬할 때 사용한다.
def natural_key(name: str) -> Tuple:
    parts = re.split(r"(\d+)", name)
    out: List[object] = []
    for part in parts:
        if part.isdigit():
            out.append(int(part))
        else:
            out.append(part.lower())
    return tuple(out)


# XML 태그에서 네임스페이스를 제거하고 로컬 태그명만 꺼낸다.
# ElementTree가 {namespace}tag 형태를 사용하므로, 분기 처리를 단순하게 만들기 위한 유틸이다.
def local_name(tag: str) -> str:
    return tag.split("}", 1)[-1]


# 여러 줄과 중복 공백을 하나의 읽기 쉬운 문자열로 정규화한다.
# XML 텍스트 노드나 오버레이 텍스트를 Markdown으로 옮기기 전에 공백 잡음을 줄이는 데 쓴다.
def normalize_text(s: str) -> str:
    s = s.replace("\n", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


# 이미 추출된 패키지 디렉터리가 현재 PPTX 원본과 동일한 입력에서 만들어졌는지 검사한다.
# .pptx_source.json 안의 경로/크기/mtime 정보를 원본 파일의 현재 상태와 비교해 캐시 재사용 가능 여부를 판단한다.
def package_marker_matches(pkg_dir: Path, pptx_path: Path) -> bool:
    marker = pkg_dir / ".pptx_source.json"
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


# PPTX(zip) 내부 엔트리를 안전하게 검증한 뒤 지정한 폴더에 압축 해제한다.
# 상대경로 탈출 같은 위험한 엔트리를 막아서, 외부 경로로 파일이 풀리는 zip slip 문제를 방지한다.
def safe_extract_pptx(pptx_path: Path, dest_dir: Path) -> None:
    with zipfile.ZipFile(pptx_path) as zf:
        for member in zf.infolist():
            member_path = Path(member.filename)
            if member_path.is_absolute() or ".." in member_path.parts:
                raise ValueError(f"Unsafe archive entry: {member.filename}")
        zf.extractall(dest_dir)


# 원본 PPTX 파일을 target_slides 아래 관리되는 패키지 디렉터리로 추출한다.
# 기존 추출 결과가 같은 원본에서 생성된 경우 재사용하고, 아니라면 필요 시 삭제 후 다시 풀며,
# 추적용 marker 파일과 target_pptx 쪽의 staged 복사본도 함께 맞춰 둔다.
def extract_pptx_to_target(
    package: PreparedPackage,
    extraction_root: Path,
    staged_pptx_root: Path,
    allow_replace_unmanaged: bool = False,
) -> PreparedPackage:
    pptx_path = package.source_pptx_path
    stat = pptx_path.stat()
    extraction_root.mkdir(parents=True, exist_ok=True)
    pkg_name = package.package_dir_name
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
        return package.with_package_dir(pkg_dir.resolve())

    marker = pkg_dir / ".pptx_source.json"
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

    return package.with_package_dir(pkg_dir.resolve())


# Office가 임시로 만드는 잠금 파일인지 판별한다.
# "~$"로 시작하는 PPTX는 실제 입력으로 처리하면 안 되므로 자동 탐색에서 제외한다.
def is_ignored_pptx_file(path: Path) -> bool:
    # Skip Office lock/temp files like "~$sample1.pptx".
    return path.name.startswith("~$")


# 사용자가 넘긴 입력 문자열을 실제 PPTX 파일 경로로 해석한다.
# 현재 작업 디렉터리, 기본 입력 디렉터리, 추출 디렉터리를 차례로 후보에 넣고
# 확장자가 생략된 경우 ".pptx"를 보완해서 찾는다. 실패 시에는 확인한 후보 목록도 함께 돌려준다.
def resolve_input_pptx_path(cwd: Path, item: str) -> Tuple[Optional[Path], List[Path]]:
    search_roots = [cwd, default_pptx_input_dir(cwd), default_target_dir(cwd)]

    candidates: List[Path] = []
    seen: set[str] = set()

    # 같은 경로 후보가 여러 번 들어오지 않도록 중복을 제거하면서 순서를 유지한다.
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


# 입력 인자를 실제 추출 대상 패키지 경로 목록으로 준비한다.
# 각 입력을 PPTX 파일로 해석한 뒤 target_slides 아래로 추출하고,
# 못 찾은 입력은 어떤 경로들을 확인했는지와 함께 별도로 수집한다.
def prepare_package_inputs(
    cwd: Path,
    raw_inputs: Sequence[str],
    force_extract: bool = False,
) -> Tuple[List[PreparedPackage], List[Dict[str, object]]]:
    if not raw_inputs:
        return [], []

    prepared: List[PreparedPackage] = []
    missing_inputs: List[Dict[str, object]] = []
    extraction_root = default_target_dir(cwd)
    staged_pptx_root = default_pptx_input_dir(cwd)

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
        package = PreparedPackage(
            package_dir=extraction_root / picked_file.stem,
            source_pptx_path=picked_file,
        )
        package = extract_pptx_to_target(
            package,
            extraction_root,
            staged_pptx_root=staged_pptx_root,
            allow_replace_unmanaged=force_extract,
        )
        prepared.append(package)
    return prepared, missing_inputs


# 기본 입력 디렉터리 아래의 모든 PPTX 파일을 자동 수집한다.
# 잠금 파일은 제외하고, 파일명은 자연 정렬한 뒤 중복 없는 절대경로 문자열 목록으로 반환한다.
def collect_target_pptx_inputs(cwd: Path) -> List[str]:
    files: List[Path] = []
    for path in default_pptx_input_dir(cwd).glob("*.pptx"):
        if path.is_file() and not is_ignored_pptx_file(path):
            files.append(path.resolve())
    files = sorted(files, key=lambda p: natural_key(p.name))
    uniq: Dict[str, Path] = {str(p): p for p in files}
    return [str(p) for p in uniq.values()]


# slide12.xml 같은 파일명에서 슬라이드 번호를 추출한다.
# 정규식으로 번호를 찾지 못하는 예외 케이스에서는 호출자가 넘긴 기본 인덱스를 대신 사용한다.
def parse_slide_number(filename: str, default_idx: int) -> int:
    m = re.search(r"slide(\d+)", filename, re.IGNORECASE)
    if m:
        return int(m.group(1))
    return default_idx


# reordered XML 같은 파생 파일에서도 원본 슬라이드 기준 이름을 복원한다.
# slide2.reordered.xml -> slide2.xml 형태로 맞춰서 rels나 sidecar 파일을 찾을 때 사용한다.
def slide_base_name(slide_xml: Path) -> str:
    m = re.search(r"(slide\d+)", slide_xml.stem, re.IGNORECASE)
    if m:
        return f"{m.group(1)}.xml"
    return slide_xml.name


# 슬라이드 XML 옆에 생성된 구조 분석 sidecar JSON 파일을 찾는다.
# reordered 버전과 원본 버전 모두 고려하며, structure_analysis.json과 reading_order.json 두 이름 체계를 모두 지원한다.
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


# 구조 분석 단계에서 생성한 heading 힌트를 읽어 shape_id 기준 맵으로 바꾼다.
# 이후 본문/제목 판별 로직이 XML을 다시 계산하지 않고도 점수, depth, placeholder 정보를 바로 참조할 수 있게 한다.
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


# sidecar JSON 안의 input_xml 정보를 이용해 원본 슬라이드의 rels 파일을 찾는다.
# reordered 슬라이드처럼 현재 파일 위치만으로는 관계 파일을 바로 찾기 어려운 경우를 보완하는 용도다.
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


# 현재 슬라이드 XML 주변의 전형적인 위치들에서 rels 파일을 찾는다.
# sidecar 정보가 없을 때 같은 패키지 안에서 가장 자연스러운 후보들을 순서대로 확인하는 fallback 역할이다.
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


# 원본 슬라이드 XML이 따로 주어진 경우 그 파일의 rels 경로를 계산한다.
# Surya처럼 reordered XML과 source XML이 분리되는 모드에서 마지막 보조 수단으로 사용된다.
def source_slide_rels(source_slide_xml: Optional[Path]) -> Optional[Path]:
    if source_slide_xml is None or not source_slide_xml.exists():
        return None
    rels = source_slide_xml.parent / "_rels" / f"{source_slide_xml.name}.rels"
    if rels.exists() and rels.is_file():
        return rels
    return None


# slide rels XML을 rId -> target 경로 맵으로 파싱한다.
# 이미지, 다이어그램, 기타 임베디드 리소스의 실제 파일 위치를 나중에 빠르게 해석하기 위한 전처리 단계다.
# slide rels XML을 rId -> target 경로 맵으로 파싱한다.
# 이미지, 다이어그램, 기타 임베디드 리소스의 실제 파일 위치를
# 나중에 빠르게 해석하기 위한 전처리 단계다.
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


# 슬라이드 하나를 처리할 때 사용할 rels 파일을 우선순위에 따라 결정한다.
# sidecar 기반 경로를 먼저 시도하고, 실패하면 현재 슬라이드 주변 fallback, 마지막으로 source slide rels까지 본다.
# 슬라이드 하나를 처리할 때 사용할 rels 파일을 우선순위에 따라 결정한다.
# sidecar 기반 경로를 먼저 시도하고, 실패하면 현재 슬라이드 주변 fallback,
# 마지막으로 source slide rels까지 확인한다.
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


# r:embed 값과 rels 정보를 이용해 실제 이미지 파일 경로를 결정한다.
# 실패 시에는 [unresolved-image:*] 형태의 플레이스홀더와 이유를 반환하고,
# 성공 시에는 target_slides 패키지 구조까지 반영한 절대경로를 돌려준다.
# r:embed 값과 rels 정보를 이용해 실제 이미지 파일 경로를 결정한다.
# 실패 시에는 [unresolved-image:*] 형태의 플레이스홀더와 이유를 반환하고,
# 성공 시에는 target_slides 패키지 구조까지 반영한 절대경로를 돌려준다.
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


# 절대경로나 외부 기준 경로를 출력 Markdown 기준 상대경로로 바꾼다.
# unresolved placeholder는 그대로 두고, 실제 파일 경로만 output 디렉터리 기준으로 재기록한다.
def relativize_markdown_path(path: str, output_dir: Optional[Path]) -> str:
    if path.startswith("[unresolved-image") or output_dir is None:
        return path
    try:
        return os.path.relpath(path, start=str(output_dir))
    except Exception:
        return path


# 단일 이미지 파일을 VLM에 보내 Markdown 설명으로 바꾼다.
# provider/model 설정을 정리하고, 파이프라인 예외를 사용자 경고 메시지와
# "사용 불가" 여부로 정규화해 반환한다.
def convert_picture_to_markdown(
    image_path: str,
    *,
    provider: str = DEFAULT_IMAGE_VLM_PROVIDER,
    model_spec: Optional[str] = None,
    prompt: str = DEFAULT_IMAGE_VLM_PROMPT,
    max_new_tokens: int = DEFAULT_IMAGE_VLM_MAX_NEW_TOKENS,
    gemini_api_key_env: str = DEFAULT_GEMINI_API_KEY_ENV,
    ignore_cache: bool = False,
) -> Tuple[Optional[str], Optional[str], bool, Optional[Dict[str, object]]]:
    # 의존성 미설치나 API 키 누락처럼 설정 문제로 파이프라인이 아예 못 도는 경우를 분류한다.
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
    if normalized_provider == "openai" and not effective_model:
        effective_model = DEFAULT_OPENAI_MODEL
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
            ignore_cache=ignore_cache,
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


# 이미지 경로를 커스텀 이미지 태그 문자열로 렌더링한다.
# downstream 파서가 기대하는 [img(src="...")] 포맷으로 통일한다.
def render_image_tag(path: str) -> str:
    return f'[img(src="{path}")]'


# 이미지 경로를 커스텀 이미지 태그로 렌더링한다.
# 필요하면 media 디렉터리로 복사하고, 이미지 VLM이 켜져 있으면
# 단순 링크 대신 생성된 Markdown 설명을 우선 사용한다.
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
    ignore_image_vlm_cache: bool = False,
) -> Tuple[str, Optional[str], bool, bool, bool]:
    if path.startswith("[unresolved-image"):
        return path, None, False, False, False

    normalized_provider = normalize_provider(image_vlm_provider)
    image_vlm_enabled = bool(image_vlm_model) or normalized_provider in {"gemini", "openai"}
    if image_vlm_enabled:
        image_md, image_warn, unavailable, result = convert_picture_to_markdown(
            path,
            provider=normalized_provider,
            model_spec=image_vlm_model,
            prompt=image_vlm_prompt,
            max_new_tokens=image_vlm_max_new_tokens,
            gemini_api_key_env=image_vlm_api_key_env,
            ignore_cache=ignore_image_vlm_cache,
        )
        if image_md is not None:
            return annotate_generated_image_markdown(image_md, path), None, unavailable, True, False
        skipped_no_markdown = isinstance(result, dict) and str(result.get("status", "")) == "no_markdown"
        copied_path = copy_media_asset(path, media_dir=media_dir, copied_media=copied_media)
        relative_path = relativize_markdown_path(copied_path, output_dir)
        return render_image_tag(relative_path), image_warn, unavailable, False, skipped_no_markdown

    copied_path = copy_media_asset(path, media_dir=media_dir, copied_media=copied_media)
    relative_path = relativize_markdown_path(copied_path, output_dir)
    return render_image_tag(relative_path), None, False, False, False


def _sanitize_inline_latex(latex: str) -> str:
    sanitized = " ".join(part.strip() for part in latex.splitlines() if part.strip())
    sanitized = re.sub(r"\s+", " ", sanitized).strip()
    return sanitized.strip("$").strip()


def _sanitize_block_latex(latex: str) -> str:
    sanitized = latex.strip()
    if sanitized.startswith("$$") and sanitized.endswith("$$"):
        sanitized = sanitized[2:-2].strip()
    sanitized = sanitized.strip("$").strip()
    return sanitized


def _build_inline_math_segment(math_elem: ET.Element) -> ParagraphSegment:
    latex = _sanitize_inline_latex(convert_omml(math_elem))
    if not latex:
        raise ValueError("empty latex")
    return ParagraphSegment(kind="math_inline", text=f"${latex}$")


def _build_block_math_segment(math_elem: ET.Element) -> ParagraphSegment:
    latex = _sanitize_block_latex(convert_omml(math_elem))
    if not latex:
        raise ValueError("empty latex")
    return ParagraphSegment(kind="math_block", text=f"$$\n{latex}\n$$")


def _extract_run_text(run_elem: ET.Element) -> str:
    texts = [node.text or "" for node in run_elem.findall(".//a:t", NS) if (node.text or "").strip()]
    return normalize_text(" ".join(texts))


def _append_math_segments(container: ET.Element, segments: List[ParagraphSegment]) -> bool:
    container_tag = local_name(container.tag)
    if container_tag == "oMath":
        try:
            segments.append(_build_inline_math_segment(container))
        except Exception:
            fallback = normalize_text(" ".join(node.text or "" for node in container.findall(".//m:t", NS)))
            segments.append(ParagraphSegment(kind="math_error", text=fallback or "[unsupported-math]"))
        return True
    if container_tag == "oMathPara":
        try:
            segments.append(_build_block_math_segment(container))
        except Exception:
            fallback = normalize_text(" ".join(node.text or "" for node in container.findall(".//m:t", NS)))
            segments.append(ParagraphSegment(kind="math_error", text=fallback or "[unsupported-math]"))
        return True

    appended = False
    for math_elem in list(container):
        tag = local_name(math_elem.tag)
        if tag == "oMath":
            appended = True
            try:
                segments.append(_build_inline_math_segment(math_elem))
            except Exception:
                fallback = normalize_text(" ".join(node.text or "" for node in math_elem.findall(".//m:t", NS)))
                segments.append(ParagraphSegment(kind="math_error", text=fallback or "[unsupported-math]"))
            continue
        if tag == "oMathPara":
            appended = True
            try:
                segments.append(_build_block_math_segment(math_elem))
            except Exception:
                fallback = normalize_text(" ".join(node.text or "" for node in math_elem.findall(".//m:t", NS)))
                segments.append(ParagraphSegment(kind="math_error", text=fallback or "[unsupported-math]"))
            continue
    return appended


def parse_paragraph_segments(paragraph: ET.Element) -> List[ParagraphSegment]:
    segments: List[ParagraphSegment] = []
    for child in list(paragraph):
        tag = local_name(child.tag)
        if tag in {"r", "fld"}:
            text = _extract_run_text(child)
            if text:
                segments.append(ParagraphSegment(kind="text", text=text))
            continue
        if tag == "br":
            segments.append(ParagraphSegment(kind="break", text=""))
            continue
        if tag == "m":
            if _append_math_segments(child, segments):
                continue
        if tag == "oMath":
            _append_math_segments(child, segments)
            continue
        if tag == "oMathPara":
            _append_math_segments(child, segments)
            continue
    return segments


def paragraph_text(paragraph: ET.Element) -> str:
    parts: List[str] = []
    for segment in parse_paragraph_segments(paragraph):
        if segment.kind == "break":
            continue
        plain = segment.text
        if segment.kind in {"math_inline", "math_block"}:
            plain = plain.replace("$", " ")
        normalized = normalize_text(plain)
        if normalized:
            parts.append(normalized)
    return normalize_text(" ".join(parts))


# 문단의 목록 레벨(lvl)을 읽어 Markdown 들여쓰기 깊이 계산에 쓴다.
# lvl이 없거나 숫자로 해석할 수 없으면 None을 반환한다.
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


# 문단이 단순 텍스트가 아니라 목록 항목 의미를 가지는지 판정한다.
# lvl, buChar, buAutoNum 존재 여부를 통해 unordered/ordered list 후보를 식별한다.
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


# shape 내부 블록들 중 짧은 일반 텍스트를 주변 리스트 문맥에 맞춰 리스트 항목으로 승격한다.
# PowerPoint가 시각적으로만 맞춰 둔 문단을 Markdown 리스트 구조로 더 자연스럽게 복원하기 위한 보정이다.
def promote_plain_text_to_list(
    blocks: Sequence[ShapeBlock],
) -> List[ShapeBlock]:
    if not any(block.kind in {"list_ul", "list_ol"} for block in blocks):
        return list(blocks)

    text_blocks = [(idx, block.plain_text) for idx, block in enumerate(blocks) if block.kind == "text"]
    if len(text_blocks) < 3:
        return list(blocks)

    first_list_kind = next((block.kind for block in blocks if block.kind in {"list_ul", "list_ol"}), "list_ul")
    promoted = list(blocks)
    for idx, text in text_blocks:
        if len(text) > 80 or text.endswith((".", ":")):
            continue
        promoted[idx] = ShapeBlock(kind=first_list_kind, segments=promoted[idx].segments, level=0)
    return promoted


# PowerPoint의 불연속 목록 레벨을 0,1,2... 형태의 연속 깊이로 정규화한다.
# 원본 lvl 값이 듬성듬성해도 Markdown 렌더링 들여쓰기가 과도하게 깊어지지 않도록 막는다.
def normalize_list_levels(blocks: Sequence[ShapeBlock]) -> List[ShapeBlock]:
    levels = sorted({int(block.level or 0) for block in blocks if block.kind in {"list_ul", "list_ol"}})
    if not levels:
        return list(blocks)
    remap = {level: idx for idx, level in enumerate(levels)}
    normalized: List[ShapeBlock] = []
    for block in blocks:
        if block.kind not in {"list_ul", "list_ol"}:
            normalized.append(block)
            continue
        mapped = remap[int(block.level or 0)]
        normalized.append(ShapeBlock(kind=block.kind, segments=block.segments, level=mapped))
    return normalized


# shape 하나에서 텍스트 문단들을 추출해 중간 표현 블록 목록으로 바꾼다.
# 각 문단을 plain text / unordered list / ordered list로 분류하고 필요한 level도 함께 기록한다.
def extract_shape_blocks(shape_elem: ET.Element) -> List[ShapeBlock]:
    blocks: List[ShapeBlock] = []
    for p in shape_elem.findall(".//p:txBody/a:p", NS):
        segments = parse_paragraph_segments(p)
        if not segments:
            continue
        plain_text = paragraph_text(p)
        if not plain_text and not any(segment.kind in {"math_inline", "math_block"} for segment in segments):
            continue
        p_pr = p.find("./a:pPr", NS)
        has_auto_num = p_pr is not None and p_pr.find("./a:buAutoNum", NS) is not None
        if paragraph_has_list_semantics(p):
            level = paragraph_level(p)
            blocks.append(
                ShapeBlock(
                    kind=("list_ol" if has_auto_num else "list_ul"),
                    segments=segments,
                    level=0 if level is None else level,
                )
            )
        else:
            blocks.append(ShapeBlock(kind="text", segments=segments, level=None))
    return blocks


# 중간 표현 블록 목록을 최종 Markdown 문자열로 렌더링한다.
# 리스트 번호 재계산, 들여쓰기, 일반 문단 출력까지 한 번에 수행한다.
def render_shape_blocks(blocks: Sequence[ShapeBlock]) -> str:
    if not blocks:
        return ""
    blocks = normalize_list_levels(promote_plain_text_to_list(blocks))

    rendered: List[str] = []
    ordered_counters: Dict[int, int] = {}
    for idx, block in enumerate(blocks):
        text = block.markdown_text
        if block.kind in {"list_ul", "list_ol"}:
            indent = "  " * max(0, int(block.level or 0))
            if block.kind == "list_ol":
                clean_text = re.sub(r"^\d+\s*\.\s*", "", text).strip() or text
                lvl = max(0, int(block.level or 0))
                ordered_counters[lvl] = ordered_counters.get(lvl, 0) + 1
                for deeper in [k for k in ordered_counters.keys() if k > lvl]:
                    del ordered_counters[deeper]
                rendered.append(f"{indent}{ordered_counters[lvl]}. {clean_text}")
            else:
                rendered.append(f"{indent}- {text}")
            continue

        prev_kind = blocks[idx - 1].kind if idx > 0 else None
        next_kind = blocks[idx + 1].kind if idx + 1 < len(blocks) else None
        if prev_kind in {"list_ul", "list_ol"} and next_kind in {"list_ul", "list_ol"}:
            prev_level = max(0, int(blocks[idx - 1].level or 0))
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


# graphicFrame이 어떤 종류의 객체인지 식별한다.
# 현재는 diagram/chart를 구분해 이후 전용 처리 로직으로 분기하는 데 사용한다.
def graphic_frame_kind(graphic_frame: ET.Element) -> Optional[str]:
    graphic_data = graphic_frame.find("./a:graphic/a:graphicData", NS)
    if graphic_data is None:
        return None
    uri = graphic_data.attrib.get("uri", "").strip()
    if uri.endswith("/diagram"):
        return "diagram"
    if "chart" in uri:
        return "chart"
    for element in graphic_data.iter():
        if local_name(element.tag) == "chart":
            return "chart"
    if uri.endswith("/chart"):
        return "chart"
    return None


def _resolve_ooxml_target(base_part_path: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base_dir = posixpath.dirname(base_part_path)
    return posixpath.normpath(posixpath.join(base_dir, target)).lstrip("/")


def _chart_relationship_id(graphic_frame: ET.Element) -> Optional[str]:
    graphic_data = graphic_frame.find("./a:graphic/a:graphicData", NS)
    if graphic_data is None:
        return None
    for element in graphic_data.iter():
        if local_name(element.tag) != "chart":
            continue
        for attr_name, value in element.attrib.items():
            if attr_name == "r:id" or attr_name.endswith("}id"):
                return value
    return None


def _part_path_from_rels_path(rels_path: Path) -> Optional[str]:
    parts = rels_path.parts
    try:
        ppt_index = parts.index("ppt")
    except ValueError:
        return None
    rels_part_path = Path(*parts[ppt_index:]).as_posix()
    return rels_part_path.replace("/_rels/", "/").removesuffix(".rels")


def _smartart_media_markdown(
    markdown: str,
    images: Sequence[Tuple[bytes, str]],
    *,
    output_dir: Optional[Path],
    media_dir: Optional[Path],
) -> str:
    if not images:
        return re.sub(r"@@IMG:\d+@@", "", markdown).strip()

    rendered = markdown
    for index, (data, ext) in enumerate(images):
        normalized_ext = (ext or "png").lower().strip(".") or "png"
        if normalized_ext == "jpeg":
            normalized_ext = "jpg"

        if media_dir is None:
            replacement = ""
        else:
            media_dir.mkdir(parents=True, exist_ok=True)
            asset_id = next(_SMARTART_ASSET_COUNTER)
            asset_path = media_dir / f"smartart-{asset_id}.{normalized_ext}"
            asset_path.write_bytes(data)
            relative_path = relativize_markdown_path(str(asset_path), output_dir)
            replacement = render_image_tag(relative_path)
        rendered = rendered.replace(f"@@IMG:{index}@@", replacement)

    return re.sub(r"@@IMG:\d+@@", "", rendered).strip()


def convert_chart_to_markdown(
    graphic_frame: ET.Element,
    rels_path: Optional[Path],
    rels_map: Dict[str, str],
    source_pptx_path: Optional[Path],
) -> Tuple[Optional[str], Optional[str]]:
    if source_pptx_path is None:
        return None, "chart conversion failed: missing source pptx path"
    if rels_path is None:
        return None, "chart conversion failed: missing slide rels file"
    rid = _chart_relationship_id(graphic_frame)
    if not rid:
        return None, "chart conversion failed: missing chart relationship id"
    target = rels_map.get(rid)
    if not target:
        return None, f"chart conversion failed: relationship not found: {rid}"

    slide_part_path = _part_path_from_rels_path(rels_path)
    if not slide_part_path:
        return None, f"chart conversion failed: unsupported rels path: {rels_path}"
    chart_part_path = _resolve_ooxml_target(slide_part_path, target)

    try:
        with zipfile.ZipFile(source_pptx_path) as zf:
            chart_root = ET.fromstring(zf.read(chart_part_path))
            ctx = ZipContext(zf, chart_part_path)
            markdown = convert_chart(chart_root, ctx).strip()
    except KeyError:
        return None, f"chart conversion failed: missing chart part: {chart_part_path}"
    except ET.ParseError as exc:
        return None, f"chart conversion failed: invalid chart xml: {chart_part_path}: {exc}"
    except Exception as exc:  # noqa: BLE001
        return None, f"chart conversion failed: {chart_part_path}: {type(exc).__name__}: {exc}"

    if not markdown:
        return None, f"chart conversion failed: empty markdown output: {chart_part_path}"
    return markdown, None


def convert_smartart_to_markdown(
    graphic_frame: ET.Element,
    rels_path: Optional[Path],
    rels_map: Dict[str, str],
    source_pptx_path: Optional[Path],
    output_dir: Optional[Path],
    media_dir: Optional[Path],
) -> Tuple[Optional[str], Optional[str]]:
    if source_pptx_path is None:
        return None, "smartart conversion failed: missing source pptx path"
    diagram_data_xml = diagram_data_path(graphic_frame, rels_path, rels_map)
    if diagram_data_xml is None:
        return None, "smartart conversion failed: missing diagram data path"

    slide_part_path = _part_path_from_rels_path(rels_path) if rels_path is not None else None
    if not slide_part_path:
        return None, f"smartart conversion failed: unsupported rels path: {rels_path}"

    dm_rel_ids = graphic_frame.find("./a:graphic/a:graphicData/dgm:relIds", NS)
    dm_rid = dm_rel_ids.attrib.get(f"{{{NS['r']}}}dm") if dm_rel_ids is not None else None
    target = rels_map.get(dm_rid) if dm_rid else None
    if not target:
        return None, "smartart conversion failed: missing diagram relationship target"

    data_part_path = _resolve_ooxml_target(slide_part_path, target)

    try:
        with zipfile.ZipFile(source_pptx_path) as zf:
            data_root = ET.fromstring(zf.read(data_part_path))
            ctx = SmartArtZipContext(zf, data_part_path)
            markdown, images = convert_smartart(data_root, ctx)
    except KeyError:
        return None, f"smartart conversion failed: missing diagram part: {data_part_path}"
    except ET.ParseError as exc:
        return None, f"smartart conversion failed: invalid diagram xml: {data_part_path}: {exc}"
    except Exception as exc:  # noqa: BLE001
        return None, f"smartart conversion failed: {data_part_path}: {type(exc).__name__}: {exc}"

    rendered = _smartart_media_markdown(
        markdown,
        images,
        output_dir=output_dir,
        media_dir=media_dir,
    )
    if not rendered:
        return None, f"smartart conversion failed: empty markdown output: {data_part_path}"
    return rendered, None


# 다이어그램 graphicFrame에서 실제 데이터 XML 파일 경로를 해석한다.
# rels를 따라가 dgm data model 파일을 찾고, 존재하는 실제 파일일 때만 반환한다.
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


# 삼각형 기호로 시작하는 가짜 불릿 텍스트를 표준 Markdown 불릿 형태로 바꾼다.
# 시각적 문자 불릿을 구조적 리스트로 복원하기 위한 1차 정규화 함수다.
def normalize_triangle_bullet(text: str) -> str:
    raw = re.sub(r"\s+", " ", (text or "").strip())
    if not raw:
        return ""
    raw = re.sub(r"^(\d+)\s+\.\s*", r"\1. ", raw)
    if re.match(r"^▶+\s*", raw):
        return re.sub(r"^▶+\s*", "- ", raw)
    return raw


# 한 줄 안에 여러 개의 삼각형 불릿이 이어진 경우 개별 리스트 항목으로 분리한다.
# 결과는 render 단계에서 바로 붙일 수 있는 Markdown 라인 목록이다.
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


# XML 요소에서 PowerPoint 내부 shape id를 꺼낸다.
# heading hints, overlay 매핑 등 다른 분석 결과와 현재 shape를 연결하는 공통 키다.
def shape_id_of(elem: ET.Element) -> str:
    return shape_id_of_core(elem, NS)


# 테이블 오버레이 이미지를 커스텀 이미지 태그 문자열로 바꾼다.
# 필요 시 media 디렉터리로 복사하고, 문서 출력 위치 기준 상대경로로 다시 쓴다.
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
    return render_image_tag(path)


# 이미지 VLM이 생성한 Markdown 앞에 원본 이미지 출처 정보를 덧붙인다.
# 사람이 결과를 검토할 때 어떤 파일에서 생성된 설명인지 추적할 수 있게 한다.
def annotate_generated_image_markdown(markdown: str, image_path: str) -> str:
    image_name = Path(image_path).name
    body = markdown.strip()
    if not body:
        return f"[image-vlm-source: {image_name}]"
    return f"[image-vlm-source: {image_name}]\n\n{body}"


# 테이블 오버레이 이미지 하나를 최종 텍스트로 변환한다.
# VLM 설명을 우선 시도하고, 실패하거나 비활성화된 경우에는 파일 링크로 대체한다.
def overlay_content_text(
    path: str,
    output_dir: Optional[Path],
    media_dir: Optional[Path] = None,
    copied_media: Optional[Dict[str, Path]] = None,
    image_vlm_provider: str = DEFAULT_IMAGE_VLM_PROVIDER,
    image_vlm_model: Optional[str] = None,
    image_vlm_prompt: str = DEFAULT_IMAGE_VLM_PROMPT,
    image_vlm_max_new_tokens: int = DEFAULT_IMAGE_VLM_MAX_NEW_TOKENS,
    image_vlm_api_key_env: str = DEFAULT_GEMINI_API_KEY_ENV,
    ignore_image_vlm_cache: bool = False,
) -> Tuple[str, Optional[str], bool, bool, bool]:
    if path.startswith("[unresolved-image"):
        return path, None, False, False, False

    normalized_provider = normalize_provider(image_vlm_provider)
    image_vlm_enabled = bool(image_vlm_model) or normalized_provider in {"gemini", "openai"}
    if image_vlm_enabled:
        image_md, image_warn, unavailable, result = convert_picture_to_markdown(
            path,
            provider=normalized_provider,
            model_spec=image_vlm_model,
            prompt=image_vlm_prompt,
            max_new_tokens=image_vlm_max_new_tokens,
            gemini_api_key_env=image_vlm_api_key_env,
            ignore_cache=ignore_image_vlm_cache,
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


# 슬라이드 전체에서 테이블 위에 겹쳐진 picture overlay들을 수집한다.
# 실제 구현은 table_overlay 모듈에 두고, 여기서는 현재 파일의 namespace와
# 이미지 경로 해석 함수를 주입하는 어댑터 역할만 맡는다.
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


# graphicFrame 테이블을 Markdown 표 문자열로 변환한다.
# 오버레이 텍스트 생성, 이미지 경로 해석, 텍스트 정규화 같은
# 현재 모듈의 정책 함수를 core 구현에 넘겨주는 어댑터다.
def convert_table_to_markdown(
    graphic_frame: ET.Element,
    overlays: Optional[Sequence[Dict[str, object]]] = None,
    output_dir: Optional[Path] = None,
    media_dir: Optional[Path] = None,
    copied_media: Optional[Dict[str, Path]] = None,
    rels_path: Optional[Path] = None,
    rels_map: Optional[Dict[str, str]] = None,
    image_vlm_provider: str = DEFAULT_IMAGE_VLM_PROVIDER,
    image_vlm_model: Optional[str] = None,
    image_vlm_prompt: str = DEFAULT_IMAGE_VLM_PROMPT,
    image_vlm_max_new_tokens: int = DEFAULT_IMAGE_VLM_MAX_NEW_TOKENS,
    image_vlm_api_key_env: str = DEFAULT_GEMINI_API_KEY_ENV,
    ignore_image_vlm_cache: bool = False,
) -> Tuple[Optional[str], Optional[str]]:
    return convert_table_to_markdown_core(
        graphic_frame,
        overlays=overlays,
        output_dir=output_dir,
        media_dir=media_dir,
        copied_media=copied_media,
        rels_path=rels_path,
        rels_map=rels_map,
        image_vlm_provider=image_vlm_provider,
        image_vlm_model=image_vlm_model,
        image_vlm_prompt=image_vlm_prompt,
        image_vlm_max_new_tokens=image_vlm_max_new_tokens,
        image_vlm_api_key_env=image_vlm_api_key_env,
        ignore_image_vlm_cache=ignore_image_vlm_cache,
        ns=NS,
        normalize_text_fn=normalize_text,
        overlay_content_text_fn=overlay_content_text,
        resolve_image_path_fn=resolve_image_path,
    )


# slide_converter core가 필요로 하는 의존성 묶음을 구성한다.
# 이 파일에 정의된 XML 파싱/렌더링 정책을 하나의 객체로 모아 core 구현에 주입한다.
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
        format_markdown_image=format_markdown_image,
        convert_table_to_markdown=convert_table_to_markdown,
        graphic_frame_kind=graphic_frame_kind,
        convert_chart_to_markdown=convert_chart_to_markdown,
        convert_smartart_to_markdown=convert_smartart_to_markdown,
        normalize_single_heading_to_h1=normalize_single_heading_to_h1,
    )


# 슬라이드 XML 하나를 Markdown과 통계 정보로 변환하는 진입점이다.
# 실제 본문 순회는 core 구현이 담당하고, 이 함수는 현재 모듈의 정책과 옵션을 연결한다.
def convert_one_slide(
    *,
    context: SlideConversionContext,
    assets: SlideRenderAssets,
    strict_headings: bool = False,
) -> Tuple[str, SlideStats]:
    return convert_one_slide_core(
        context=context,
        assets=assets,
        strict_headings=strict_headings,
        deps=_slide_conversion_deps(),
    )


# CLI 인자를 정의하고 파싱한다.
# 입력 PPTX 목록, 읽기 순서 모드, heading strict 모드, 이미지 VLM 옵션 등
# 전체 변환 파이프라인을 제어하는 설정을 여기서 받는다.
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
        "--not-strict",
        dest="strict",
        action="store_false",
        default=True,
        help="Disable strict heading detection in xml reading-order mode.",
    )
    parser.add_argument(
        "--reuse-surya-cache",
        action="store_true",
        help="Reuse existing Surya structure_ready outputs instead of re-running the Surya pipeline.",
    )
    parser.add_argument(
        "--image-vlm-provider",
        choices=("local", "gemini", "openai"),
        default=DEFAULT_IMAGE_VLM_PROVIDER,
        help="Image VLM backend. local uses Qwen2.5-VL, gemini uses the Gemini API, openai uses the OpenAI Responses API.",
    )
    parser.add_argument(
        "--image-vlm-model",
        help=(
            "Image VLM model identifier. "
            "Use 3b/7b (or a Hugging Face model id) for --image-vlm-provider local, "
            "a Gemini model id such as gemini-2.5-flash for --image-vlm-provider gemini, "
            "or an OpenAI model id such as gpt-4.1-mini for --image-vlm-provider openai."
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
        help="Environment variable name containing the provider API key when --image-vlm-provider gemini or openai is used.",
    )
    parser.add_argument(
        "--ignore-image-vlm-cache",
        "--ignore-vlm-cache",
        action="store_true",
        help="Ignore in-memory and disk cache for image VLM results and recompute them from scratch.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable verbose debug logging.",
    )
    return parser.parse_args()

# 파싱된 CLI 인자를 내부 설정 모델로 변환한다.
# 작업 기준 디렉터리, 출력 위치, 읽기 순서 모드, 이미지 VLM 관련 값을 이후 로직이 일관되게 사용할 수 있는 ConverterConfig로 정규화한다.
def _build_config(args: argparse.Namespace) -> ConverterConfig:
    main_converter_root = REPO_ROOT/ "main_converter"
    normalized_provider = normalize_provider(args.image_vlm_provider)
    api_key_env = str(args.image_vlm_api_key_env).strip()
    if not api_key_env or (
        normalized_provider == "openai" and api_key_env == DEFAULT_GEMINI_API_KEY_ENV
    ):
        api_key_env = DEFAULT_OPENAI_API_KEY_ENV if normalized_provider == "openai" else DEFAULT_GEMINI_API_KEY_ENV
    return ConverterConfig(
        cwd=main_converter_root,
        repo_root=REPO_ROOT,
        output_dir=main_converter_root / "output" / args.reading_order,
        inputs=list(args.inputs),
        reading_order=str(args.reading_order),
        strict=bool(args.strict),
        reuse_surya_cache=bool(args.reuse_surya_cache),
        image_vlm_provider=normalized_provider,
        image_vlm_model=(str(args.image_vlm_model).strip() if args.image_vlm_model else None),
        image_vlm_prompt=str(args.image_vlm_prompt),
        image_vlm_max_new_tokens=max(1, int(args.image_vlm_max_new_tokens)),
        image_vlm_api_key_env=api_key_env,
        ignore_image_vlm_cache=bool(args.ignore_image_vlm_cache),
    )


# 실제 처리에 사용할 입력 패키지 경로 목록을 확정한다.
# 사용자가 직접 넘긴 입력이 있으면 그것만 검증/추출하고,
# 없으면 target_pptx 아래 파일들을 자동 탐색해서 동일한 형식으로 준비한다.
def _resolve_prepared_inputs(config: ConverterConfig) -> List[PreparedPackage]:
    # 누락된 입력에 대해 어떤 경로들을 확인했는지 자세히 로그로 남긴다.
    def _log_missing_inputs(missing_inputs: Sequence[Dict[str, object]]) -> None:
        if not missing_inputs:
            return
        logger.error("Input .pptx file not found.")
        for row in missing_inputs:
            requested = str(row.get("input", "")).strip()
            if requested:
                logger.error("- requested: %s", requested)
            checked = row.get("checked")
            if isinstance(checked, list):
                for candidate in checked:
                    logger.error("  checked: %s", candidate)

    if config.inputs:
        non_pptx_inputs = [x for x in config.inputs if Path(x).suffix.lower() != ".pptx"]
        if non_pptx_inputs:
            logger.error("Only .pptx inputs are allowed.")
            logger.error("Provide files like: sample1.pptx sample2.pptx")
            for item in non_pptx_inputs:
                logger.error("- %s", item)
            raise ValueError("invalid non-pptx inputs")
        prepared_inputs, missing_inputs = prepare_package_inputs(config.cwd, config.inputs, force_extract=True)
        if missing_inputs:
            _log_missing_inputs(missing_inputs)
            raise ValueError("missing pptx inputs")
        return prepared_inputs

    auto_pptx_inputs = collect_target_pptx_inputs(config.cwd)
    if not auto_pptx_inputs:
        logger.info("No .pptx files found in: %s", default_pptx_input_dir(config.cwd).resolve())
        return []
    prepared_inputs, missing_inputs = prepare_package_inputs(config.cwd, auto_pptx_inputs, force_extract=True)
    if missing_inputs:
        _log_missing_inputs(missing_inputs)
        raise ValueError("missing auto-discovered pptx inputs")
    return prepared_inputs


# 패키지 단위의 선행 stage가 실패했을 때, 해당 패키지의 모든 슬라이드를 실패로 기록한다.
# 예를 들어 구조 분석 단계에서 패키지 전체가 막히면 슬라이드별 결과 대신 공통 실패 행을 남긴다.
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
        pkg_row["slides"].append(row)
        manifest.summary.failed += 1
    manifest.packages.append(pkg_row)
    logger.error("[%s] %s failed: %s", package_name, stage_label, error_message)

# PPTX를 압축 해제한 패키지 하나를 끝까지 변환하는 핵심 오케스트레이션 함수다.
# 슬라이드 목록 수집, 읽기 순서 전처리, 슬라이드별 Markdown 변환, 패키지 단위 Markdown 생성과 manifest 누적까지 담당한다.
def _convert_package(
    config: ConverterConfig,
    package: PreparedPackage,
    surya_structure_root: Optional[Path],
    manifest: ConversionManifest,
) -> None:
    pkg = package.package_dir
    pkg_name = pkg.name
    slides_dir = pkg / "ppt" / "slides"
    slide_xmls = sorted(
        [p for p in slides_dir.glob("slide*.xml") if p.is_file()],
        key=lambda p: natural_key(p.name),
    )
    pkg_out = config.output_dir / pkg_name
    output_md_path = package.output_markdown_path(config.output_dir)
    media_dir = pkg_out / "media"
    copied_media: Dict[str, Path] = {}
    pkg_out.mkdir(parents=True, exist_ok=True)

    pkg_row: Dict[str, object] = {
        "package": str(pkg),
        "name": pkg_name,
        "slides": [],
        "result_md": str(output_md_path),
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
            context = SlideConversionContext(
                slide_xml=ordered_slide_xml,
                page_no=page_no,
                ns=NS,
                source_slide_xml=(slide_xml if config.reading_order == "surya" else None),
                source_pptx_path=package.source_pptx_path,
            )
            assets = SlideRenderAssets(
                output_dir=pkg_out,
                media_dir=media_dir,
                copied_media=copied_media,
                image_vlm_provider=config.image_vlm_provider,
                image_vlm_model=config.image_vlm_model,
                image_vlm_prompt=config.image_vlm_prompt,
                image_vlm_max_new_tokens=config.image_vlm_max_new_tokens,
                image_vlm_api_key_env=config.image_vlm_api_key_env,
                ignore_image_vlm_cache=config.ignore_image_vlm_cache,
            )
            merged_md_text, stats = convert_one_slide(
                context=context,
                assets=assets,
                strict_headings=config.strict,
            )
            if config.reading_order == "surya":
                row["surya_source"] = str(structure_output_dir)
            all_chunks.append(merged_md_text.rstrip())

            row.update(
                {"status": "ok", **stats.to_slide_row_fields()}
            )
            manifest.summary.add_slide(stats)
            logger.info("[%s] [md-convert] Processed: %s", pkg_name, slide_xml.name)
            if stats.warnings:
                seen_warnings: set[str] = set()
                for warning in stats.warnings:
                    text = str(warning or "").strip()
                    if not text or text in seen_warnings:
                        continue
                    seen_warnings.add(text)
                    logger.warning(
                        "[%s] Warning: %s (page=%s, slide=%s)",
                        pkg_name,
                        text,
                        page_no,
                        slide_xml.name,
                    )
        except Exception as e:  # noqa: BLE001
            row.update({"status": "failed", "error": str(e)})
            manifest.summary.failed += 1
            logger.error("[%s] Failed: %s -> %s", pkg_name, slide_xml.name, e)
        pkg_row["slides"].append(row)

    merged = "\n\n".join(all_chunks).strip()
    if merged:
        merged += "\n"
    output_md_path.write_text(merged, encoding="utf-8")
    manifest.summary.processed_packages += 1
    manifest.packages.append(pkg_row)


# 전체 변환 파이프라인의 CLI 엔트리포인트다.
# 인자 파싱, 입력 준비, 패키지 선택, Surya 준비, manifest 저장,
# 최종 요약 로그 출력까지 전체 흐름을 조율한다.
def main() -> int:
    args = _parse_args() 
    _configure_logging(verbose=bool(getattr(args, "verbose", False)))
    config = _build_config(args)
    config.output_dir.mkdir(parents=True, exist_ok=True)

    if str(config.repo_root) not in sys.path:
        sys.path.insert(0, str(config.repo_root))
    try:
        prepared_inputs = _resolve_prepared_inputs(config)
    except ValueError:
        return 1
    if not prepared_inputs:
        return 0

    packages = prepared_inputs
    if not packages:
        logger.info("No valid PPTX package directories found.")
        logger.info("Checked default directory:")
        logger.info("- %s", default_target_dir(config.cwd).resolve())
        return 0

    surya_structure_root: Optional[Path] = None
    if config.reading_order == "surya":
        surya_structure_root = prepare_surya_structure_root(
            force=not config.reuse_surya_cache,
            reuse_existing_output=config.reuse_surya_cache,
            targets=[pkg.name for pkg in packages],
            target_pptx_dir=default_pptx_input_dir(config.cwd).resolve(),
            target_slides_dir=default_target_dir(config.cwd).resolve(),
        )

    manifest = ConversionManifest()

    # 각 패키지에 대하여 일괄적으로 메인 컨버터 로직인 _convert_package를 수행한다.
    for pkg in packages:
        _convert_package(config, pkg, surya_structure_root, manifest)

    manifest.mark_finished()
    manifest_path = config.output_dir / "convert_manifest.json"
    manifest_path.write_text(
        json.dumps(manifest.model_dump(mode="python"), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    logger.info("Wrote package outputs under: %s", config.output_dir.resolve())
    logger.info("Wrote: %s", manifest_path.resolve())
    logger.info(
        "Summary: "
        "packages=%s "
        "slides=%s "
        "failed=%s "
        "charts=%s "
        "smartarts=%s "
        "tables=%s "
        "table_skipped=%s "
        "images_resolved=%s "
        "images_unresolved=%s",
        manifest.summary.processed_packages,
        manifest.summary.processed_slides,
        manifest.summary.failed,
        manifest.summary.chart_blocks,
        manifest.summary.smartart_blocks,
        manifest.summary.table_blocks,
        manifest.summary.table_skipped_blocks,
        manifest.summary.resolved_images,
        manifest.summary.unresolved_images,
    )
    return 1 if manifest.summary.failed > 0 else 0


if __name__ == "__main__":
    raise SystemExit(main())
