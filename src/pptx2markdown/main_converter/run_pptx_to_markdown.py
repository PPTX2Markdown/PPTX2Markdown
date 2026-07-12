#!/usr/bin/env python3
"""
Convert PPTX/PPT presentation(s) to Markdown or JSON.

Usage:
  pptx2markdown [file.pptx ...]

Rules:
  - If no positional args are provided, all .pptx/.ppt files in the current
    directory are processed.
  - Output is written to ./output/<name>/<name>.md by default
    (override with --output-dir).
  - Intermediate files (extracted packages, caches) live in ./.pptx2markdown
    by default (override with --work-dir).
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import posixpath
import re
import xml.etree.ElementTree as ET
import zipfile
from itertools import count
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

from chart2md import ZipContext, convert_chart
from omml2latex import convert_omml
from smartart2md import ZipContext as SmartArtZipContext
from smartart2md import convert_smartart

from .asset_utils import copy_media_asset
from .converter_models import (
    ConversionManifest,
    ConverterConfig,
    ParagraphSegment,
    PreparedPackage,
    PresentationDocument,
    ShapeBlock,
    SlideDocument,
    SlideStats,
    render_presentation_markdown,
)
from .package_inputs import (
    collect_target_presentation_inputs,
    default_target_dir,
    natural_key,
    prepare_package_inputs,
)
from .slide_converter import (
    SlideConversionContext,
    SlideConversionDeps,
    SlideRenderAssets,
)
from .slide_converter import (
    convert_one_slide as convert_one_slide_core,
)
from .structure_analysis_pipeline import run_structure_analysis_stage
from .table_overlay import (
    collect_table_overlay_pictures as collect_table_overlay_pictures_core,
)
from .table_overlay import (
    convert_table_to_markdown as convert_table_to_markdown_core,
)
from .table_overlay import (
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

# 실행 중인 변환 작업의 작업 루트(추출/캐시/중간 산출물 저장 위치).
# config 객체가 전달되지 않는 하위 헬퍼들이 fallback 경로를 계산할 때 사용한다.
_WORK_ROOT: Optional[Path] = None


def default_work_root() -> Path:
    return _WORK_ROOT if _WORK_ROOT is not None else (Path.cwd() / ".pptx2markdown")


# 로거 출력 레벨과 포맷을 한 번에 설정한다.
# verbose 여부에 따라 DEBUG/INFO를 전환하고, 이후 전체 변환 파이프라인의 로그 형식을 통일한다.
def _configure_logging(verbose: bool = False) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(level=level, format="%(message)s")


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


# reordered와 원본 슬라이드 이름으로 구조 분석 sidecar를 찾는다.
def find_sidecar_json(slide_xml: Path) -> Optional[Path]:
    # e.g. slide2.reordered.xml -> slide2.structure_analysis.json
    stem = slide_xml.stem
    candidates = []
    if stem.endswith(".reordered"):
        candidates.append(
            slide_xml.with_name(f"{stem[: -len('.reordered')]}.structure_analysis.json")
        )
    candidates.append(slide_xml.with_name(f"{stem}.structure_analysis.json"))
    for c in candidates:
        if c.exists() and c.is_file():
            return c
    return None


# 구조 분석의 heading 힌트를 shape_id 기준 맵으로 읽는다.
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
        return {}
    out: Dict[str, Dict[str, object]] = {}
    for order_index, row in enumerate(rows):
        if not isinstance(row, dict):
            continue
        sid = str(row.get("shape_id", "")).strip()
        if not sid:
            continue
        out[sid] = {
            "order_index": order_index,
            "xml_index": row.get("xml_index"),
            "is_heading_candidate": bool(row.get("is_heading_candidate", False)),
            "heading_score": float(row.get("heading_score", 0.0)),
            "heading_depth_hint": row.get("heading_depth_hint"),
            "font_pt": row.get("font_pt"),
            "ph_type": row.get("ph_type"),
            "is_title_placeholder": bool(row.get("is_title_placeholder", False)),
        }
    return out


def load_effective_properties(slide_xml: Path) -> Dict[str, Dict[str, object]]:
    sidecar = find_sidecar_json(slide_xml)
    if sidecar is None:
        return {}
    try:
        payload = json.loads(sidecar.read_text(encoding="utf-8"))
    except Exception:
        return {}
    rows = payload.get("structure_order")
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
            "list_kind": row.get("list_kind"),
            "list_level": row.get("list_level"),
            "has_list_semantics": bool(row.get("has_list_semantics", False)),
            "ph_type": row.get("ph_type"),
            "is_decorative": bool(row.get("is_decorative", False)),
        }
    return out


# sidecar의 input_xml을 이용해 원본 슬라이드 관계 파일을 찾는다.
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
    work_root = default_work_root()
    if not work_root.exists():
        return None
    for cand in work_root.rglob(base_name):
        if "/ppt/slides/" not in cand.as_posix():
            continue
        rels2 = cand.parent / "_rels" / f"{cand.name}.rels"
        if rels2.exists():
            return rels2
    return None


# 현재 슬라이드 주변의 일반적인 위치에서 관계 파일을 찾는다.
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


# slide 관계 파일을 rId -> target 경로 맵으로 파싱한다.
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


# 같은 패키지 안에서 슬라이드 관계 파일을 우선순위에 따라 선택한다.
def choose_rels_in_package(slide_xml: Path) -> Optional[Path]:
    # Strictly stay inside same ppt package to avoid cross-package mismatches.
    fallback = fallback_rels(slide_xml)
    if fallback:
        return fallback
    sidecar = rels_from_sidecar(slide_xml)
    if sidecar:
        return sidecar
    return None


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
    pkg_name: Optional[str] = None
    try:
        if rels_path.parents[2].name == "ppt":
            pkg_name = rels_path.parents[3].name
    except IndexError:
        pkg_name = None
    if pkg_name:
        target_name = Path(target).name
        remapped = default_work_root() / "target_slides" / pkg_name / "ppt" / "media" / target_name
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


# 이미지 경로를 커스텀 이미지 태그 문자열로 렌더링한다.
# downstream 파서가 기대하는 [img(src="...")] 포맷으로 통일한다.
def render_image_tag(path: str) -> str:
    return f'[img(src="{path}")]'


# 이미지를 media 디렉터리로 복사하고 정적 Markdown 태그로 렌더링한다.
def format_markdown_image(
    path: str,
    output_dir: Optional[Path],
    media_dir: Optional[Path] = None,
    copied_media: Optional[Dict[str, Path]] = None,
) -> str:
    if path.startswith("[unresolved-image"):
        return path

    copied_path = copy_media_asset(path, media_dir=media_dir, copied_media=copied_media)
    relative_path = relativize_markdown_path(copied_path, output_dir)
    return render_image_tag(relative_path)


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
    return "".join(node.text or "" for node in run_elem.findall(".//a:t", NS))


def _append_text_segment(segments: List[ParagraphSegment], text: str) -> None:
    if not text:
        return
    if segments and segments[-1].kind == "text":
        previous = segments[-1]
        segments[-1] = ParagraphSegment(kind="text", text=previous.text + text)
        return
    segments.append(ParagraphSegment(kind="text", text=text))


def _append_math_segments(container: ET.Element, segments: List[ParagraphSegment]) -> bool:
    container_tag = local_name(container.tag)
    if container_tag == "oMath":
        try:
            segments.append(_build_inline_math_segment(container))
        except Exception:
            fallback = normalize_text(
                " ".join(node.text or "" for node in container.findall(".//m:t", NS))
            )
            segments.append(
                ParagraphSegment(kind="math_error", text=fallback or "[unsupported-math]")
            )
        return True
    if container_tag == "oMathPara":
        try:
            segments.append(_build_block_math_segment(container))
        except Exception:
            fallback = normalize_text(
                " ".join(node.text or "" for node in container.findall(".//m:t", NS))
            )
            segments.append(
                ParagraphSegment(kind="math_error", text=fallback or "[unsupported-math]")
            )
        return True

    appended = False
    for math_elem in list(container):
        tag = local_name(math_elem.tag)
        if tag == "oMath":
            appended = True
            try:
                segments.append(_build_inline_math_segment(math_elem))
            except Exception:
                fallback = normalize_text(
                    " ".join(node.text or "" for node in math_elem.findall(".//m:t", NS))
                )
                segments.append(
                    ParagraphSegment(kind="math_error", text=fallback or "[unsupported-math]")
                )
            continue
        if tag == "oMathPara":
            appended = True
            try:
                segments.append(_build_block_math_segment(math_elem))
            except Exception:
                fallback = normalize_text(
                    " ".join(node.text or "" for node in math_elem.findall(".//m:t", NS))
                )
                segments.append(
                    ParagraphSegment(kind="math_error", text=fallback or "[unsupported-math]")
                )
            continue
    return appended


def parse_paragraph_segments(paragraph: ET.Element) -> List[ParagraphSegment]:
    segments: List[ParagraphSegment] = []
    for child in list(paragraph):
        tag = local_name(child.tag)
        if tag in {"r", "fld"}:
            text = _extract_run_text(child)
            _append_text_segment(segments, text)
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


# 주변 리스트 문맥에 맞는 짧은 일반 텍스트를 리스트 항목으로 승격한다.
def promote_plain_text_to_list(
    blocks: Sequence[ShapeBlock],
) -> List[ShapeBlock]:
    if not any(block.kind in {"list_ul", "list_ol"} for block in blocks):
        return list(blocks)

    text_blocks = [
        (idx, block.plain_text) for idx, block in enumerate(blocks) if block.kind == "text"
    ]
    if len(text_blocks) < 3:
        return list(blocks)

    first_list_kind = next(
        (block.kind for block in blocks if block.kind in {"list_ul", "list_ol"}), "list_ul"
    )
    promoted = list(blocks)
    for idx, text in text_blocks:
        if len(text) > 80 or text.endswith((".", ":")):
            continue
        promoted[idx] = ShapeBlock(kind=first_list_kind, segments=promoted[idx].segments, level=0)
    return promoted


# PowerPoint의 불연속 목록 레벨을 0,1,2... 형태의 연속 깊이로 정규화한다.
# 원본 lvl 값이 듬성듬성해도 Markdown 렌더링 들여쓰기가 과도하게 깊어지지 않도록 막는다.
def normalize_list_levels(blocks: Sequence[ShapeBlock]) -> List[ShapeBlock]:
    levels = sorted(
        {int(block.level or 0) for block in blocks if block.kind in {"list_ul", "list_ol"}}
    )
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
    for p in shape_elem.findall("./p:txBody/a:p", NS):
        segments = parse_paragraph_segments(p)
        if not segments:
            continue
        plain_text = paragraph_text(p)
        if not plain_text and not any(
            segment.kind in {"math_inline", "math_block"} for segment in segments
        ):
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


def _ooxml_part_from_relationship(
    rels_path: Path,
    target: str,
) -> Optional[str]:
    normalized_target = posixpath.normpath(target.replace("\\", "/"))
    target_parts = normalized_target.split("/")
    if "ppt" in target_parts:
        ppt_index = target_parts.index("ppt")
        return "/".join(target_parts[ppt_index:])

    source_part = _part_path_from_rels_path(rels_path)
    if source_part is None:
        return None
    return _resolve_ooxml_target(source_part, target)


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

    chart_part_path = _ooxml_part_from_relationship(rels_path, target)
    if chart_part_path is None:
        return None, f"chart conversion failed: unsupported rels path: {rels_path}"

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

    dm_rel_ids = graphic_frame.find("./a:graphic/a:graphicData/dgm:relIds", NS)
    dm_rid = dm_rel_ids.attrib.get(f"{{{NS['r']}}}dm") if dm_rel_ids is not None else None
    target = rels_map.get(dm_rid) if dm_rid else None
    if not target:
        return None, "smartart conversion failed: missing diagram relationship target"

    if rels_path is None:
        return None, "smartart conversion failed: missing slide rels file"
    data_part_path = _ooxml_part_from_relationship(rels_path, target)
    if data_part_path is None:
        return None, f"smartart conversion failed: unsupported rels path: {rels_path}"

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
def diagram_data_path(
    graphic_frame: ET.Element, rels_path: Optional[Path], rels_map: Dict[str, str]
) -> Optional[Path]:
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


# 테이블 오버레이 이미지를 정적 asset 링크로 변환한다.
def overlay_content_text(
    path: str,
    output_dir: Optional[Path],
    media_dir: Optional[Path] = None,
    copied_media: Optional[Dict[str, Path]] = None,
) -> str:
    return overlay_link_text(
        path,
        output_dir,
        media_dir=media_dir,
        copied_media=copied_media,
    )


# 슬라이드 전체에서 테이블 위에 겹쳐진 picture overlay들을 수집한다.
# 실제 구현은 table_overlay 모듈에 두고, 여기서는 현재 파일의 namespace와
# 이미지 경로 해석 함수를 주입하는 어댑터 역할만 맡는다.
def collect_table_overlay_pictures(
    shape_items: Sequence[Dict[str, object]],
    slide_xml: Path,
    rels_map: Dict[str, str],
    rels_path: Optional[Path],
) -> Tuple[Dict[str, List[Dict[str, object]]], set[str], List[str], int, int]:
    _ = slide_xml
    return collect_table_overlay_pictures_core(
        shape_items,
        rels_map,
        rels_path,
        ns=NS,
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
) -> Tuple[Optional[str], Optional[str]]:
    return convert_table_to_markdown_core(
        graphic_frame,
        overlays=overlays,
        output_dir=output_dir,
        media_dir=media_dir,
        copied_media=copied_media,
        rels_path=rels_path,
        rels_map=rels_map,
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
        load_effective_properties=load_effective_properties,
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
    )


# 슬라이드 XML 하나를 Markdown과 통계 정보로 변환하는 진입점이다.
# 실제 본문 순회는 core 구현이 담당하고, 이 함수는 현재 모듈의 정책과 옵션을 연결한다.
def convert_one_slide(
    *,
    context: SlideConversionContext,
    assets: SlideRenderAssets,
    strict_headings: bool = False,
    heading_mode: str = "auto",
) -> Tuple[SlideDocument, SlideStats]:
    context.heading_mode = heading_mode
    return convert_one_slide_core(
        context=context,
        assets=assets,
        strict_headings=strict_headings,
        deps=_slide_conversion_deps(),
    )


# CLI 인자를 정의하고 파싱한다.
# 입력 PPTX 목록과 heading/placeholder 정책 등
# 전체 변환 파이프라인을 제어하는 설정을 여기서 받는다.
def _parse_args() -> argparse.Namespace:
    # parser 생성
    parser = argparse.ArgumentParser(description="Convert PPTX/PPT files to Markdown or JSON.")
    # 필요한 인자들 추가
    parser.add_argument(
        "inputs",
        nargs="*",
        help=(
            "Optional .pptx/.ppt selections (e.g., sample3.pptx sample4.ppt). "
            "If omitted, all .pptx/.ppt files in the current directory are processed."
        ),
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        default=None,
        help="Directory for converted output. Default: ./output",
    )
    parser.add_argument(
        "--output-format",
        choices=("markdown", "json"),
        default="markdown",
        help="Final output format. Default: markdown.",
    )
    parser.add_argument(
        "--work-dir",
        default=None,
        help=(
            "Directory for intermediate files (extracted packages, caches, "
            "structure-analysis artifacts). Default: ./.pptx2markdown"
        ),
    )
    parser.add_argument(
        "--headings",
        choices=("auto", "strict"),
        default="auto",
        help=(
            "Heading detection strategy. auto uses placeholder, numbering, font, "
            "and position signals; strict uses title placeholders only."
        ),
    )
    parser.add_argument(
        "--placeholder-inheritance",
        dest="pptx_inheritance",
        choices=("none", "geometry", "style"),
        default="style",
        help=(
            "Placeholder inheritance depth for markdown extraction. "
            "none uses slide XML only; geometry inherits placeholder type/bbox; "
            "style also inherits text style signals such as font size and list semantics."
        ),
    )
    parser.add_argument(
        "--inherited-shapes",
        choices=("none", "visible", "all"),
        default="visible",
        help=(
            "Materialize layout/master-only shapes into effective structure properties. "
            "visible keeps slideshow-visible text/images while filtering placeholder prompts; "
            "all keeps every shape."
        ),
    )
    parser.add_argument(
        "--ppt-converter",
        choices=("auto", "powerpoint", "libreoffice"),
        default="auto",
        help=(
            "Converter used for legacy .ppt inputs. auto tries PowerPoint on Windows, "
            "then LibreOffice."
        ),
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable verbose debug logging.",
    )
    return parser.parse_args()


# CLI 인자를 내부 ConverterConfig로 정규화한다.
def _build_config(args: argparse.Namespace) -> ConverterConfig:
    work_root = (
        Path(args.work_dir).expanduser().resolve()
        if getattr(args, "work_dir", None)
        else (Path.cwd() / ".pptx2markdown").resolve()
    )
    output_root = (
        Path(args.output_dir).expanduser().resolve()
        if getattr(args, "output_dir", None)
        else (Path.cwd() / "output").resolve()
    )
    heading_mode = str(args.headings)
    if heading_mode not in {"auto", "strict"}:
        raise ValueError("headings must be 'auto' or 'strict'")
    output_format = str(getattr(args, "output_format", "markdown"))
    if output_format not in {"markdown", "json"}:
        raise ValueError("output_format must be 'markdown' or 'json'")
    pptx_inheritance = str(args.pptx_inheritance)
    if pptx_inheritance not in {"none", "geometry", "style"}:
        raise ValueError("placeholder_inheritance must be 'none', 'geometry', or 'style'")
    inherited_shapes = str(args.inherited_shapes)
    if inherited_shapes not in {"none", "visible", "all"}:
        raise ValueError("inherited_shapes must be 'none', 'visible', or 'all'")
    return ConverterConfig(
        cwd=work_root,
        output_dir=output_root,
        inputs=list(args.inputs),
        output_format=output_format,
        heading_mode=heading_mode,
        strict=heading_mode == "strict",
        pptx_inheritance=pptx_inheritance,
        inherited_shapes=inherited_shapes,
        ppt_converter=str(args.ppt_converter),
    )


# 실제 처리에 사용할 입력 패키지 경로 목록을 확정한다.
# 사용자가 직접 넘긴 입력이 있으면 그것만 검증/추출하고,
# 없으면 target_pptx 아래 파일들을 자동 탐색해서 동일한 형식으로 준비한다.
def _resolve_prepared_inputs(config: ConverterConfig) -> List[PreparedPackage]:
    # 누락된 입력에 대해 어떤 경로들을 확인했는지 자세히 로그로 남긴다.
    def _log_missing_inputs(missing_inputs: Sequence[Dict[str, object]]) -> None:
        if not missing_inputs:
            return
        logger.error("Input .pptx/.ppt file not found.")
        for row in missing_inputs:
            requested = str(row.get("input", "")).strip()
            if requested:
                logger.error("- requested: %s", requested)
            checked = row.get("checked")
            if isinstance(checked, list):
                for candidate in checked:
                    logger.error("  checked: %s", candidate)

    if config.inputs:
        unsupported_inputs = [
            x
            for x in config.inputs
            if Path(x).suffix and Path(x).suffix.lower() not in {".pptx", ".ppt"}
        ]
        if unsupported_inputs:
            logger.error("Only .pptx/.ppt inputs are allowed.")
            logger.error("Provide files like: sample1.pptx sample2.ppt")
            for item in unsupported_inputs:
                logger.error("- %s", item)
            raise ValueError("invalid presentation inputs")
        prepared_inputs, missing_inputs = prepare_package_inputs(
            config.cwd,
            config.inputs,
            force_extract=True,
            ppt_converter=config.ppt_converter,
        )
        if missing_inputs:
            _log_missing_inputs(missing_inputs)
            raise ValueError("missing presentation inputs")
        return prepared_inputs

    auto_presentation_inputs = collect_target_presentation_inputs()
    if not auto_presentation_inputs:
        logger.info("No .pptx/.ppt files found in: %s", Path.cwd().resolve())
        return []
    prepared_inputs, missing_inputs = prepare_package_inputs(
        config.cwd,
        auto_presentation_inputs,
        force_extract=True,
        ppt_converter=config.ppt_converter,
    )
    if missing_inputs:
        _log_missing_inputs(missing_inputs)
        raise ValueError("missing auto-discovered presentation inputs")
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


# 추출된 패키지 하나를 변환하고 결과와 manifest를 누적한다.
def _convert_package(
    config: ConverterConfig,
    package: PreparedPackage,
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
    output_path = (
        package.output_json_path(config.output_dir)
        if config.output_format == "json"
        else package.output_markdown_path(config.output_dir)
    )
    media_dir = pkg_out / "media"
    copied_media: Dict[str, Path] = {}
    pkg_out.mkdir(parents=True, exist_ok=True)

    pkg_row: Dict[str, object] = {
        "package": str(pkg),
        "name": pkg_name,
        "slides": [],
        "result": str(output_path),
        "output_format": config.output_format,
    }
    pkg_row[f"result_{'md' if config.output_format == 'markdown' else 'json'}"] = str(output_path)

    slides: List[SlideDocument] = []
    try:
        ro_map, ro_output = run_structure_analysis_stage(
            work_root=config.cwd,
            package_name=pkg_name,
            slide_xmls=slide_xmls,
            strict=config.strict,
            pptx_inheritance=config.pptx_inheritance,
            inherited_shapes=config.inherited_shapes,
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

    for i, slide_xml in enumerate(slide_xmls, 1):
        page_no = parse_slide_number(slide_xml.name, i)
        ordered_slide_xml = ro_map.get(str(slide_xml.resolve()), slide_xml)
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
                source_pptx_path=package.source_pptx_path,
            )
            assets = SlideRenderAssets(
                output_dir=pkg_out,
                media_dir=media_dir,
                copied_media=copied_media,
            )
            slide_document, stats = convert_one_slide(
                context=context,
                assets=assets,
                strict_headings=config.strict,
                heading_mode=config.heading_mode,
            )
            slides.append(slide_document)

            row.update({"status": "ok", **stats.to_slide_row_fields()})
            manifest.summary.add_slide(stats)
            logger.info("[%s] [convert] Processed: %s", pkg_name, slide_xml.name)
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

    document = PresentationDocument(
        source=str(package.source_pptx_path),
        slides=slides,
    )
    if config.output_format == "json":
        output_text = json.dumps(document.model_dump(mode="json"), ensure_ascii=False, indent=2)
        output_text += "\n"
    else:
        output_text = render_presentation_markdown(document)
    output_path.write_text(output_text, encoding="utf-8")
    manifest.summary.processed_packages += 1
    manifest.packages.append(pkg_row)


# 전체 변환 파이프라인의 CLI 엔트리포인트다.
# 인자 파싱, 입력 준비, 패키지 선택, manifest 저장,
# 최종 요약 로그 출력까지 전체 흐름을 조율한다.
def main() -> int:
    args = _parse_args()
    _configure_logging(verbose=bool(getattr(args, "verbose", False)))
    config = _build_config(args)
    return run(config)


# CLI와 Python API가 공유하는 실행 본체.
# ConverterConfig 하나를 받아 전체 파이프라인을 수행하고 종료 코드를 반환한다.
def run(config: ConverterConfig) -> int:
    global _WORK_ROOT
    _WORK_ROOT = config.cwd
    config.output_dir.mkdir(parents=True, exist_ok=True)

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

    manifest = ConversionManifest()

    # 각 패키지에 대하여 일괄적으로 메인 컨버터 로직인 _convert_package를 수행한다.
    for pkg in packages:
        _convert_package(config, pkg, manifest)

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
