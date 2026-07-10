from __future__ import annotations

import re
from pathlib import Path
from typing import List, Optional, Tuple
import xml.etree.ElementTree as ET

from .constants import LARGE_INT, NS


def local_name(tag: str) -> str:
    return tag.split("}", 1)[-1]


def natural_key(path: Path) -> Tuple[object, ...]:
    parts = re.split(r"(\d+)", path.name)
    out: List[object] = []
    for part in parts:
        if part.isdigit():
            out.append(int(part))
        else:
            out.append(part.lower())
    return tuple(out)


def parse_int(value: Optional[str], default: int = LARGE_INT) -> int:
    if value is None:
        return default
    try:
        return int(value)
    except ValueError:
        return default


def get_nvpr_paths(tag: str) -> Tuple[str, str]:
    if tag == "sp":
        return "./p:nvSpPr/p:cNvPr", "./p:nvSpPr/p:nvPr/p:ph"
    if tag == "pic":
        return "./p:nvPicPr/p:cNvPr", "./p:nvPicPr/p:nvPr/p:ph"
    if tag == "graphicFrame":
        return "./p:nvGraphicFramePr/p:cNvPr", "./p:nvGraphicFramePr/p:nvPr/p:ph"
    if tag == "grpSp":
        return "./p:nvGrpSpPr/p:cNvPr", "./p:nvGrpSpPr/p:nvPr/p:ph"
    if tag == "cxnSp":
        return "./p:nvCxnSpPr/p:cNvPr", "./p:nvCxnSpPr/p:nvPr/p:ph"
    return ".//p:cNvPr", ".//p:ph"


def first_off(elem: ET.Element) -> Optional[ET.Element]:
    for path in (
        "./p:spPr/a:xfrm/a:off",
        "./p:grpSpPr/a:xfrm/a:off",
        "./p:xfrm/a:off",
        ".//a:off",
    ):
        off = elem.find(path, NS)
        if off is not None:
            return off
    return None


def first_ext(elem: ET.Element) -> Optional[ET.Element]:
    for path in (
        "./p:spPr/a:xfrm/a:ext",
        "./p:grpSpPr/a:xfrm/a:ext",
        "./p:xfrm/a:ext",
        ".//a:ext",
    ):
        ext = elem.find(path, NS)
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
    if any(v >= LARGE_INT for v in (x, y, w, h)):
        return None
    if w <= 0 or h <= 0:
        return None
    return (x, y, x + w, y + h)


def register_xml_namespaces() -> None:
    ET.register_namespace("a", NS["a"])
    ET.register_namespace("p", NS["p"])
    ET.register_namespace("r", NS["r"])
