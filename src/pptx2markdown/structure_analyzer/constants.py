from __future__ import annotations

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

REL_NS = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}

SLIDE_LAYOUT_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
)
SLIDE_MASTER_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
)

REORDERABLE = {"sp", "pic", "graphicFrame", "grpSp", "cxnSp"}
FOOTER_TYPES = {"sldNum", "ftr", "dt"}
TITLE_TYPES = {"title", "ctrTitle", "subTitle"}
STRICT_HEADING_PLACEHOLDER_TYPES = {"title", "ctrTitle", "subTitle"}

LARGE_INT = 10**18
