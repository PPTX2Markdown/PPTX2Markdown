"""OMML-to-LaTeX conversion backed by omml2latex."""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET

from omml2latex import convert_omml


def omml_to_latex(element: ET.Element) -> str:
    """Convert m:oMath or m:oMathPara to single-line LaTeX without $ delimiters.

    Returns an empty string when the equation cannot be converted.
    """

    try:
        latex = convert_omml(element)
    except Exception:  # third-party converter: callers report empty results as warnings
        return ""
    latex = re.sub(r"\s+", " ", latex).strip()
    for delimiter in ("$$", "$"):
        if (
            len(latex) >= 2 * len(delimiter)
            and latex.startswith(delimiter)
            and latex.endswith(delimiter)
        ):
            return latex[len(delimiter) : -len(delimiter)].strip()
    return latex
