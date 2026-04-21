"""CJK-aware font application.

The hidden bug in 90% of python-docx projects: `run.font.name = ...` only
sets the `w:ascii` and `w:hAnsi` slots. CJK characters fall back to the
`w:eastAsia` slot, which stays on whatever the template defaults to. Use
`apply_font` instead — it sets all three explicitly.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt


@dataclass
class FontSpec:
    latin: str
    cjk: str
    size_pt: Optional[float] = None


# Sensible defaults — users can override via FontProfile.
DEFAULTS = {
    "body":    FontSpec(latin="Times New Roman", cjk="宋体",   size_pt=11),
    "heading": FontSpec(latin="Arial",           cjk="黑体",   size_pt=None),
    "code":    FontSpec(latin="Consolas",        cjk="宋体",   size_pt=10),
}


def apply_font(run, spec: FontSpec) -> None:
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"),    spec.latin)
    rfonts.set(qn("w:hAnsi"),    spec.latin)
    rfonts.set(qn("w:eastAsia"), spec.cjk)
    rfonts.set(qn("w:cs"),       spec.latin)
    if spec.size_pt is not None:
        run.font.size = Pt(spec.size_pt)
