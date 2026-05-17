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


def apply_font(run, spec: FontSpec, fill_missing_only: bool = False) -> None:
    """Set the four font slots on `run`.

    fill_missing_only=False (default): full override — all four slots get
    written from `spec`. This is what we want when rendering with the
    builtin template, where we own the styling end-to-end.

    fill_missing_only=True: respect whatever the run/style already has;
    only write a slot if it is currently unset. This is what we want when
    the user passed their own template: we don't want to clobber their
    carefully chosen fonts, but we do want to backfill the eastAsia slot
    if they forgot it (the classic CJK-renders-in-Calibri bug).

    Size is only written when the run has no size set (under
    fill_missing_only) or unconditionally (under full override) and `spec`
    actually carries one.
    """
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)

    def _set(slot: str, value: str) -> None:
        if value is None:
            return
        if fill_missing_only and rfonts.get(qn(f"w:{slot}")):
            return
        rfonts.set(qn(f"w:{slot}"), value)

    _set("ascii",    spec.latin)
    _set("hAnsi",    spec.latin)
    _set("eastAsia", spec.cjk)
    _set("cs",       spec.latin)

    if spec.size_pt is not None:
        if not (fill_missing_only and run.font.size is not None):
            run.font.size = Pt(spec.size_pt)


def style_has_eastasia_font(style) -> bool:
    """Return True if `style` (a docx Style) declares a w:eastAsia font.

    Used to decide whether a user template's heading/normal style already
    handles CJK on its own, or whether we need to backfill at the run level.
    """
    elem = style.element if hasattr(style, "element") else style._element
    rpr = elem.find(qn("w:rPr"))
    if rpr is None:
        return False
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        return False
    return bool(rfonts.get(qn("w:eastAsia")))
