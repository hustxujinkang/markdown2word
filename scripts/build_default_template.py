"""Generate md2word's default Chinese-friendly template.

Run this once; commit the produced `md2word/templates/default_zh.docx`
into the repo as a distribution asset.

Design notes — what's special about this template:

* Every style has `w:rFonts` with an explicit `w:eastAsia` slot, so Chinese
  characters don't fall back to the host machine's default CJK font.
* Body paragraphs use `firstLineChars=200` (two characters) rather than a
  fixed-width indent — this auto-scales with font size, which is how
  properly-typeset Chinese documents work.
* A `Code Block` paragraph style is created (Word doesn't ship one).
* Page size is A4 with Chinese-convention margins.
"""
from __future__ import annotations

import os

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


LATIN_BODY    = "Times New Roman"
LATIN_HEADING = "Arial"
LATIN_CODE    = "Consolas"
CJK_BODY      = "宋体"
CJK_HEADING   = "黑体"
CJK_CODE      = "宋体"


def _set_rfonts(rpr, latin: str, cjk: str) -> None:
    """Set all four font slots — the CJK fix that matters."""
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:ascii"),    latin)
    rfonts.set(qn("w:hAnsi"),    latin)
    rfonts.set(qn("w:eastAsia"), cjk)
    rfonts.set(qn("w:cs"),       latin)


def _style_rpr(style):
    """Get-or-create the style's run-properties element."""
    elem = style.element
    rpr = elem.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        elem.append(rpr)
    return rpr


def _style_ppr(style):
    elem = style.element
    ppr = elem.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        # pPr must come before rPr in a style element
        rpr = elem.find(qn("w:rPr"))
        if rpr is not None:
            elem.insert(list(elem).index(rpr), ppr)
        else:
            elem.append(ppr)
    return ppr


def _set_first_line_chars(style, chars: int = 2) -> None:
    """Indent first line by N Chinese characters (scales with font size)."""
    ppr = _style_ppr(style)
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    ind.set(qn("w:firstLineChars"), str(chars * 100))
    ind.set(qn("w:firstLine"), "0")   # required alongside firstLineChars


def _set_line_spacing(style, multiple: float) -> None:
    ppr = _style_ppr(style)
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    spacing.set(qn("w:line"), str(int(multiple * 240)))
    spacing.set(qn("w:lineRule"), "auto")


def _set_spacing_before_after(style, before_pt: float, after_pt: float) -> None:
    ppr = _style_ppr(style)
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    spacing.set(qn("w:before"), str(int(before_pt * 20)))
    spacing.set(qn("w:after"),  str(int(after_pt  * 20)))


def _set_shading(style, fill_hex: str) -> None:
    ppr = _style_ppr(style)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    ppr.append(shd)


def _set_left_border(style, color_hex: str, size: int = 24) -> None:
    ppr = _style_ppr(style)
    pBdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), str(size))
    left.set(qn("w:space"), "8")
    left.set(qn("w:color"), color_hex)
    pBdr.append(left)
    ppr.append(pBdr)


def build_template(out_path: str) -> None:
    doc = Document()

    # ------------ page setup (A4 with Chinese-convention margins) ----------
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width  = Cm(21.0)
    section.top_margin    = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin   = Cm(3.18)
    section.right_margin  = Cm(3.18)

    # ------------ Normal (body text) ----------
    normal = doc.styles["Normal"]
    normal.font.size = Pt(12)                              # 小四
    _set_rfonts(_style_rpr(normal), LATIN_BODY, CJK_BODY)
    _set_line_spacing(normal, 1.5)
    _set_first_line_chars(normal, 2)
    _set_spacing_before_after(normal, 0, 0)

    # ------------ Headings 1..6 ----------
    # All six levels get the same heading-font treatment. Earlier
    # versions only configured 1..4 and left 5/6 inheriting Word's
    # default, which left the eastAsia slot empty — Chinese text in
    # H5/H6 would fall back to Calibri. `md2word inspect` will flag
    # exactly this kind of gap on user templates.
    heading_sizes = {1: 16, 2: 15, 3: 14, 4: 12, 5: 11, 6: 10}      # 三号 / 小三 / 四号 / 小四 / 五号 / 小五
    heading_spacing = {1: (18, 12), 2: (16, 10), 3: (14, 8), 4: (12, 6), 5: (10, 4), 6: (8, 4)}
    for level, size in heading_sizes.items():
        style = doc.styles[f"Heading {level}"]
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        _set_rfonts(_style_rpr(style), LATIN_HEADING, CJK_HEADING)
        _set_line_spacing(style, 1.5)
        before, after = heading_spacing[level]
        _set_spacing_before_after(style, before, after)
        # explicitly kill the first-line-indent that Normal applies
        ppr = _style_ppr(style)
        ind = ppr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind"); ppr.append(ind)
        ind.set(qn("w:firstLineChars"), "0")
        ind.set(qn("w:firstLine"), "0")

    # ------------ List styles ----------
    for sty_name in ("List Bullet", "List Number"):
        style = doc.styles[sty_name]
        style.font.size = Pt(12)
        _set_rfonts(_style_rpr(style), LATIN_BODY, CJK_BODY)
        _set_line_spacing(style, 1.5)
        # lists shouldn't have first-line indent
        ppr = _style_ppr(style)
        ind = ppr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind"); ppr.append(ind)
        ind.set(qn("w:firstLineChars"), "0")

    # ------------ Quote ----------
    quote = doc.styles["Quote"]
    quote.font.size = Pt(12)
    quote.font.italic = False                              # Chinese italic looks awful
    _set_rfonts(_style_rpr(quote), LATIN_BODY, CJK_BODY)
    _set_line_spacing(quote, 1.5)
    _set_spacing_before_after(quote, 6, 6)
    _set_left_border(quote, "4A90E2", size=24)
    ppr = _style_ppr(quote)
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind"); ppr.append(ind)
    ind.set(qn("w:left"), "480")                           # 0.33 inch
    ind.set(qn("w:firstLineChars"), "0")
    ind.set(qn("w:firstLine"), "0")

    # ------------ Code Block (new style) ----------
    code = doc.styles.add_style("Code Block", WD_STYLE_TYPE.PARAGRAPH)
    code.base_style = doc.styles["Normal"]
    code.font.size = Pt(10)                                # 小五
    _set_rfonts(_style_rpr(code), LATIN_CODE, CJK_CODE)
    _set_line_spacing(code, 1.15)
    _set_spacing_before_after(code, 6, 6)
    _set_shading(code, "F5F5F5")
    ppr = _style_ppr(code)
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind"); ppr.append(ind)
    ind.set(qn("w:firstLineChars"), "0")
    ind.set(qn("w:firstLine"), "0")
    ind.set(qn("w:left"), "240")

    # ------------ Caption ----------
    caption = doc.styles["Caption"]
    caption.font.size = Pt(9)                              # 小五
    caption.font.italic = False
    caption.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    _set_rfonts(_style_rpr(caption), LATIN_BODY, CJK_BODY)
    _set_line_spacing(caption, 1.15)
    _set_spacing_before_after(caption, 2, 12)
    ppr = _style_ppr(caption)
    jc = ppr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc"); ppr.append(jc)
    jc.set(qn("w:val"), "center")
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind"); ppr.append(ind)
    ind.set(qn("w:firstLineChars"), "0")

    # ------------ Enable CJK/Latin auto-spacing in document settings ----------
    settings = doc.settings.element
    for tag in ("autoSpaceDE", "autoSpaceDN"):
        el = settings.find(qn(f"w:{tag}"))
        if el is None:
            el = OxmlElement(f"w:{tag}"); settings.append(el)
        el.set(qn("w:val"), "1")

    # save (empty body — only styles matter)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    doc.save(out_path)
    print(f"✓ 模板已生成: {out_path}")


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    out = os.path.join(here, "..", "md2word", "templates", "default_zh.docx")
    build_template(os.path.abspath(out))
