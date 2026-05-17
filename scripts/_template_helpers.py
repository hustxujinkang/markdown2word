"""Shared helpers for template-build scripts.

Each `scripts/build_<name>.py` produces one .docx template under
`md2word/templates/`. The helpers in this module wrap the raw OOXML
fiddling that python-docx doesn't expose nicely — font slot setting,
fixed line-height, paragraph borders, etc.

Why a private module under `scripts/` rather than inside `md2word/`?
These helpers are only used at template-build time, never at
runtime. Keeping them outside the runtime package makes the
dependency direction obvious: build scripts depend on python-docx +
this module; the runtime package depends on neither.
"""
from __future__ import annotations

import os

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt


# ---------- rPr / pPr accessors -------------------------------------------

def style_rpr(style):
    """Get-or-create the style's run-properties element."""
    elem = style.element
    rpr = elem.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        elem.append(rpr)
    return rpr


def style_ppr(style):
    """Get-or-create the style's paragraph-properties element.

    pPr must come before rPr in a style element, so we insert it
    accordingly when both need to exist.
    """
    elem = style.element
    ppr = elem.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        rpr = elem.find(qn("w:rPr"))
        if rpr is not None:
            elem.insert(list(elem).index(rpr), ppr)
        else:
            elem.append(ppr)
    return ppr


# ---------- font slot setting ---------------------------------------------

def set_rfonts(rpr, latin: str, cjk: str) -> None:
    """Set all four font slots on an rPr element.

    The CJK fix that matters: w:eastAsia is what Chinese text actually
    resolves through. python-docx's run.font.name only writes ascii/
    hAnsi. Without eastAsia, Chinese falls back to the host machine's
    default CJK font — usually Calibri on Windows, which can't render
    CJK at all.

    Also drop any `*Theme` attributes (asciiTheme, eastAsiaTheme, etc.).
    Theme references take priority over direct font names in Word's
    resolution order — if both are present, Word looks up the font in
    the theme XML, which python-docx's blank template configures as
    日文 (MS Gothic / MS Mincho). Removing the Theme refs forces Word
    to use our explicit `w:eastAsia` value.
    """
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    # Strip Theme attributes — they win over direct font names.
    for theme_attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        attr_name = qn(f"w:{theme_attr}")
        if rfonts.get(attr_name) is not None:
            del rfonts.attrib[attr_name]
    rfonts.set(qn("w:ascii"),    latin)
    rfonts.set(qn("w:hAnsi"),    latin)
    rfonts.set(qn("w:eastAsia"), cjk)
    rfonts.set(qn("w:cs"),       latin)


def configure_style(style, *, latin: str, cjk: str, size_pt: float = None,
                    bold: bool = False, italic: bool = False) -> None:
    """One-shot: font slots + size + bold/italic on a paragraph style."""
    if size_pt is not None:
        style.font.size = Pt(size_pt)
    style.font.bold = bold
    style.font.italic = italic
    set_rfonts(style_rpr(style), latin, cjk)


# ---------- spacing & indentation ------------------------------------------

def set_line_spacing_multiple(style, multiple: float) -> None:
    """Multiple-of-single line spacing (e.g. 1.5 = "1.5 倍行距")."""
    ppr = style_ppr(style)
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    spacing.set(qn("w:line"), str(int(multiple * 240)))
    spacing.set(qn("w:lineRule"), "auto")


def set_line_spacing_fixed_pt(style, line_pt: float) -> None:
    """Exact-pt line height — required by GB/T 9704 (公文固定 28-30 磅行距).

    w:line is in 1/20 pt units when lineRule=exact, which is why we
    multiply by 20. Don't confuse with the *240 multiplier used for
    'multiple' line rules — that one is in 240ths-of-a-line.
    """
    ppr = style_ppr(style)
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    spacing.set(qn("w:line"), str(int(line_pt * 20)))
    spacing.set(qn("w:lineRule"), "exact")


def set_spacing_before_after(style, before_pt: float, after_pt: float) -> None:
    ppr = style_ppr(style)
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    spacing.set(qn("w:before"), str(int(before_pt * 20)))
    spacing.set(qn("w:after"),  str(int(after_pt  * 20)))


def set_first_line_chars(style, chars: int = 2) -> None:
    """Indent first line by N Chinese characters (scales with font size).

    Using firstLineChars rather than a fixed-width firstLine indent is
    the convention for Chinese typesetting — when the font size
    changes, the indent rescales automatically.
    """
    ppr = style_ppr(style)
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    ind.set(qn("w:firstLineChars"), str(chars * 100))
    ind.set(qn("w:firstLine"), "0")


def clear_first_line_indent(style) -> None:
    """Force no first-line indent (for headings, lists, code, captions
    that inherit from Normal but shouldn't be indented)."""
    ppr = style_ppr(style)
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    ind.set(qn("w:firstLineChars"), "0")
    ind.set(qn("w:firstLine"), "0")


def set_left_indent_chars(style, chars: int) -> None:
    """Hanging-style left indent in character units."""
    ppr = style_ppr(style)
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    ind.set(qn("w:leftChars"), str(chars * 100))


def set_alignment(style, align: str) -> None:
    """align: left / center / right / both (= justify)."""
    ppr = style_ppr(style)
    jc = ppr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        ppr.append(jc)
    jc.set(qn("w:val"), align)


# ---------- visual decorations ---------------------------------------------

def set_shading(style, fill_hex: str) -> None:
    """Solid background fill (e.g. for code blocks)."""
    ppr = style_ppr(style)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    ppr.append(shd)


def set_left_border(style, color_hex: str, size: int = 24) -> None:
    """Single left border (e.g. for the bar on a Quote block)."""
    ppr = style_ppr(style)
    pBdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), str(size))
    left.set(qn("w:space"), "8")
    left.set(qn("w:color"), color_hex)
    pBdr.append(left)
    ppr.append(pBdr)


# ---------- document-level settings ----------------------------------------

def set_default_eastasia_language_zh_cn(doc) -> None:
    """Set the document's default eastAsia language to zh-CN.

    Without this, the docDefaults block stays at the python-docx default
    (`w:eastAsia="en-US"`), and Word's CJK font resolution falls back to
    日文 fonts (MS Gothic / MS Mincho on Windows). With zh-CN set, the
    fallback chain points to 简体中文 fonts (SimSun / SimHei etc.).

    This is in addition to setting w:eastAsia on each style's rFonts —
    belt and suspenders, since the lang attribute and the font slot are
    independent inputs to Word's font-resolution.
    """
    styles_part = doc.styles.element
    # Find or create docDefaults > rPrDefault > rPr > lang
    doc_defaults = styles_part.find(qn("w:docDefaults"))
    if doc_defaults is None:
        doc_defaults = OxmlElement("w:docDefaults")
        styles_part.insert(0, doc_defaults)
    rpr_default = doc_defaults.find(qn("w:rPrDefault"))
    if rpr_default is None:
        rpr_default = OxmlElement("w:rPrDefault")
        doc_defaults.append(rpr_default)
    rpr = rpr_default.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        rpr_default.append(rpr)
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr.append(lang)
    lang.set(qn("w:val"),      "zh-CN")
    lang.set(qn("w:eastAsia"), "zh-CN")
    lang.set(qn("w:bidi"),     "ar-SA")


def set_style_eastasia_language_zh_cn(style) -> None:
    """Tag a specific style's rPr with eastAsia=zh-CN.

    Useful when one style needs a different language than the doc
    default — most of the time set_default_eastasia_language_zh_cn() on
    the doc is enough.
    """
    rpr = style_rpr(style)
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr.append(lang)
    lang.set(qn("w:val"),      "zh-CN")
    lang.set(qn("w:eastAsia"), "zh-CN")
    lang.set(qn("w:bidi"),     "ar-SA")


def enable_cjk_latin_autospace(doc) -> None:
    """Turn on the doc-level flags that auto-space CJK against Latin/digits.

    Without these, "Python 是 the 默认" runs together with no visual
    breathing room between scripts. Word does this for you when these
    settings are on; with them off, you have to type spaces yourself.
    """
    settings = doc.settings.element
    for tag in ("autoSpaceDE", "autoSpaceDN"):
        el = settings.find(qn(f"w:{tag}"))
        if el is None:
            el = OxmlElement(f"w:{tag}"); settings.append(el)
        el.set(qn("w:val"), "1")


def patch_theme_cjk_fonts(docx_path: str, major_ea: str = "黑体",
                          minor_ea: str = "宋体") -> None:
    """Post-save fix: fill the empty CJK slots in word/theme/theme1.xml.

    python-docx ships its blank template with theme1.xml's CJK font
    fields (`<a:ea typeface=""/>`) blank. Any style that carries an
    `eastAsiaTheme="majorEastAsia"` (which is the majority of Word's
    built-in styles, like the Table Grid family, No Spacing, etc.)
    resolves through the theme — and because ea is empty, Word falls
    back through its language chain to whatever the OS prefers, often
    MS Gothic (日文) on default-locale Windows installs.

    Filling these two slots in the theme is one-line operation that
    fixes the issue for every style that uses theme refs, without us
    having to strip Theme attributes from each style manually.

    Must be called after doc.save() because python-docx doesn't expose
    the theme part as a first-class object.
    """
    import shutil
    import re
    import tempfile
    import zipfile

    # Read the existing theme, patch it, write back. Done by copying the
    # whole zip into a new file (zipfile doesn't support in-place edit
    # of a single member).
    fd, tmp_path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    try:
        with zipfile.ZipFile(docx_path, "r") as zin, \
             zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/theme/theme1.xml":
                    text = data.decode("utf-8")
                    # Fill ea slot inside the <a:majorFont> block
                    text = re.sub(
                        r'(<a:majorFont>.*?<a:ea typeface=")[^"]*(")',
                        rf'\1{major_ea}\2',
                        text, count=1, flags=re.DOTALL,
                    )
                    # And inside <a:minorFont>
                    text = re.sub(
                        r'(<a:minorFont>.*?<a:ea typeface=")[^"]*(")',
                        rf'\1{minor_ea}\2',
                        text, count=1, flags=re.DOTALL,
                    )
                    data = text.encode("utf-8")
                zout.writestr(item, data)
        shutil.move(tmp_path, docx_path)
    except Exception:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        raise


# ---------- font-family conventions ----------------------------------------
# Centralizing the CJK names here so changing a typography decision is one
# edit, not five. These are the "most-portable" names we picked — see the
# README's font-portability note for the reasoning.

# General-purpose CJK fonts (broadly available across Windows / macOS /
# Linux Chinese installs).
CJK_SONG    = "宋体"       # 衬线主力,正文首选
CJK_HEI     = "黑体"       # 无衬线,标题主力
CJK_FANGSONG = "仿宋"      # 公文正文(general fallback)
CJK_KAITI   = "楷体"       # 公文二级标题、注释
CJK_YAHEI   = "微软雅黑"   # 现代风格标题(技术文档常用)

# GB2312 / 方正 variants specifically required by formal 公文 规范.
# These names match what Windows 中文版 ships as (note the underscore).
# When the system doesn't have these exact names, Word falls back to
# nearest match — usually 仿宋 / 楷体 / 黑体 — which still reads as
# the right typographic genre.
CJK_FANGSONG_GB = "仿宋_GB2312"
CJK_KAITI_GB    = "楷体_GB2312"
CJK_FZXBS       = "方正小标宋简体"   # 公文文档大标题专用(二号)

LATIN_SERIF = "Times New Roman"
LATIN_SANS  = "Arial"
LATIN_MODERN = "Calibri"
LATIN_MONO  = "Consolas"
