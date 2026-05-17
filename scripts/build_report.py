"""Generate the work-report template (report_zh.docx).

Designed for the high-frequency, lower-ceremony stuff:
- 周报 / 月报 / 季度总结
- 项目进展汇报
- 部门工作总结

Not 公文 — no fixed line-height requirement, no 仿宋 mandate, no
red banner. Just clean, readable Chinese business document styling.

Key choices:
- A4 with standard Word margins (25.4 mm all around) — Microsoft's
  default. Mostly because that's what most users' Word is configured
  for, and reports often get pasted between docs.
- Title: 小一 (24pt) 黑体 center, with a Date/Author subtitle line.
- H1–H4: 黑体 阶梯 (16/14/12/11 pt). H5/H6 also configured so deep
  nesting doesn't break.
- Body: 小四 (12pt) 宋体, 1.5x line spacing, 2-character first-line
  indent. This is the standard "office Chinese" look.
- Tables: Word's built-in Table Grid + bolded header row (md2word
  already bolds header cells at render time, so no template work).
- Code blocks: small monospace with light grey shading (useful in
  tech-leaning reports without being obnoxious).
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Pt, RGBColor

from _template_helpers import (  # noqa: E402
    CJK_HEI, CJK_SONG, LATIN_MONO, LATIN_SANS, LATIN_SERIF,
    clear_first_line_indent, configure_style, enable_cjk_latin_autospace, set_default_eastasia_language_zh_cn,
    set_alignment, set_first_line_chars, set_left_border,
    set_line_spacing_multiple, set_shading, set_spacing_before_after,
)


def build_template(out_path: str) -> None:
    doc = Document()

    set_default_eastasia_language_zh_cn(doc)
    # ------------ page setup ----------
    section = doc.sections[0]
    section.page_height   = Cm(29.7)
    section.page_width    = Cm(21.0)
    section.top_margin    = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin   = Cm(3.18)
    section.right_margin  = Cm(3.18)

    # ------------ Normal (body) ----------
    # 小四 (12pt) 宋体, 1.5 倍行距, 首行缩进 2 字符
    normal = doc.styles["Normal"]
    configure_style(normal, latin=LATIN_SERIF, cjk=CJK_SONG, size_pt=12)
    set_line_spacing_multiple(normal, 1.5)
    set_first_line_chars(normal, 2)
    set_spacing_before_after(normal, 0, 0)

    # ------------ Title ----------
    # 小一 (24pt) 黑体, center, decent spacing below to separate from
    # the subtitle / date line.
    title = doc.styles["Title"]
    configure_style(title, latin=LATIN_SANS, cjk=CJK_HEI, size_pt=24, bold=True)
    set_line_spacing_multiple(title, 1.5)
    set_spacing_before_after(title, 0, 12)
    set_alignment(title, "center")
    clear_first_line_indent(title)

    # ------------ Subtitle (for date / author / period) ----------
    subtitle = doc.styles["Subtitle"]
    configure_style(subtitle, latin=LATIN_SERIF, cjk=CJK_SONG, size_pt=12)
    set_line_spacing_multiple(subtitle, 1.5)
    set_spacing_before_after(subtitle, 0, 18)
    set_alignment(subtitle, "center")
    clear_first_line_indent(subtitle)
    subtitle.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    # ------------ Headings 1..6 ----------
    # 黑体 阶梯: 三号 / 小三 / 四号 / 小四 / 五号 / 小五
    # 一律加粗,1.5 行距,无首行缩进
    heading_sizes = {1: 16, 2: 14, 3: 12, 4: 11, 5: 10.5, 6: 10}
    heading_spacing = {1: (18, 12), 2: (14, 8), 3: (10, 6), 4: (8, 4), 5: (6, 4), 6: (6, 4)}
    for level, size in heading_sizes.items():
        h = doc.styles[f"Heading {level}"]
        configure_style(h, latin=LATIN_SANS, cjk=CJK_HEI, size_pt=size, bold=True)
        set_line_spacing_multiple(h, 1.5)
        before, after = heading_spacing[level]
        set_spacing_before_after(h, before, after)
        set_alignment(h, "left")
        clear_first_line_indent(h)

    # ------------ List styles ----------
    for sty_name in ("List Bullet", "List Number"):
        sty = doc.styles[sty_name]
        configure_style(sty, latin=LATIN_SERIF, cjk=CJK_SONG, size_pt=12)
        set_line_spacing_multiple(sty, 1.5)
        clear_first_line_indent(sty)

    # ------------ Quote ----------
    # 报告里偶尔会有要强调的"领导原话"或者背景资料,加左边蓝色竖条更醒目
    quote = doc.styles["Quote"]
    configure_style(quote, latin=LATIN_SERIF, cjk=CJK_SONG, size_pt=12)
    quote.font.italic = False  # 中文斜体很丑
    set_line_spacing_multiple(quote, 1.5)
    set_spacing_before_after(quote, 6, 6)
    set_left_border(quote, "4A90E2", size=24)
    clear_first_line_indent(quote)
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    ppr = quote.element.find(qn("w:pPr"))
    if ppr is not None:
        ind = ppr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind"); ppr.append(ind)
        ind.set(qn("w:left"), "480")

    # ------------ Code Block ----------
    code = doc.styles.add_style("Code Block", WD_STYLE_TYPE.PARAGRAPH)
    code.base_style = doc.styles["Normal"]
    configure_style(code, latin=LATIN_MONO, cjk=CJK_SONG, size_pt=10)
    set_line_spacing_multiple(code, 1.15)
    set_spacing_before_after(code, 6, 6)
    set_shading(code, "F5F5F5")
    clear_first_line_indent(code)

    # ------------ Caption ----------
    caption = doc.styles["Caption"]
    configure_style(caption, latin=LATIN_SERIF, cjk=CJK_SONG, size_pt=10.5)
    caption.font.italic = False
    caption.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    set_line_spacing_multiple(caption, 1.15)
    set_spacing_before_after(caption, 2, 12)
    set_alignment(caption, "center")
    clear_first_line_indent(caption)

    enable_cjk_latin_autospace(doc)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    doc.save(out_path)
    print(f"✓ 模板已生成: {out_path}")


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    out = os.path.join(here, "..", "md2word", "templates", "report_zh.docx")
    build_template(os.path.abspath(out))
