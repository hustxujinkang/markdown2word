"""Generate the tech-doc template (tech_doc_zh.docx).

For technical writing where the audience is engineers/PMs:
- 技术方案书
- 产品需求文档 (PRD)
- 系统设计文档
- 评审材料

Differences vs the work-report template:
- Narrower margins (20 mm) — tech docs often have wide diagrams,
  code samples, and tables. Margins eat horizontal real estate.
- 微软雅黑 for headings instead of 黑体 — more modern look, common
  in tech-industry templates.
- Deeper heading hierarchy emphasis (4 levels styled differently,
  not just font-size shrink).
- Code blocks more prominent — bigger, more contrast — because
  this is a doc style where readers actually read code.
- Captions in 楷体 with a deliberate "image label" look.
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Mm, Pt, RGBColor

from _template_helpers import (  # noqa: E402
    CJK_HEI, CJK_KAITI, CJK_SONG, CJK_YAHEI, LATIN_MODERN, LATIN_MONO,
    LATIN_SANS, LATIN_SERIF, clear_first_line_indent, configure_style,
    enable_cjk_latin_autospace, set_default_eastasia_language_zh_cn, set_alignment, set_first_line_chars,
    set_left_border, set_line_spacing_multiple, set_shading,
    set_spacing_before_after,
)


def build_template(out_path: str) -> None:
    doc = Document()

    set_default_eastasia_language_zh_cn(doc)
    # ------------ page setup ----------
    # 20 mm margins — tighter than office default to make room for
    # wide diagrams and tables.
    section = doc.sections[0]
    section.page_height   = Cm(29.7)
    section.page_width    = Cm(21.0)
    section.top_margin    = Mm(20)
    section.bottom_margin = Mm(20)
    section.left_margin   = Mm(20)
    section.right_margin  = Mm(20)

    # ------------ Normal ----------
    # 小四 (12pt) 宋体, 1.5 倍行距, 首行缩进 2 字符
    normal = doc.styles["Normal"]
    configure_style(normal, latin=LATIN_MODERN, cjk=CJK_SONG, size_pt=12)
    set_line_spacing_multiple(normal, 1.5)
    set_first_line_chars(normal, 2)
    set_spacing_before_after(normal, 0, 0)

    # ------------ Title ----------
    # 小二 (18pt) 微软雅黑, 居中
    title = doc.styles["Title"]
    configure_style(title, latin=LATIN_SANS, cjk=CJK_YAHEI, size_pt=22, bold=True)
    set_line_spacing_multiple(title, 1.5)
    set_spacing_before_after(title, 0, 12)
    set_alignment(title, "center")
    clear_first_line_indent(title)

    # ------------ Subtitle ----------
    subtitle = doc.styles["Subtitle"]
    configure_style(subtitle, latin=LATIN_MODERN, cjk=CJK_SONG, size_pt=12)
    set_line_spacing_multiple(subtitle, 1.5)
    set_spacing_before_after(subtitle, 0, 18)
    set_alignment(subtitle, "center")
    clear_first_line_indent(subtitle)
    subtitle.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    # ------------ Headings 1..6 — modern hierarchy ----------
    # H1/H2: 微软雅黑 加粗,大字号 — 用作"模块/章节"分界
    # H3/H4: 微软雅黑 加粗,中字号 — 用作功能/子模块
    # H5/H6: 黑体 加粗,小字号 — 细分要点
    heading_config = [
        # level, size_pt, cjk, latin,           bold,  before, after
        (1, 18, CJK_YAHEI, LATIN_SANS,    True,  18, 10),
        (2, 15, CJK_YAHEI, LATIN_SANS,    True,  14, 8),
        (3, 13, CJK_YAHEI, LATIN_SANS,    True,  10, 6),
        (4, 12, CJK_YAHEI, LATIN_SANS,    True,  8,  4),
        (5, 11, CJK_HEI,   LATIN_SANS,    True,  6,  4),
        (6, 11, CJK_HEI,   LATIN_SANS,    False, 6,  4),
    ]
    for level, size, cjk, latin, bold, before, after in heading_config:
        h = doc.styles[f"Heading {level}"]
        configure_style(h, latin=latin, cjk=cjk, size_pt=size, bold=bold)
        set_line_spacing_multiple(h, 1.5)
        set_spacing_before_after(h, before, after)
        set_alignment(h, "left")
        clear_first_line_indent(h)

    # ------------ List styles ----------
    for sty_name in ("List Bullet", "List Number"):
        sty = doc.styles[sty_name]
        configure_style(sty, latin=LATIN_MODERN, cjk=CJK_SONG, size_pt=12)
        set_line_spacing_multiple(sty, 1.5)
        clear_first_line_indent(sty)

    # ------------ Quote ----------
    # 蓝色边条,稍微加粗的字号 - 技术文档常用来引用 RFC/规范原文
    quote = doc.styles["Quote"]
    configure_style(quote, latin=LATIN_MODERN, cjk=CJK_SONG, size_pt=12)
    quote.font.italic = False
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
    # 技术文档里代码块要醒目:更大的字号 (五号 10.5pt) + 更明显的底色
    code = doc.styles.add_style("Code Block", WD_STYLE_TYPE.PARAGRAPH)
    code.base_style = doc.styles["Normal"]
    configure_style(code, latin=LATIN_MONO, cjk=CJK_SONG, size_pt=10.5)
    set_line_spacing_multiple(code, 1.2)
    set_spacing_before_after(code, 6, 6)
    set_shading(code, "F0F0F2")  # 略带蓝灰
    clear_first_line_indent(code)

    # ------------ Caption ----------
    # 楷体 — 区别于正文宋体,视觉上立即识别为"图/表说明"
    caption = doc.styles["Caption"]
    configure_style(caption, latin=LATIN_SERIF, cjk=CJK_KAITI, size_pt=10.5)
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
    out = os.path.join(here, "..", "md2word", "templates", "tech_doc_zh.docx")
    build_template(os.path.abspath(out))
