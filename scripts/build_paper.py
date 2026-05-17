"""Generate the academic-paper template (paper_zh.docx).

A *generic* Chinese academic-paper layout — not tailored to any one
university or journal. Use cases:
- 期刊投稿初稿
- 学位论文章节(配合各校自己的最终模板)
- 学术研讨班 / 课程论文
- 项目结题报告

Conventions chosen to be the "most common" Chinese academic look,
based on the typesetting of mainstream education-ministry journals
(教育部认定的核心期刊):
- 边距: 左 30 mm (装订边), 右 25 mm, 上下 25 mm
- 标题: 小二 (18 pt) 宋体 加粗 居中
- 摘要标题 / 关键词: 小四 黑体, 加粗
- 一级标题: 三号 (16 pt) 黑体
- 二级: 四号 (14 pt) 黑体
- 三级: 小四 (12 pt) 黑体
- 正文: 小四 宋体, 1.5 倍行距, 首行缩进 2 字符
- 图注 / 表注: 五号 (10.5 pt) 宋体 居中
- 参考文献: 五号 宋体, 悬挂缩进

Each university has its own dissertation format — we don't try to
match any specific one. Users with a school-issued template should
use that template via `-t` instead. This file is for "I don't have
a template and need to write something that looks like a paper".
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Mm, Pt, RGBColor

from _template_helpers import (  # noqa: E402
    CJK_HEI, CJK_KAITI, CJK_SONG, LATIN_MONO, LATIN_SERIF,
    clear_first_line_indent, configure_style, enable_cjk_latin_autospace, set_default_eastasia_language_zh_cn,
    set_alignment, set_first_line_chars, set_line_spacing_multiple,
    set_shading, set_spacing_before_after,
)


def build_template(out_path: str) -> None:
    doc = Document()

    set_default_eastasia_language_zh_cn(doc)
    # ------------ page setup ----------
    # Left margin larger than right — standard for printed/bound academic
    # work where the left edge gets stapled or bound.
    section = doc.sections[0]
    section.page_height   = Cm(29.7)
    section.page_width    = Cm(21.0)
    section.top_margin    = Mm(25)
    section.bottom_margin = Mm(25)
    section.left_margin   = Mm(30)
    section.right_margin  = Mm(25)

    # ------------ Normal ----------
    # 小四 (12pt) 宋体, 1.5 倍行距, 首行缩进 2 字符
    normal = doc.styles["Normal"]
    configure_style(normal, latin=LATIN_SERIF, cjk=CJK_SONG, size_pt=12)
    set_line_spacing_multiple(normal, 1.5)
    set_first_line_chars(normal, 2)
    set_spacing_before_after(normal, 0, 0)
    set_alignment(normal, "both")  # justify, academic convention

    # ------------ Title (论文题目) ----------
    # 小二 (18pt) 宋体 加粗 居中
    title = doc.styles["Title"]
    configure_style(title, latin=LATIN_SERIF, cjk=CJK_SONG, size_pt=18, bold=True)
    set_line_spacing_multiple(title, 1.5)
    set_spacing_before_after(title, 0, 18)
    set_alignment(title, "center")
    clear_first_line_indent(title)

    # ------------ Subtitle (作者 / 单位 / 投稿信息) ----------
    subtitle = doc.styles["Subtitle"]
    configure_style(subtitle, latin=LATIN_SERIF, cjk=CJK_SONG, size_pt=12)
    set_line_spacing_multiple(subtitle, 1.5)
    set_spacing_before_after(subtitle, 0, 18)
    set_alignment(subtitle, "center")
    clear_first_line_indent(subtitle)

    # ------------ Headings ----------
    # 一级: 三号 黑体
    # 二级: 四号 黑体
    # 三级: 小四 黑体
    # 四级及以下:宋体加粗,字号递减 (避免过深的标题层级)
    heading_config = [
        (1, 16, CJK_HEI,  True,  18, 12),
        (2, 14, CJK_HEI,  True,  14, 8),
        (3, 12, CJK_HEI,  True,  10, 6),
        (4, 12, CJK_SONG, True,  8,  4),
        (5, 11, CJK_SONG, True,  6,  4),
        (6, 10.5, CJK_SONG, True, 6, 4),
    ]
    for level, size, cjk, bold, before, after in heading_config:
        h = doc.styles[f"Heading {level}"]
        configure_style(h, latin=LATIN_SERIF, cjk=cjk, size_pt=size, bold=bold)
        set_line_spacing_multiple(h, 1.5)
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
    # 学术引用:楷体,左缩进 2 字符 (与正文区分),无装饰边框
    quote = doc.styles["Quote"]
    configure_style(quote, latin=LATIN_SERIF, cjk=CJK_KAITI, size_pt=12)
    quote.font.italic = False
    set_line_spacing_multiple(quote, 1.5)
    set_spacing_before_after(quote, 6, 6)
    clear_first_line_indent(quote)
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    ppr = quote.element.find(qn("w:pPr"))
    if ppr is not None:
        ind = ppr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind"); ppr.append(ind)
        ind.set(qn("w:leftChars"), "200")

    # ------------ Code Block ----------
    # 学术论文里代码不常见,保持低调
    code = doc.styles.add_style("Code Block", WD_STYLE_TYPE.PARAGRAPH)
    code.base_style = doc.styles["Normal"]
    configure_style(code, latin=LATIN_MONO, cjk=CJK_SONG, size_pt=10)
    set_line_spacing_multiple(code, 1.15)
    set_spacing_before_after(code, 6, 6)
    set_shading(code, "F8F8F8")
    clear_first_line_indent(code)

    # ------------ Caption ----------
    # 五号 (10.5pt) 宋体 居中,小灰
    caption = doc.styles["Caption"]
    configure_style(caption, latin=LATIN_SERIF, cjk=CJK_SONG, size_pt=10.5)
    caption.font.italic = False
    caption.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
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
    out = os.path.join(here, "..", "md2word", "templates", "paper_zh.docx")
    build_template(os.path.abspath(out))
