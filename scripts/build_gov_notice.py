"""Generate the gov-notice template (gov_notice_zh.docx).

Implements the typographical core of the user's organization 公文 spec
(see /mnt/project/44441bbf-...-docx). Where that spec differs from
GB/T 9704-2012, the user's spec wins — they're the ones who'll
review the output:

| 元素        | 字体                | 字号  | 备注                  |
| ----------- | ------------------- | ----- | --------------------- |
| 标题        | 方正小标宋简体      | 二号  | 居中,不加粗,行距36 |
| 一级标题    | 黑体                | 三号  | "一、",不加粗       |
| 二级标题    | 楷体_GB2312         | 三号  | "(一)"              |
| 三级标题    | 仿宋_GB2312         | 三号  | "1." — 不加粗         |
| 四级标题    | 仿宋_GB2312         | 三号  | "(1)"                |
| 正文        | 仿宋_GB2312         | 三号  | 首行缩进2字符         |
| 边距        | 上33 下27 左27 右27 mm                                |
| 行距        | 固定 28 磅 (或 29)                                    |
| 段前段后    | 0 字符 / 0 行                                         |

Things we intentionally don't model in the .docx:
- 红头 (red banner with issuing agency name): every agency has its own
  wordmark; the user adds it in Word once and saves as their own
  template. We expose 发文机关 in frontmatter and render it as a
  centered paragraph at the top of the body for placeholder use.
- 印章 (seal): user adds an image in Word.
- 版记 / 骑缝章: page-layout decoration, out of scope.
- 方正小标宋简体 is a licensed font — we write the name in the
  template so users who have it (most Chinese government Windows
  installs do) get the right rendering, and those who don't will
  see Word substitute with the nearest available CJK heading font.
"""
from __future__ import annotations

import os
import sys

# Allow running this script directly (python scripts/build_gov_notice.py)
# without first installing the package. The helpers live next door.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Mm, Pt

from _template_helpers import (  # noqa: E402
    CJK_FANGSONG_GB, CJK_FZXBS, CJK_HEI, CJK_KAITI_GB,
    LATIN_MONO, LATIN_SERIF,
    clear_first_line_indent, configure_style, enable_cjk_latin_autospace,
    set_alignment, set_default_eastasia_language_zh_cn, set_first_line_chars,
    set_line_spacing_fixed_pt, set_line_spacing_multiple, set_shading,
    set_spacing_before_after, style_ppr,
)


# The user's spec says "固定 28 或 29 磅" — pick 28 (the more common one).
BODY_LINE_HEIGHT_PT  = 28
TITLE_LINE_HEIGHT_PT = 36


def build_template(out_path: str) -> None:
    doc = Document()

    # IMPORTANT: set the doc-default eastAsia language to zh-CN *first*,
    # before any styles touch w:rFonts. Without this, Word resolves CJK
    # characters via the document's default ja-JP fallback chain
    # (MS Gothic / MS Mincho), which is what users see as "为啥标题
    # 全是 MS Gothic". Setting zh-CN here flips that fallback chain to
    # 简体中文 fonts (SimSun / SimHei / etc.).
    set_default_eastasia_language_zh_cn(doc)

    # ------------ page setup (user's spec: 上33 下27 左27 右27) ----------
    section = doc.sections[0]
    section.page_height   = Cm(29.7)
    section.page_width    = Cm(21.0)
    section.top_margin    = Mm(33)
    section.bottom_margin = Mm(27)
    section.left_margin   = Mm(27)
    section.right_margin  = Mm(27)

    # ------------ Normal (正文: 仿宋GB2312, 三号, 首行缩进2字符, 行距28磅) ----------
    normal = doc.styles["Normal"]
    configure_style(normal, latin=LATIN_SERIF, cjk=CJK_FANGSONG_GB, size_pt=16)
    set_line_spacing_fixed_pt(normal, BODY_LINE_HEIGHT_PT)
    set_first_line_chars(normal, 2)
    set_spacing_before_after(normal, 0, 0)
    set_alignment(normal, "both")  # 两端对齐 (justify)

    # ------------ Heading 1 — 文档主标题 (方正小标宋简体 二号 居中 不加粗 行距36) ----------
    # markdown 的 `# 标题` 在公文场景里就是整篇文档的主标题。
    # 用户规范明示:方正小标宋简体, 二号 (22pt), 居中, 不加粗, 行间距 36。
    h1 = doc.styles["Heading 1"]
    configure_style(h1, latin=LATIN_SERIF, cjk=CJK_FZXBS, size_pt=22, bold=False)
    set_line_spacing_fixed_pt(h1, TITLE_LINE_HEIGHT_PT)
    set_spacing_before_after(h1, 0, 24)
    set_alignment(h1, "center")
    clear_first_line_indent(h1)

    # ------------ Heading 2 — "一、" 一级标题 (黑体 三号 不加粗) ----------
    h2 = doc.styles["Heading 2"]
    configure_style(h2, latin=LATIN_SERIF, cjk=CJK_HEI, size_pt=16, bold=False)
    set_line_spacing_fixed_pt(h2, BODY_LINE_HEIGHT_PT)
    set_spacing_before_after(h2, 0, 0)
    set_alignment(h2, "left")
    clear_first_line_indent(h2)
    set_first_line_chars(h2, 2)  # "一、" 保留两字符首行缩进位置

    # ------------ Heading 3 — "(一)" 二级标题 (楷体_GB2312 三号) ----------
    h3 = doc.styles["Heading 3"]
    configure_style(h3, latin=LATIN_SERIF, cjk=CJK_KAITI_GB, size_pt=16, bold=False)
    set_line_spacing_fixed_pt(h3, BODY_LINE_HEIGHT_PT)
    set_spacing_before_after(h3, 0, 0)
    set_alignment(h3, "left")
    clear_first_line_indent(h3)
    set_first_line_chars(h3, 2)

    # ------------ Heading 4 — "1." 三级标题 (仿宋_GB2312 三号 不加粗) ----------
    h4 = doc.styles["Heading 4"]
    configure_style(h4, latin=LATIN_SERIF, cjk=CJK_FANGSONG_GB, size_pt=16, bold=False)
    set_line_spacing_fixed_pt(h4, BODY_LINE_HEIGHT_PT)
    set_spacing_before_after(h4, 0, 0)
    set_alignment(h4, "left")
    clear_first_line_indent(h4)
    set_first_line_chars(h4, 2)

    # ------------ Heading 5 — "(1)" 四级标题 (仿宋_GB2312 三号 不加粗) ----------
    h5 = doc.styles["Heading 5"]
    configure_style(h5, latin=LATIN_SERIF, cjk=CJK_FANGSONG_GB, size_pt=16, bold=False)
    set_line_spacing_fixed_pt(h5, BODY_LINE_HEIGHT_PT)
    set_spacing_before_after(h5, 0, 0)
    set_alignment(h5, "left")
    clear_first_line_indent(h5)
    set_first_line_chars(h5, 2)

    # Heading 6 — beyond 公文 convention but kept so deep Markdown
    # doesn't degrade silently. Same look as Heading 5.
    h6 = doc.styles["Heading 6"]
    configure_style(h6, latin=LATIN_SERIF, cjk=CJK_FANGSONG_GB, size_pt=16, bold=False)
    set_line_spacing_fixed_pt(h6, BODY_LINE_HEIGHT_PT)
    set_alignment(h6, "left")
    clear_first_line_indent(h6)
    set_first_line_chars(h6, 2)

    # ------------ Title style — keep but make it cosmetically similar to H1 ----------
    # markdown emit goes through Heading 1, so Title is only used if a
    # user manually applies it. Keep it consistent so it doesn't look
    # jarring next to H1.
    title = doc.styles["Title"]
    configure_style(title, latin=LATIN_SERIF, cjk=CJK_FZXBS, size_pt=22, bold=False)
    set_line_spacing_fixed_pt(title, TITLE_LINE_HEIGHT_PT)
    set_spacing_before_after(title, 0, 24)
    set_alignment(title, "center")
    clear_first_line_indent(title)

    # ------------ Subtitle (副标题 / 文号 / 签发人: 仿宋, 三号, 居中) ----------
    subtitle = doc.styles["Subtitle"]
    configure_style(subtitle, latin=LATIN_SERIF, cjk=CJK_FANGSONG_GB, size_pt=16, bold=False)
    set_line_spacing_fixed_pt(subtitle, BODY_LINE_HEIGHT_PT)
    set_spacing_before_after(subtitle, 0, 6)
    set_alignment(subtitle, "center")
    clear_first_line_indent(subtitle)

    # ------------ List styles ----------
    for sty_name in ("List Bullet", "List Number"):
        sty = doc.styles[sty_name]
        configure_style(sty, latin=LATIN_SERIF, cjk=CJK_FANGSONG_GB, size_pt=16)
        set_line_spacing_fixed_pt(sty, BODY_LINE_HEIGHT_PT)
        clear_first_line_indent(sty)

    # ------------ Quote ----------
    quote = doc.styles["Quote"]
    configure_style(quote, latin=LATIN_SERIF, cjk=CJK_FANGSONG_GB, size_pt=16)
    set_line_spacing_fixed_pt(quote, BODY_LINE_HEIGHT_PT)
    set_spacing_before_after(quote, 6, 6)
    clear_first_line_indent(quote)
    # 2-character left indent for the quote block — restrained, no
    # decorative border (公文 doesn't do those)
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    ppr = style_ppr(quote)
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind"); ppr.append(ind)
    ind.set(qn("w:leftChars"), "200")

    # ------------ Code Block ----------
    code = doc.styles.add_style("Code Block", WD_STYLE_TYPE.PARAGRAPH)
    code.base_style = doc.styles["Normal"]
    configure_style(code, latin=LATIN_MONO, cjk=CJK_FANGSONG_GB, size_pt=10.5)
    set_line_spacing_multiple(code, 1.15)
    set_spacing_before_after(code, 6, 6)
    set_shading(code, "F5F5F5")
    clear_first_line_indent(code)

    # ------------ Caption (图注 / 表注) ----------
    # The user's spec doesn't address captions, but 公文 附件 typically
    # uses 仿宋 with a slightly smaller size for table/figure captions.
    caption = doc.styles["Caption"]
    configure_style(caption, latin=LATIN_SERIF, cjk=CJK_KAITI_GB, size_pt=14, bold=False)
    set_line_spacing_multiple(caption, 1.15)
    set_spacing_before_after(caption, 2, 12)
    set_alignment(caption, "center")
    clear_first_line_indent(caption)

    # ------------ Document-level settings ----------
    enable_cjk_latin_autospace(doc)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    doc.save(out_path)
    print(f"✓ 模板已生成: {out_path}")


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    out = os.path.join(here, "..", "md2word", "templates", "gov_notice_zh.docx")
    build_template(os.path.abspath(out))
