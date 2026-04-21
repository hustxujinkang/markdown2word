"""AST → docx.

Walks the Document AST and emits into a python-docx `Document`. If you pass
a reference.docx, we load it and clear its body so only your template's
styles, theme, sections, headers/footers remain. If not, python-docx's
built-in default template is used (it already ships with Heading 1..9,
Normal, List Bullet, List Number, Quote, Intense Quote).
"""
from __future__ import annotations

import os
import urllib.request
import tempfile
from dataclasses import dataclass
from typing import Optional

from docx import Document as DocxDocument
from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu, Inches, Pt

from . import ast as A
from .fonts import DEFAULTS, FontSpec, apply_font


# ------------------------- configuration ----------------------------------

@dataclass
class FontProfile:
    body: FontSpec = None
    heading: FontSpec = None
    code: FontSpec = None

    def __post_init__(self):
        self.body = self.body or DEFAULTS["body"]
        self.heading = self.heading or DEFAULTS["heading"]
        self.code = self.code or DEFAULTS["code"]


# ---------------- style name resolution with fallback --------------------

def _style_or_fallback(doc, preferred: list) -> Optional[str]:
    """Return the first style name from `preferred` that actually exists."""
    existing = {s.name for s in doc.styles}
    for name in preferred:
        if name in existing:
            return name
    return None


# ---------------------------- main renderer -------------------------------

_BUILTIN_TEMPLATE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "templates", "default_zh.docx"
)


class Renderer:
    def __init__(self, template: Optional[str] = None, fonts: Optional[FontProfile] = None):
        resolved = template or (_BUILTIN_TEMPLATE if os.path.exists(_BUILTIN_TEMPLATE) else None)
        self.doc = DocxDocument(resolved) if resolved else DocxDocument()
        self.fonts = fonts or FontProfile()
        if resolved:
            self._clear_body()
        self._cache_styles()

    # purge the loaded template's body so we start clean
    def _clear_body(self) -> None:
        body = self.doc.element.body
        sectPr = body.find(qn("w:sectPr"))
        for child in list(body):
            if child is not sectPr:
                body.remove(child)

    def _cache_styles(self) -> None:
        d = self.doc
        self.sty_heading = [
            _style_or_fallback(d, [f"Heading {i}", f"标题 {i}"]) for i in range(1, 7)
        ]
        self.sty_normal  = _style_or_fallback(d, ["Normal", "正文"])
        self.sty_bullet  = _style_or_fallback(d, ["List Bullet", "无序列表"])
        self.sty_number  = _style_or_fallback(d, ["List Number", "有序列表"])
        self.sty_quote   = _style_or_fallback(d, ["Quote", "Intense Quote", "引用"])
        self.sty_code    = _style_or_fallback(d, ["Code Block", "Code", "代码块", "HTML Preformatted"])
        self.sty_caption = _style_or_fallback(d, ["Caption", "图注"])

    # --------- entry point ---------

    def render(self, document: A.Document, output_path: str) -> None:
        for block in document.blocks:
            self._render_block(block)
        self.doc.save(output_path)

    # --------- block dispatch ---------

    def _render_block(self, block, parent=None) -> None:
        if isinstance(block, A.Heading):
            self._render_heading(block)
        elif isinstance(block, A.Paragraph):
            self._render_paragraph(block)
        elif isinstance(block, A.CodeBlock):
            self._render_code(block)
        elif isinstance(block, A.Quote):
            for child in block.children:
                self._render_block(child)
                # tag the just-added paragraph with the Quote style
                if self.sty_quote and self.doc.paragraphs:
                    self.doc.paragraphs[-1].style = self.doc.styles[self.sty_quote]
        elif isinstance(block, A.BulletList):
            self._render_list(block, numbered=False, level=0)
        elif isinstance(block, A.OrderedList):
            self._render_list(block, numbered=True, level=0)
        elif isinstance(block, A.Table):
            self._render_table(block)
        elif isinstance(block, A.Image):
            self._render_block_image(block)
        elif isinstance(block, A.ThematicBreak):
            self._render_hr()

    # --------- individual blocks ---------

    def _render_heading(self, h: A.Heading) -> None:
        style_name = self.sty_heading[h.level - 1] or self.sty_normal
        p = self.doc.add_paragraph(style=self.doc.styles[style_name]) if style_name else self.doc.add_paragraph()
        self._write_inlines(p, h.inlines, font=self.fonts.heading)

    def _render_paragraph(self, p_node: A.Paragraph) -> None:
        p = self.doc.add_paragraph(style=self.doc.styles[self.sty_normal]) if self.sty_normal else self.doc.add_paragraph()
        self._write_inlines(p, p_node.inlines, font=self.fonts.body)

    def _render_code(self, cb: A.CodeBlock) -> None:
        style = self.doc.styles[self.sty_code] if self.sty_code else (self.doc.styles[self.sty_normal] if self.sty_normal else None)
        p = self.doc.add_paragraph(style=style) if style else self.doc.add_paragraph()
        # preserve line breaks inside a single paragraph
        lines = cb.code.split("\n")
        for i, line in enumerate(lines):
            run = p.add_run(line)
            apply_font(run, self.fonts.code)
            if i < len(lines) - 1:
                run.add_break(WD_BREAK.LINE)

    def _render_list(self, node, numbered: bool, level: int) -> None:
        style_name = self.sty_number if numbered else self.sty_bullet
        for item in node.items:
            # each item's first block becomes a list paragraph; subsequent
            # blocks (continuation paragraphs) attach at the same level
            first = True
            for child in item.children:
                if isinstance(child, (A.BulletList, A.OrderedList)):
                    self._render_list(child, isinstance(child, A.OrderedList), level + 1)
                else:
                    if isinstance(child, A.Paragraph) and first:
                        p = self.doc.add_paragraph(
                            style=self.doc.styles[style_name]) if style_name else self.doc.add_paragraph()
                        self._write_inlines(p, child.inlines, font=self.fonts.body)
                        if level > 0:
                            p.paragraph_format.left_indent = Inches(0.25 * (level + 1))
                        first = False
                    else:
                        self._render_block(child)

    def _render_table(self, t: A.Table) -> None:
        header_cells = t.header.cells if t.header else None
        ncols = len(header_cells) if header_cells else (len(t.rows[0].cells) if t.rows else 0)
        if ncols == 0:
            return
        total_rows = (1 if header_cells else 0) + len(t.rows)
        table = self.doc.add_table(rows=total_rows, cols=ncols)
        try:
            table.style = "Table Grid"
        except KeyError:
            pass

        r_idx = 0
        if header_cells:
            for j, inlines in enumerate(header_cells):
                cell = table.cell(0, j)
                cell.text = ""
                p = cell.paragraphs[0]
                self._write_inlines(p, inlines, font=self.fonts.body, force_bold=True)
            r_idx = 1
        for row in t.rows:
            for j, inlines in enumerate(row.cells):
                if j >= ncols:
                    break
                cell = table.cell(r_idx, j)
                cell.text = ""
                p = cell.paragraphs[0]
                self._write_inlines(p, inlines, font=self.fonts.body)
            r_idx += 1

    def _render_block_image(self, img: A.Image) -> None:
        p = self.doc.add_paragraph()
        run = p.add_run()
        path = _resolve_image(img.src)
        if path:
            try:
                run.add_picture(path, width=self._content_width())
            except Exception:
                run.add_text(f"[图片: {img.alt or img.src}]")
        else:
            run.add_text(f"[图片: {img.alt or img.src}]")
        if img.alt and self.sty_caption:
            cap = self.doc.add_paragraph(img.alt, style=self.doc.styles[self.sty_caption])
            for r in cap.runs:
                apply_font(r, self.fonts.body)

    def _render_hr(self) -> None:
        p = self.doc.add_paragraph()
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "auto")
        pBdr.append(bottom)
        pPr.append(pBdr)

    # --------- inlines ---------

    def _write_inlines(self, p, inlines, font: FontSpec, force_bold: bool = False) -> None:
        for node in inlines:
            if isinstance(node, A.Text):
                if node.value == "\n":
                    p.add_run().add_break(WD_BREAK.LINE); continue
                run = p.add_run(node.value)
                run.bold = node.bold or force_bold
                run.italic = node.italic
                apply_font(run, self.fonts.code if node.code else font)
            elif isinstance(node, A.Link):
                self._add_hyperlink(p, node.url, node.text, font)
            elif isinstance(node, A.InlineImage):
                run = p.add_run()
                path = _resolve_image(node.src)
                if path:
                    try:
                        run.add_picture(path, height=Pt(14))
                    except Exception:
                        run.add_text(f"[{node.alt or 'image'}]")
                else:
                    run.add_text(f"[{node.alt or 'image'}]")

    def _add_hyperlink(self, paragraph, url: str, text: str, font: FontSpec) -> None:
        part = paragraph.part
        r_id = part.relate_to(
            url,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            is_external=True,
        )
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("r:id"), r_id)

        run = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        rStyle = OxmlElement("w:rStyle"); rStyle.set(qn("w:val"), "Hyperlink"); rPr.append(rStyle)
        run.append(rPr)
        t = OxmlElement("w:t"); t.text = text; t.set(qn("xml:space"), "preserve"); run.append(t)
        hyperlink.append(run)
        paragraph._p.append(hyperlink)

        # also set CJK font on the hyperlink's run
        rfonts = OxmlElement("w:rFonts")
        rfonts.set(qn("w:ascii"), font.latin); rfonts.set(qn("w:hAnsi"), font.latin)
        rfonts.set(qn("w:eastAsia"), font.cjk); rfonts.set(qn("w:cs"), font.latin)
        rPr.insert(0, rfonts)

    # --------- helpers ---------

    def _content_width(self) -> Emu:
        section = self.doc.sections[0]
        return Emu(section.page_width - section.left_margin - section.right_margin)


# ---------------------- image resolution helpers --------------------------

def _resolve_image(src: str) -> Optional[str]:
    """Return a local path that python-docx can embed, or None."""
    if src.startswith(("http://", "https://")):
        try:
            suffix = os.path.splitext(src)[1] or ".png"
            fd, tmp = tempfile.mkstemp(suffix=suffix)
            os.close(fd)
            urllib.request.urlretrieve(src, tmp)
            return tmp
        except Exception:
            return None
    return src if os.path.exists(src) else None
