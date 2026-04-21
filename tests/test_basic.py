"""Structural tests for md2word.

Each test is one feature. We assert on docx structure and XML, not on visual
appearance — visuals you verify by opening the file in Word.

Run with:
    python -m unittest discover tests

Each test:
  1. Reads tests/fixtures/<name>.md
  2. Pipes through parse() and Renderer into a temp .docx
  3. Reopens the .docx and asserts on paragraphs / runs / tables / XML
"""
from __future__ import annotations

import os
import tempfile
import unittest

from docx import Document
from docx.oxml.ns import qn

from md2word.parser import parse
from md2word.renderer import Renderer


FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")


# --------------------------- helpers --------------------------------------

def render_fixture(name: str):
    """Render tests/fixtures/<name>.md and return an opened Document."""
    md_path = os.path.join(FIXTURES, f"{name}.md")
    with open(md_path, encoding="utf-8") as f:
        md_text = f.read()
    ast = parse(md_text)
    fd, out = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    Renderer().render(ast, out)
    return Document(out), out


def style_of(paragraph) -> str:
    return paragraph.style.name if paragraph.style else ""


def run_fonts(run) -> dict:
    """Return {ascii, hAnsi, eastAsia, cs} for a run, or {} if not set."""
    rpr = run._element.find(qn("w:rPr"))
    if rpr is None:
        return {}
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        return {}
    return {
        slot: rfonts.get(qn(f"w:{slot}"))
        for slot in ("ascii", "hAnsi", "eastAsia", "cs")
    }


def hyperlinks_in(paragraph) -> list:
    """Return [(url, text)] for every w:hyperlink in the paragraph."""
    out = []
    for h in paragraph._p.findall(qn("w:hyperlink")):
        r_id = h.get(qn("r:id"))
        target = paragraph.part.rels[r_id].target_ref if r_id else ""
        text = "".join(t.text or "" for t in h.iter(qn("w:t")))
        out.append((target, text))
    return out


def table_as_grid(table) -> list:
    return [[cell.text for cell in row.cells] for row in table.rows]


# --------------------------- tests ----------------------------------------

class TestHeadings(unittest.TestCase):
    def test_levels_1_through_6(self):
        doc, _ = render_fixture("headings")
        levels = [style_of(p) for p in doc.paragraphs if p.text]
        expected = [f"Heading {i}" for i in range(1, 7)]
        self.assertEqual(levels, expected)

    def test_heading_text_preserved(self):
        doc, _ = render_fixture("headings")
        texts = [p.text for p in doc.paragraphs if p.text]
        self.assertEqual(texts, ["一级", "二级", "三级", "四级", "五级", "六级"])


class TestInlineFormatting(unittest.TestCase):
    def test_bold_italic_code_runs(self):
        doc, _ = render_fixture("inline")
        p = next(p for p in doc.paragraphs if p.text)
        # concatenated runs should reconstruct the paragraph
        self.assertIn("加粗", p.text)
        self.assertIn("斜体", p.text)
        self.assertIn("行内代码", p.text)

        # at least one bold run and one italic run must exist
        has_bold = any(r.bold for r in p.runs)
        has_italic = any(r.italic for r in p.runs)
        self.assertTrue(has_bold, "expected at least one bold run")
        self.assertTrue(has_italic, "expected at least one italic run")

    def test_inline_code_uses_code_font(self):
        doc, _ = render_fixture("inline")
        p = next(p for p in doc.paragraphs if p.text)
        code_run = next((r for r in p.runs if "行内代码" in r.text), None)
        self.assertIsNotNone(code_run)
        fonts = run_fonts(code_run)
        self.assertEqual(fonts.get("ascii"), "Consolas")

    def test_combined_bold_italic(self):
        doc, _ = render_fixture("inline")
        p = next(p for p in doc.paragraphs if p.text)
        combined = next((r for r in p.runs if "又粗又斜" in r.text), None)
        self.assertIsNotNone(combined)
        self.assertTrue(combined.bold)
        self.assertTrue(combined.italic)


class TestLists(unittest.TestCase):
    def test_bullet_list_style(self):
        doc, _ = render_fixture("lists")
        bullet_texts = [p.text for p in doc.paragraphs if style_of(p) == "List Bullet"]
        self.assertEqual(bullet_texts, ["苹果", "香蕉", "橘子"])

    def test_ordered_list_style(self):
        doc, _ = render_fixture("lists")
        number_texts = [p.text for p in doc.paragraphs if style_of(p) == "List Number"]
        self.assertEqual(number_texts, ["第一条", "第二条", "第三条"])


class TestNestedList(unittest.TestCase):
    def test_all_items_emitted(self):
        doc, _ = render_fixture("nested_list")
        items = [p.text for p in doc.paragraphs if p.text]
        expected = ["第一层甲", "第二层甲一", "第二层甲二",
                    "第一层乙", "第二层乙一", "第三层"]
        self.assertEqual(items, expected)

    def test_deeper_levels_are_indented(self):
        doc, _ = render_fixture("nested_list")
        by_text = {p.text: p for p in doc.paragraphs if p.text}
        level1 = by_text["第一层甲"].paragraph_format.left_indent
        level3 = by_text["第三层"].paragraph_format.left_indent
        # level1 is 0 or None (uses style's own indent); level3 got explicit extra indent
        l1 = level1.emu if level1 else 0
        l3 = level3.emu if level3 else 0
        self.assertGreater(l3, l1, "deeper list levels should be indented further")


class TestTable(unittest.TestCase):
    def test_dimensions_and_content(self):
        doc, _ = render_fixture("table")
        self.assertEqual(len(doc.tables), 1)
        grid = table_as_grid(doc.tables[0])
        self.assertEqual(grid, [
            ["参数", "值", "说明"],
            ["字号", "12pt", "小四"],
            ["行距", "1.5", "倍数"],
        ])

    def test_header_is_bold(self):
        doc, _ = render_fixture("table")
        header_row = doc.tables[0].rows[0]
        for cell in header_row.cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    if r.text.strip():
                        self.assertTrue(r.bold, f"header cell run {r.text!r} not bold")


class TestCodeBlock(unittest.TestCase):
    def test_uses_code_block_style(self):
        doc, _ = render_fixture("code")
        code_paras = [p for p in doc.paragraphs if style_of(p) == "Code Block"]
        self.assertEqual(len(code_paras), 1)

    def test_multiline_content_preserved(self):
        doc, _ = render_fixture("code")
        p = next(p for p in doc.paragraphs if style_of(p) == "Code Block")
        self.assertIn("def hello():", p.text)
        self.assertIn("你好", p.text)
        # should have line breaks inside (multiple w:br)
        breaks = p._p.findall(".//" + qn("w:br"))
        self.assertGreater(len(breaks), 0, "code block should preserve line breaks")

    def test_uses_monospace_font(self):
        doc, _ = render_fixture("code")
        p = next(p for p in doc.paragraphs if style_of(p) == "Code Block")
        r = p.runs[0]
        self.assertEqual(run_fonts(r).get("ascii"), "Consolas")


class TestQuote(unittest.TestCase):
    def test_uses_quote_style(self):
        doc, _ = render_fixture("quote")
        quote_paras = [p for p in doc.paragraphs if style_of(p) == "Quote"]
        self.assertGreater(len(quote_paras), 0)


class TestHyperlink(unittest.TestCase):
    def test_real_hyperlink_with_relationship(self):
        doc, _ = render_fixture("link")
        p = next(p for p in doc.paragraphs if p.text)
        links = hyperlinks_in(p)
        self.assertEqual(len(links), 1)
        url, text = links[0]
        self.assertEqual(url, "https://www.python.org/")
        self.assertEqual(text, "Python 官网")


class TestMissingImage(unittest.TestCase):
    def test_degrades_to_alt_text(self):
        doc, _ = render_fixture("missing_image")
        full_text = "\n".join(p.text for p in doc.paragraphs)
        self.assertIn("[图片: 架构图]", full_text)
        # shouldn't have crashed — rest of doc is there too
        self.assertIn("降级为文字占位符", full_text)


class TestCJKFontSlots(unittest.TestCase):
    def test_both_latin_and_eastasia_slots_set(self):
        doc, _ = render_fixture("cjk_mixed")
        p = next(p for p in doc.paragraphs if p.text)
        self.assertGreater(len(p.runs), 0)
        for r in p.runs:
            fonts = run_fonts(r)
            self.assertEqual(fonts.get("ascii"), "Times New Roman",
                             f"run {r.text!r} missing Latin font")
            self.assertEqual(fonts.get("eastAsia"), "宋体",
                             f"run {r.text!r} missing CJK font")


class TestThematicBreak(unittest.TestCase):
    def test_hr_rendered_as_bottom_border(self):
        doc, _ = render_fixture("hr")
        # should have three paragraphs: "上面的段落" / HR / "下面的段落"
        found_border = False
        for p in doc.paragraphs:
            pPr = p._p.find(qn("w:pPr"))
            if pPr is None:
                continue
            pBdr = pPr.find(qn("w:pBdr"))
            if pBdr is not None and pBdr.find(qn("w:bottom")) is not None:
                found_border = True
                break
        self.assertTrue(found_border, "expected at least one paragraph with bottom border")


class TestEndToEndSample(unittest.TestCase):
    """Smoke test: the comprehensive sample.md renders without errors."""

    def test_sample_renders(self):
        doc, out_path = render_fixture("sample")
        self.assertGreater(len(doc.paragraphs), 10)
        self.assertGreater(len(doc.tables), 0)
        os.unlink(out_path)


if __name__ == "__main__":
    unittest.main()
