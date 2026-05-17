"""Template inspector.

`md2word inspect template.docx` opens a `.docx` template and reports
whether the styles md2word looks for actually exist, whether the styles
declare CJK fonts (the silent failure mode where Chinese ends up in
Calibri), and a few page-setup basics. The goal is to remove the "why
isn't my template taking effect?" support burden — the answer is usually
"the style isn't named what we expect, or its eastAsia slot is empty",
and now the tool can say so out loud.

This is read-only; we never modify the template.
"""
from __future__ import annotations

import json
import os
import sys
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from docx import Document as DocxDocument
from docx.oxml.ns import qn


# What md2word looks for, in priority order. Mirror of the lookup tables
# in renderer.Renderer._cache_styles — if you change one, change the other.
STYLE_LOOKUPS = {
    "一级标题":     ["Heading 1", "标题 1"],
    "二级标题":     ["Heading 2", "标题 2"],
    "三级标题":     ["Heading 3", "标题 3"],
    "四级标题":     ["Heading 4", "标题 4"],
    "五级标题":     ["Heading 5", "标题 5"],
    "六级标题":     ["Heading 6", "标题 6"],
    "正文":         ["Normal", "正文"],
    "无序列表":     ["List Bullet", "无序列表"],
    "有序列表":     ["List Number", "有序列表"],
    "引用块":       ["Quote", "Intense Quote", "引用"],
    "代码块":       ["Code Block", "Code", "代码块", "HTML Preformatted"],
    "图注":         ["Caption", "图注"],
}


@dataclass
class StyleReport:
    role: str                         # human label, e.g. "一级标题"
    expected: List[str]               # names we tried
    matched: Optional[str]            # the name that actually exists, or None
    matched_rank: int = -1            # 0 = first preference, 1 = second, ...
    fonts: Dict[str, Optional[str]] = field(default_factory=dict)
    cjk_set: bool = False             # eastAsia slot present
    latin_set: bool = False           # ascii slot present
    size_pt: Optional[float] = None


@dataclass
class PageReport:
    width_mm: float
    height_mm: float
    margin_top_mm: float
    margin_bottom_mm: float
    margin_left_mm: float
    margin_right_mm: float


@dataclass
class TemplateReport:
    path: str
    style_count: int
    styles: List[StyleReport]
    page: Optional[PageReport]
    warnings: List[str] = field(default_factory=list)


# --------------------------- public API ----------------------------------

def inspect_template(path: str) -> TemplateReport:
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    doc = DocxDocument(path)
    style_count = sum(1 for _ in doc.styles)
    style_reports = [_inspect_style(doc, role, expected) for role, expected in STYLE_LOOKUPS.items()]
    page = _inspect_page(doc)
    warnings = _collect_warnings(style_reports)
    return TemplateReport(
        path=path,
        style_count=style_count,
        styles=style_reports,
        page=page,
        warnings=warnings,
    )


def format_report_text(report: TemplateReport) -> str:
    """Render the report as a human-friendly text block."""
    lines = []
    lines.append(f"模板文件: {report.path}")
    if report.page:
        p = report.page
        lines.append(
            f"页面:    {p.width_mm:.1f} × {p.height_mm:.1f} mm  "
            f"边距 上{p.margin_top_mm:.1f} / 下{p.margin_bottom_mm:.1f} / "
            f"左{p.margin_left_mm:.1f} / 右{p.margin_right_mm:.1f} mm"
        )
    lines.append(f"样式总数: {report.style_count}")
    lines.append("")
    lines.append("样式匹配")
    lines.append("─" * 60)
    for s in report.styles:
        mark = "✓" if s.matched else "✗"
        if s.matched:
            font_desc = _describe_fonts(s)
            rank_note = "" if s.matched_rank == 0 else f"  (备选 #{s.matched_rank + 1})"
            lines.append(f"  {mark} {s.role:<8} → {s.matched}{rank_note}")
            lines.append(f"      字体: {font_desc}")
        else:
            expected = " / ".join(s.expected)
            lines.append(f"  {mark} {s.role:<8} 未找到 ({expected})")
            lines.append(f"      → 将降级为 Normal/正文 样式")
    lines.append("")
    if report.warnings:
        lines.append(f"发现 {len(report.warnings)} 个潜在问题:")
        for w in report.warnings:
            lines.append(f"  • {w}")
    else:
        lines.append("✓ 未发现潜在问题")
    return "\n".join(lines)


def format_report_json(report: TemplateReport) -> str:
    def to_dict(obj):
        if hasattr(obj, "__dict__"):
            return {k: to_dict(v) for k, v in obj.__dict__.items()}
        if isinstance(obj, list):
            return [to_dict(x) for x in obj]
        if isinstance(obj, dict):
            return {k: to_dict(v) for k, v in obj.items()}
        return obj
    return json.dumps(to_dict(report), ensure_ascii=False, indent=2)


# --------------------------- internals -----------------------------------

def _inspect_style(doc, role: str, expected: List[str]) -> StyleReport:
    existing = {s.name: s for s in doc.styles}
    matched = None
    matched_rank = -1
    for i, name in enumerate(expected):
        if name in existing:
            matched = name
            matched_rank = i
            break
    if matched is None:
        return StyleReport(role=role, expected=expected, matched=None)

    style = existing[matched]
    fonts, size_pt = _style_fonts(style)
    return StyleReport(
        role=role,
        expected=expected,
        matched=matched,
        matched_rank=matched_rank,
        fonts=fonts,
        cjk_set=bool(fonts.get("eastAsia")),
        latin_set=bool(fonts.get("ascii")),
        size_pt=size_pt,
    )


def _style_fonts(style) -> tuple:
    """Read the four font slots and size from a Style's rPr, if present."""
    elem = style.element if hasattr(style, "element") else style._element
    rpr = elem.find(qn("w:rPr"))
    fonts: Dict[str, Optional[str]] = {"ascii": None, "hAnsi": None, "eastAsia": None, "cs": None}
    size_pt: Optional[float] = None
    if rpr is None:
        return fonts, size_pt
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is not None:
        for slot in fonts:
            fonts[slot] = rfonts.get(qn(f"w:{slot}"))
    sz = rpr.find(qn("w:sz"))
    if sz is not None:
        val = sz.get(qn("w:val"))
        if val:
            try:
                # w:sz is in half-points
                size_pt = float(val) / 2.0
            except ValueError:
                pass
    return fonts, size_pt


def _describe_fonts(s: StyleReport) -> str:
    latin = s.fonts.get("ascii")
    cjk = s.fonts.get("eastAsia")
    size = f", {s.size_pt:g}pt" if s.size_pt is not None else ""
    if not latin and not cjk:
        return f"未在样式中显式设置(继承默认){size}"
    parts = []
    parts.append(f"西文 {latin or '未设置'}")
    if cjk:
        parts.append(f"中文 {cjk}")
    else:
        parts.append("中文 ⚠ 未设置")
    return ", ".join(parts) + size


def _inspect_page(doc) -> Optional[PageReport]:
    if not doc.sections:
        return None
    sec = doc.sections[0]
    # python-docx returns these as Emu objects (or None). Convert to mm.
    def mm(x) -> float:
        if x is None:
            return 0.0
        # An Emu is an int subclass; .emu attribute available, but plain
        # int division also works since the class defines no __int__
        # surprises.
        return float(x) / 36000.0
    return PageReport(
        width_mm=mm(sec.page_width),
        height_mm=mm(sec.page_height),
        margin_top_mm=mm(sec.top_margin),
        margin_bottom_mm=mm(sec.bottom_margin),
        margin_left_mm=mm(sec.left_margin),
        margin_right_mm=mm(sec.right_margin),
    )


def _collect_warnings(reports: List[StyleReport]) -> List[str]:
    warnings = []
    for s in reports:
        if s.matched is None:
            warnings.append(
                f"{s.role} 样式缺失({' / '.join(s.expected)}),"
                "对应 Markdown 元素将降级到 Normal"
            )
            continue
        # Only flag missing CJK font for styles where Chinese text is
        # actually expected. Code Block / Caption can reasonably skip it.
        if s.role in {"正文", "一级标题", "二级标题", "三级标题",
                      "四级标题", "五级标题", "六级标题", "引用块"}:
            if not s.cjk_set:
                warnings.append(
                    f"{s.role}({s.matched})未设置中文字体(eastAsia),"
                    "中文可能回退到 Calibri,建议在样式里显式设置"
                )
        if s.matched_rank > 0:
            warnings.append(
                f"{s.role} 匹配到的是备选名 {s.matched}(首选 {s.expected[0]} 不存在),"
                "建议把模板里的样式重命名为首选名以保证兼容性"
            )
    return warnings


# --------------------------- CLI entry -----------------------------------

def main(argv=None) -> int:
    import argparse
    ap = argparse.ArgumentParser(
        prog="md2word inspect",
        description="检查一份 .docx 模板,看 md2word 能识别到哪些样式、缺哪些字体设置。",
    )
    ap.add_argument("template", help="待检查的 .docx 模板")
    ap.add_argument("--json", action="store_true", help="以 JSON 格式输出")
    args = ap.parse_args(argv)

    if not os.path.exists(args.template):
        print(f"错误: 找不到文件 {args.template}", file=sys.stderr); return 2
    try:
        report = inspect_template(args.template)
    except Exception as e:
        print(f"错误: 无法读取模板 ({e})", file=sys.stderr); return 1

    if args.json:
        print(format_report_json(report))
    else:
        print(format_report_text(report))
    return 0


if __name__ == "__main__":
    sys.exit(main())
