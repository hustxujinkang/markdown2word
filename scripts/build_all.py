"""Build every bundled template in one go.

Run this whenever you change a template build script or
_template_helpers.py — regenerates all the .docx assets that ship
with the package.

Two-pass build:
1. Run each build_*.py to generate the bare styled template (only
   paragraph styles, page setup, fonts — empty body).
2. Use md2word itself to render the matching samples/*.md back into
   the template file, so opening the .docx in Word shows a worked
   sample instead of a blank page.

The double-duty design works because Renderer.__init__ calls _clear_body
when loading a template — any body content is discarded and only the
styles survive. So a template that *also* contains sample content is
still a valid md2word template.
"""
from __future__ import annotations

import os
import shutil
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

sys.path.insert(0, HERE)
sys.path.insert(0, ROOT)

import build_default_template
import build_gov_notice
import build_paper
import build_report
import build_tech_doc

# Helper for the MS-Gothic / theme-CJK-empty fix
from _template_helpers import patch_theme_cjk_fonts


def _overlay_sample(template_path: str, example_md: str) -> bool:
    """Render `example_md` against `template_path`, write result back
    to the same path. Returns True if overlay happened, False if the
    example file doesn't exist."""
    if not os.path.exists(example_md):
        return False
    # Lazy import — md2word depends on being installable, importing
    # at module top would slow the pass-1 build and is unnecessary
    # for users who just want plain styled templates.
    from md2word.parser import parse
    from md2word.renderer import Renderer

    with open(example_md, "r", encoding="utf-8") as f:
        md_text = f.read()
    ast = parse(md_text)
    fd, tmp = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    try:
        Renderer(template=template_path).render(ast, tmp)
        shutil.move(tmp, template_path)
        return True
    except Exception:
        # Don't leave a stray temp file behind on failure
        if os.path.exists(tmp):
            os.unlink(tmp)
        raise


# Per-template theme CJK font pair. Mirrors the heading/body fonts
# each template's build script picked, so theme-referenced styles
# fall back through the right family rather than to MS Gothic.
_THEME_CJK = {
    "default_zh.docx":    ("黑体",   "宋体"),
    "gov_notice_zh.docx": ("方正小标宋简体", "仿宋_GB2312"),
    "report_zh.docx":     ("黑体",   "宋体"),
    "tech_doc_zh.docx":   ("微软雅黑", "宋体"),
    "paper_zh.docx":      ("黑体",   "宋体"),
}


def main():
    out_dir = os.path.abspath(os.path.join(ROOT, "md2word", "templates"))
    samples_dir = os.path.abspath(os.path.join(ROOT, "samples"))
    os.makedirs(out_dir, exist_ok=True)

    # Each entry: (template filename, build function, sample markdown
    # to overlay or None to leave bare). default stays bare so users
    # who want a clean canvas have one.
    builders = [
        ("default_zh.docx",    build_default_template.build_template, None),
        ("gov_notice_zh.docx", build_gov_notice.build_template,       "gov_notice_demo.md"),
        ("report_zh.docx",     build_report.build_template,           "report_demo.md"),
        ("tech_doc_zh.docx",   build_tech_doc.build_template,         "tech_doc_demo.md"),
        ("paper_zh.docx",      build_paper.build_template,            "paper_demo.md"),
    ]

    # Pass 1 — bare styled templates
    for name, fn, _ in builders:
        fn(os.path.join(out_dir, name))

    # Pass 2 — overlay sample content
    print()
    print("嵌入示例内容,让模板打开即是样品:")
    for name, _, sample in builders:
        if sample is None:
            print(f"  · {name}: 保持空白(供需要干净底版的用户使用)")
            continue
        tpl = os.path.join(out_dir, name)
        md  = os.path.join(samples_dir, sample)
        if _overlay_sample(tpl, md):
            print(f"  ↳ {name}: 已嵌入 {sample}")
        else:
            print(f"  · {name}: 未找到 {sample},保持空白")

    # Pass 3 — patch theme1.xml so theme-referenced styles don't fall
    # back to MS Gothic. This is post-save because python-docx doesn't
    # expose the theme part directly.
    print()
    print("修补 theme1.xml 中文字体槽位:")
    for name, _, _ in builders:
        tpl = os.path.join(out_dir, name)
        major, minor = _THEME_CJK.get(name, ("黑体", "宋体"))
        patch_theme_cjk_fonts(tpl, major_ea=major, minor_ea=minor)
        print(f"  ↳ {name}: 主题字体 → 标题 {major} / 正文 {minor}")


if __name__ == "__main__":
    main()
