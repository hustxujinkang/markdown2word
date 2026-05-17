"""Command-line entry point.

Two subcommands:
- `md2word convert IN.md -o OUT.docx [-t TPL]`  (default if omitted)
- `md2word inspect TPL.docx [--json]`

Backwards compatibility: the old form `md2word IN.md -o OUT.docx` keeps
working — if the first positional arg isn't a known subcommand, we slot
`convert` in front of it before parsing.
"""
from __future__ import annotations

import argparse
import os
import sys

from .parser import parse
from .renderer import Renderer


_SUBCOMMANDS = {"convert", "inspect"}


def _build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        prog="md2word",
        description="将 Markdown 转换为 Word 文档,按 reference.docx 模板应用样式。",
    )
    sub = ap.add_subparsers(dest="command")

    # --- convert ---
    p_conv = sub.add_parser("convert", help="把 Markdown 转换为 Word")
    p_conv.add_argument("input", help="输入的 Markdown 文件")
    p_conv.add_argument("-o", "--output", help="输出 Word 路径(默认同名 .docx)")
    p_conv.add_argument(
        "-t", "--template",
        help="参考模板 reference.docx。不指定则使用内置中文模板。",
    )
    p_conv.add_argument(
        "--override-template-fonts",
        action="store_true",
        help="即使指定了自定义模板也强制覆盖字体(默认尊重模板)。",
    )

    # --- inspect ---
    p_insp = sub.add_parser("inspect", help="检查模板里 md2word 能识别的样式/字体")
    p_insp.add_argument("template", help="待检查的 .docx 模板")
    p_insp.add_argument("--json", action="store_true", help="以 JSON 格式输出")

    return ap


def main(argv=None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)

    # Backwards compatibility: `md2word IN.md ...` (no subcommand) →
    # treat as `md2word convert IN.md ...`. Detect by checking the first
    # non-flag argument: if it isn't one of our known subcommands and
    # isn't a help flag, prepend `convert`.
    if argv:
        first = argv[0]
        if first not in _SUBCOMMANDS and first not in ("-h", "--help"):
            argv = ["convert"] + argv

    ap = _build_parser()
    args = ap.parse_args(argv)

    if args.command == "inspect":
        return _run_inspect(args)
    if args.command == "convert":
        return _run_convert(args)

    ap.print_help()
    return 0


def _run_convert(args) -> int:
    if not os.path.exists(args.input):
        print(f"错误: 找不到文件 {args.input}", file=sys.stderr); return 2
    if args.template and not os.path.exists(args.template):
        print(f"错误: 找不到模板 {args.template}", file=sys.stderr); return 2

    out = args.output or (os.path.splitext(args.input)[0] + ".docx")
    with open(args.input, "r", encoding="utf-8") as f:
        md_text = f.read()
    doc = parse(md_text)

    respect = None
    if args.override_template_fonts:
        respect = False
    Renderer(template=args.template, respect_template_fonts=respect).render(doc, out)
    print(f"✓ 已生成 {out}")
    return 0


def _run_inspect(args) -> int:
    # Lazy import — inspect pulls in nothing convert doesn't, but keeping
    # the import here makes the dependency direction obvious: inspect is
    # an optional companion tool, not part of the conversion pipeline.
    from .inspect import inspect_template, format_report_text, format_report_json

    if not os.path.exists(args.template):
        print(f"错误: 找不到文件 {args.template}", file=sys.stderr); return 2
    try:
        report = inspect_template(args.template)
    except Exception as e:
        print(f"错误: 无法读取模板 ({e})", file=sys.stderr); return 1

    print(format_report_json(report) if args.json else format_report_text(report))
    return 0


if __name__ == "__main__":
    sys.exit(main())
