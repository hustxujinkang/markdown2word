"""Command-line entry point.

Subcommands:
- `md2word convert IN.md -o OUT.docx [-t TPL|@NAME]`  (default if omitted)
- `md2word inspect TPL.docx [--json]`
- `md2word templates [list|show NAME|eject NAME -o PATH]`

Backwards compatibility: the old form `md2word IN.md -o OUT.docx` keeps
working — if the first positional arg isn't a known subcommand, we slot
`convert` in front of it before parsing.
"""
from __future__ import annotations

import argparse
import os
import shutil
import sys

from .parser import parse
from .renderer import Renderer


_SUBCOMMANDS = {"convert", "inspect", "templates"}


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
        help="模板。可以是 .docx 文件路径,或 @NAME 形式引用内置模板"
             "(如 @gov_notice / @report / @tech_doc / @paper)。"
             "不指定则用 @default。",
    )
    p_conv.add_argument(
        "--override-template-fonts",
        action="store_true",
        help="即使指定了自定义模板也强制覆盖字体(默认尊重模板)。",
    )

    # --- inspect ---
    p_insp = sub.add_parser("inspect", help="检查模板里 md2word 能识别的样式/字体")
    p_insp.add_argument("template", help="待检查的 .docx 模板,或 @NAME 引用内置模板")
    p_insp.add_argument("--json", action="store_true", help="以 JSON 格式输出")

    # --- templates ---
    p_tpl = sub.add_parser("templates", help="管理内置模板:列出 / 导出 / 查看详情")
    tpl_sub = p_tpl.add_subparsers(dest="tpl_action")

    p_tpl_list = tpl_sub.add_parser("list", help="列出所有内置模板")
    p_tpl_list.add_argument("--json", action="store_true", help="以 JSON 格式输出")

    p_tpl_show = tpl_sub.add_parser("show", help="查看某个内置模板的详细说明")
    p_tpl_show.add_argument("name", help="模板名(如 gov_notice,可省略 @ 前缀)")

    p_tpl_eject = tpl_sub.add_parser(
        "eject",
        help="把内置模板复制到当前目录,方便你二次定制",
    )
    p_tpl_eject.add_argument("name", help="模板名(如 gov_notice)")
    p_tpl_eject.add_argument(
        "-o", "--output",
        help="导出路径(默认在当前目录,文件名与模板原始文件名一致)",
    )

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
    if args.command == "templates":
        return _run_templates(args)

    ap.print_help()
    return 0


def _resolve_template_ref(ref):
    """Resolve a template reference to a filesystem path.

    Accepts:
    - None → None (use builtin default)
    - "@name" or "name" matching a builtin → path inside the package
    - any other string → returned as-is (treated as filesystem path)

    Returns False (sentinel) when an @-name was given but not found —
    the caller prints a helpful error then exits non-zero.
    """
    if ref is None:
        return None
    if ref.startswith("@"):
        from .builtin_templates import by_name, BUILTINS
        bt = by_name(ref)
        if bt is None:
            print(f"错误: 找不到内置模板 {ref}。可用模板:", file=sys.stderr)
            for t in BUILTINS:
                print(f"  @{t.name} — {t.title}", file=sys.stderr)
            return False
        return bt.path
    return ref


def _run_convert(args) -> int:
    if not os.path.exists(args.input):
        print(f"错误: 找不到文件 {args.input}", file=sys.stderr); return 2

    tpl = _resolve_template_ref(args.template)
    if tpl is False:
        return 2
    if tpl and not os.path.exists(tpl):
        print(f"错误: 找不到模板 {tpl}", file=sys.stderr); return 2

    out = args.output or (os.path.splitext(args.input)[0] + ".docx")
    with open(args.input, "r", encoding="utf-8") as f:
        md_text = f.read()
    doc = parse(md_text)

    respect = None
    if args.override_template_fonts:
        respect = False
    Renderer(template=tpl, respect_template_fonts=respect).render(doc, out)
    print(f"✓ 已生成 {out}")
    return 0


def _run_inspect(args) -> int:
    from .inspect import inspect_template, format_report_text, format_report_json

    tpl = _resolve_template_ref(args.template)
    if tpl is False:
        return 2
    if not tpl or not os.path.exists(tpl):
        print(f"错误: 找不到文件 {args.template}", file=sys.stderr); return 2
    try:
        report = inspect_template(tpl)
    except Exception as e:
        print(f"错误: 无法读取模板 ({e})", file=sys.stderr); return 1

    print(format_report_json(report) if args.json else format_report_text(report))
    return 0


def _run_templates(args) -> int:
    from .builtin_templates import BUILTINS, by_name

    # `md2word templates` with no action defaults to list.
    action = args.tpl_action or "list"

    if action == "list":
        as_json = getattr(args, "json", False)
        if as_json:
            import json
            data = [
                {"name": t.name, "filename": t.filename, "title": t.title,
                 "description": t.description, "use_cases": t.use_cases,
                 "page_setup": t.page_setup, "fonts": t.fonts}
                for t in BUILTINS
            ]
            print(json.dumps(data, ensure_ascii=False, indent=2))
            return 0
        print("内置模板:")
        for t in BUILTINS:
            print(f"\n  @{t.name}  ({t.filename})")
            print(f"    {t.title}")
            print(f"    页面: {t.page_setup}")
            print(f"    字体: {t.fonts}")
            print(f"    用途: {', '.join(t.use_cases)}")
        print("\n使用方式:")
        print("  md2word convert 你的文档.md -t @模板名 -o 输出.docx")
        print("  md2word templates show 模板名   # 查看详情")
        print("  md2word templates eject 模板名  # 导出为本地文件以便二次定制")
        return 0

    if action == "show":
        bt = by_name(args.name)
        if bt is None:
            print(f"错误: 找不到内置模板 {args.name}", file=sys.stderr); return 2
        print(f"@{bt.name}")
        print(f"  文件:  {bt.filename}")
        print(f"  标题:  {bt.title}")
        print(f"  描述:  {bt.description}")
        print(f"  页面:  {bt.page_setup}")
        print(f"  字体:  {bt.fonts}")
        print(f"  适用:  {', '.join(bt.use_cases)}")
        print()
        # Show inspect output too — most users want to see what styles
        # are in there before they commit to using one.
        from .inspect import inspect_template, format_report_text
        report = inspect_template(bt.path)
        print(format_report_text(report))
        return 0

    if action == "eject":
        bt = by_name(args.name)
        if bt is None:
            print(f"错误: 找不到内置模板 {args.name}", file=sys.stderr); return 2
        dest = args.output or bt.filename
        if os.path.exists(dest):
            print(f"错误: 目标文件 {dest} 已存在,不覆盖。指定 -o 改个名字。",
                  file=sys.stderr)
            return 2
        shutil.copy(bt.path, dest)
        print(f"✓ 已导出 {bt.filename} 到 {dest}")
        print(f"  你现在可以在 Word 里打开它,调整字体/边距/添加红头等,")
        print(f"  保存后用 -t {dest} 来引用它。")
        return 0

    print(f"未知操作: {action}", file=sys.stderr); return 2


if __name__ == "__main__":
    sys.exit(main())
