"""Command-line entry point."""
from __future__ import annotations

import argparse
import os
import sys

from .parser import parse
from .renderer import Renderer


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        prog="md2word",
        description="将 Markdown 转换为 Word 文档，按 reference.docx 模板应用样式。",
    )
    ap.add_argument("input", help="输入的 Markdown 文件")
    ap.add_argument("-o", "--output", help="输出 Word 路径（默认同名 .docx）")
    ap.add_argument(
        "-t", "--template",
        help="参考模板 reference.docx。不指定则使用内置中文模板 "
             "（A4 纸 / 宋体正文 / 黑体标题 / 首行缩进两字符 / 代码块带底纹）",
    )
    args = ap.parse_args(argv)

    if not os.path.exists(args.input):
        print(f"错误: 找不到文件 {args.input}", file=sys.stderr); return 2
    if args.template and not os.path.exists(args.template):
        print(f"错误: 找不到模板 {args.template}", file=sys.stderr); return 2

    out = args.output or (os.path.splitext(args.input)[0] + ".docx")
    with open(args.input, "r", encoding="utf-8") as f:
        md_text = f.read()
    doc = parse(md_text)
    Renderer(template=args.template).render(doc, out)
    print(f"✓ 已生成 {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
