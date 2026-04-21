"""Markdown → AST.

We rely on markdown-it-py for tokenization (it handles all the CommonMark
edge cases we don't want to debug), then fold its flat token stream into a
proper tree. Token folding is done with a small index cursor — no recursion
into the tokenizer itself.
"""
from __future__ import annotations

from typing import List, Tuple

from markdown_it import MarkdownIt

from .ast import (
    BulletList, CodeBlock, Document, Heading, Image, InlineImage, Link,
    ListItem, OrderedList, Paragraph, Quote, Table, TableRow, Text,
    ThematicBreak,
)


def parse(md_text: str) -> Document:
    md = MarkdownIt("commonmark", {"html": False}).enable("table").enable("strikethrough")
    tokens = md.parse(md_text)
    blocks, _ = _read_blocks(tokens, 0, stop_at=None)
    return Document(blocks=blocks)


# --- block-level folding --------------------------------------------------

def _read_blocks(tokens, i: int, stop_at) -> Tuple[list, int]:
    """Read blocks until we hit `stop_at` (a close tag) or run out."""
    out = []
    while i < len(tokens):
        tok = tokens[i]
        if stop_at and tok.type == stop_at:
            return out, i + 1

        t = tok.type
        if t == "heading_open":
            level = int(tok.tag[1])
            inlines = _read_inlines(tokens[i + 1])
            out.append(Heading(level=level, inlines=inlines))
            i += 3                              # heading_open, inline, heading_close
        elif t == "paragraph_open":
            inline_tok = tokens[i + 1]
            inlines = _read_inlines(inline_tok)
            # promote a lone image to a block-level Image
            if len(inlines) == 1 and isinstance(inlines[0], InlineImage):
                img = inlines[0]
                out.append(Image(src=img.src, alt=img.alt))
            else:
                out.append(Paragraph(inlines=inlines))
            i += 3
        elif t == "fence" or t == "code_block":
            out.append(CodeBlock(code=tok.content.rstrip("\n"), lang=(tok.info or "").strip()))
            i += 1
        elif t == "blockquote_open":
            children, i = _read_blocks(tokens, i + 1, stop_at="blockquote_close")
            out.append(Quote(children=children))
        elif t == "bullet_list_open":
            items, i = _read_list_items(tokens, i + 1, close="bullet_list_close")
            out.append(BulletList(items=items))
        elif t == "ordered_list_open":
            start = int(tok.attrGet("start") or 1)
            items, i = _read_list_items(tokens, i + 1, close="ordered_list_close")
            out.append(OrderedList(items=items, start=start))
        elif t == "hr":
            out.append(ThematicBreak()); i += 1
        elif t == "table_open":
            table, i = _read_table(tokens, i + 1)
            out.append(table)
        else:
            i += 1                              # unknown / already-consumed marker
    return out, i


def _read_list_items(tokens, i: int, close: str) -> Tuple[List[ListItem], int]:
    items: List[ListItem] = []
    while i < len(tokens) and tokens[i].type != close:
        if tokens[i].type == "list_item_open":
            children, i = _read_blocks(tokens, i + 1, stop_at="list_item_close")
            items.append(ListItem(children=children))
        else:
            i += 1
    return items, i + 1


def _read_table(tokens, i: int) -> Tuple[Table, int]:
    header = None
    rows: List[TableRow] = []
    while i < len(tokens) and tokens[i].type != "table_close":
        t = tokens[i].type
        if t == "thead_open":
            i += 1
            while tokens[i].type != "thead_close":
                if tokens[i].type == "tr_open":
                    cells, i = _read_tr(tokens, i + 1)
                    header = TableRow(cells=cells)
                else:
                    i += 1
            i += 1
        elif t == "tbody_open":
            i += 1
            while tokens[i].type != "tbody_close":
                if tokens[i].type == "tr_open":
                    cells, i = _read_tr(tokens, i + 1)
                    rows.append(TableRow(cells=cells))
                else:
                    i += 1
            i += 1
        else:
            i += 1
    return Table(header=header, rows=rows), i + 1


def _read_tr(tokens, i: int) -> Tuple[List[List], int]:
    cells: List[List] = []
    while tokens[i].type != "tr_close":
        if tokens[i].type in ("th_open", "td_open"):
            inline_tok = tokens[i + 1]
            cells.append(_read_inlines(inline_tok))
            i += 3                              # {th,td}_open, inline, {th,td}_close
        else:
            i += 1
    return cells, i + 1


# --- inline-level folding -------------------------------------------------

def _read_inlines(inline_tok) -> List:
    """markdown-it emits inline children as a flat list with *_open/*_close markers."""
    out: List = []
    stack_bold, stack_italic, stack_code = 0, 0, 0
    link_url = None
    kids = inline_tok.children or []
    i = 0
    while i < len(kids):
        c = kids[i]
        t = c.type
        if t == "text":
            if link_url is not None:
                out.append(Link(url=link_url, text=c.content))
            else:
                out.append(Text(
                    value=c.content,
                    bold=stack_bold > 0,
                    italic=stack_italic > 0,
                    code=False,
                ))
        elif t == "code_inline":
            out.append(Text(value=c.content, code=True))
        elif t == "strong_open":
            stack_bold += 1
        elif t == "strong_close":
            stack_bold -= 1
        elif t == "em_open":
            stack_italic += 1
        elif t == "em_close":
            stack_italic -= 1
        elif t == "link_open":
            link_url = c.attrGet("href") or ""
        elif t == "link_close":
            link_url = None
        elif t == "image":
            src = c.attrGet("src") or ""
            alt = c.content or ""
            out.append(InlineImage(src=src, alt=alt))
        elif t == "softbreak" or t == "hardbreak":
            out.append(Text(value="\n"))
        i += 1
    return out
