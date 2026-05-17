"""Document AST.

Minimal node types that cover 95% of day-to-day Markdown. Kept intentionally
small — new features should force you to think about whether they really
belong here, or in the renderer.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional, Union


# --- inline ---------------------------------------------------------------

@dataclass
class Text:
    value: str
    bold: bool = False
    italic: bool = False
    code: bool = False


@dataclass
class Link:
    url: str
    text: str


@dataclass
class InlineImage:
    src: str
    alt: str = ""


Inline = Union[Text, Link, InlineImage]


# --- block ----------------------------------------------------------------

@dataclass
class Heading:
    level: int                     # 1..6
    inlines: List[Inline]


@dataclass
class Paragraph:
    inlines: List[Inline]


@dataclass
class CodeBlock:
    code: str
    lang: str = ""


@dataclass
class Quote:
    children: List["Block"] = field(default_factory=list)


@dataclass
class ListItem:
    children: List["Block"] = field(default_factory=list)


@dataclass
class BulletList:
    items: List[ListItem] = field(default_factory=list)


@dataclass
class OrderedList:
    items: List[ListItem] = field(default_factory=list)
    start: int = 1


@dataclass
class TableRow:
    cells: List[List[Inline]]      # each cell is a list of inlines


@dataclass
class Table:
    header: Optional[TableRow]
    rows: List[TableRow] = field(default_factory=list)


@dataclass
class Image:                       # block-level (alone in a paragraph)
    src: str
    alt: str = ""


@dataclass
class ThematicBreak:
    pass


Block = Union[
    Heading, Paragraph, CodeBlock, Quote,
    BulletList, OrderedList, Table, Image, ThematicBreak,
]


@dataclass
class Document:
    blocks: List[Block] = field(default_factory=list)
    metadata: dict = field(default_factory=dict)
