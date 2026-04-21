"""md2word — 基于模板的 Markdown → Word 转换器。"""
from .parser import parse
from .renderer import Renderer, FontProfile
from .fonts import FontSpec

__all__ = ["parse", "Renderer", "FontProfile", "FontSpec"]
__version__ = "0.1.0"
