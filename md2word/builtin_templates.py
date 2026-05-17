"""Registry of built-in templates.

Centralized so the CLI (`-t @name`, `md2word templates list`), the
README generator, and the test suite can all enumerate the same set
of templates without each maintaining its own list.

Adding a new template = adding one entry here + a build script under
scripts/. Nothing else needs to know about it.
"""
from __future__ import annotations

import os
from dataclasses import dataclass
from typing import List, Optional


_TEMPLATES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")


@dataclass
class BuiltinTemplate:
    name: str            # the @-name used on the CLI
    filename: str        # filename under md2word/templates/
    title: str           # short human label
    description: str     # one-paragraph blurb
    use_cases: List[str]
    page_setup: str
    fonts: str

    @property
    def path(self) -> str:
        return os.path.join(_TEMPLATES_DIR, self.filename)

    @property
    def exists(self) -> bool:
        return os.path.exists(self.path)


BUILTINS: List[BuiltinTemplate] = [
    BuiltinTemplate(
        name="default",
        filename="default_zh.docx",
        title="通用中文文档",
        description="A4 纸,宋体正文,黑体标题,首行缩进 2 字符,1.5 倍行距。"
                    "适合日常无特定格式要求的文档。",
        use_cases=["日常笔记", "通用文档", "不确定该用哪个时的安全选择"],
        page_setup="A4, 边距 2.54/2.54/3.18/3.18 cm",
        fonts="正文 宋体 12pt / 标题 黑体阶梯",
    ),
    BuiltinTemplate(
        name="gov_notice",
        filename="gov_notice_zh.docx",
        title="党政机关公文",
        description="按 GB/T 9704-2012 党政机关公文格式实现。"
                    "三号仿宋正文,固定 28 磅行距,版心 156×225 mm,"
                    "四级标题用 黑体 / 楷体 / 仿宋加粗 / 仿宋 区分。"
                    "frontmatter 支持 发文字号 / 主送 / 签发人 / 成文日期 等字段。",
        use_cases=["机关单位通知", "请示报告", "公文转发", "情况说明"],
        page_setup="A4, 边距 37/35/28/26 mm (GB/T 9704)",
        fonts="正文 仿宋 16pt / 标题 黑体·楷体·仿宋",
    ),
    BuiltinTemplate(
        name="report",
        filename="report_zh.docx",
        title="工作汇报",
        description="A4 纸,标题小一黑体居中,正文小四宋体,1.5 倍行距,"
                    "支持表格、引用块、代码块。日常高频使用的稳妥选择。",
        use_cases=["周报 / 月报 / 季度总结", "项目进展汇报",
                   "部门工作总结", "个人述职"],
        page_setup="A4, 边距 2.54/2.54/3.18/3.18 cm",
        fonts="正文 宋体 12pt / 标题 黑体阶梯",
    ),
    BuiltinTemplate(
        name="tech_doc",
        filename="tech_doc_zh.docx",
        title="技术文档 / PRD",
        description="窄边距 (20 mm) 留出空间放图表;微软雅黑做标题更现代;"
                    "代码块更显眼;楷体图注与正文区分明显。"
                    "适合面向工程师/PM 的文档。",
        use_cases=["技术方案书", "产品需求文档 (PRD)",
                   "系统设计文档", "评审材料"],
        page_setup="A4, 边距 20 mm 各边",
        fonts="正文 宋体 12pt / 标题 微软雅黑",
    ),
    BuiltinTemplate(
        name="paper",
        filename="paper_zh.docx",
        title="学术论文",
        description="通用中文学术论文样式:小二宋体标题,三级黑体小标题,"
                    "正文 1.5 倍行距首行缩进,左边距 30 mm 留装订边。"
                    "不针对特定高校,各校学位论文请用本校发的官方模板。",
        use_cases=["期刊投稿初稿", "课程论文 / 研讨班论文",
                   "项目结题报告", "学位论文章节草稿"],
        page_setup="A4, 边距 25/25/30/25 mm",
        fonts="正文 宋体 12pt / 标题 黑体",
    ),
]


def by_name(name: str) -> Optional[BuiltinTemplate]:
    """Look up a builtin by its @-name (case-insensitive)."""
    needle = name.lower().lstrip("@")
    for t in BUILTINS:
        if t.name.lower() == needle:
            return t
    return None


def all_names() -> List[str]:
    return [t.name for t in BUILTINS]
