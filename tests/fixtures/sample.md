# 项目技术方案

本文档描述了 md2word 项目的 **技术选型** 与 *总体架构*。阅读对象为技术评审委员会成员。

## 背景与目标

当前团队在 Word 交付环节耗费大量时间处理排版。我们希望用一套 Markdown → Word 的自动化管线，把作者的精力留给内容本身。核心诉求包括：

- 支持按企业模板产出
- 中英文字体自动匹配
- 常见元素（表格、列表、代码块）排版不走样
- 面向非技术用户，开箱即用

## 总体架构

### 模块划分

整个工具划分为四个职责分明的层次：解析层负责把 Markdown 转成统一的 AST，渲染层负责把 AST 落到 docx。中间加一层模板抽象，把样式名与品牌模板解耦。

1. Parser 层：基于 `markdown-it-py`
2. AST 层：独立于输入输出的中间表示
3. Renderer 层：基于 `python-docx`
4. CLI 层：命令行入口

### 数据流

> 我们刻意避开了先转 HTML 再反解析这条路。HTML 层是信息损耗的根源——表格对齐、代码语言、图片尺寸这些元信息在走 HTML 时就丢掉了。

## 关键技术点

### 中文字体槽位

Word 的每个 run 有四个字体槽位：`w:ascii`、`w:hAnsi`、`w:eastAsia`、`w:cs`。python-docx 的 `run.font.name` 只设置前两个，中文字符实际使用的 `eastAsia` 槽位保持模板默认。这是 90% 基于 python-docx 的项目中英混排出问题的根源。

### 配套参数对比

| 参数 | 默认值 | 说明 |
|------|--------|------|
| 正文字体（中文） | 宋体 | GB/T 9704 推荐 |
| 正文字体（英文） | Times New Roman | 与中文字重匹配 |
| 标题字体（中文） | 黑体 | 视觉层级清晰 |
| 代码字体 | Consolas | 等宽，中英通用 |

## 代码示例

下面是字体应用的核心函数：

```python
def apply_font(run, spec):
    rpr = run._element.get_or_add_rPr()
    rfonts = OxmlElement("w:rFonts")
    rfonts.set(qn("w:ascii"), spec.latin)
    rfonts.set(qn("w:eastAsia"), spec.cjk)
    rpr.append(rfonts)
```

短短几行解决了一个长期被忽视的坑。

## 参考资料

详细的 OOXML 规范见 [ECMA-376 官方文档](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)。

---

*本方案经团队评审后发布。*
