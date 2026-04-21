# md2word

一个面向中文场景的 Markdown → Word 转换工具。核心目标：**拿一份看着顺眼的 Word 模板，把 Markdown 扔进去，拿到一份打开几乎就能用的 docx。**

## 为什么又一个

Markdown 转 Word 的工具不少，但对中文用户来说，现有方案有几个绕不开的痛：

- `pandoc` 的 `--reference-doc` 机制对中文样式名识别不好，中英文字体槽位（`w:eastAsia`）也不处理，导出后经常要手动改字体。
- 很多基于 `python-docx` 的工具只改了 `ascii` / `hAnsi` 两个字体槽位，中文字符依然走模板的默认 CJK 字体，混排时错乱。
- 对中文文档最常见的几个细节——首行缩进两字符、A4 纸、宋体正文黑体标题——都需要用户自己折腾。

md2word 的定位是把这些细节做在默认路径上，让"开箱即用"真的成立；同时保留传入自定义模板的能力，让企业场景能用自己的品牌模板。

## 当前能做什么

- 标题 1–6 级，映射到模板的 `Heading 1..6`
- 段落、粗体、斜体、行内代码、真超链接（带 relationship，不是改颜色糊弄）
- 无序 / 有序列表（含嵌套）
- 引用块、代码块（`Code Block` 样式，等宽字体 + 浅灰底）
- 表格（自动识别表头，表头加粗）
- 块级图片（本地 / URL，自动按页面可用宽度缩放；文件找不到时降级为文字占位符）
- 水平分隔线
- 中英文字体槽位分别设置：`ascii` / `hAnsi` / `eastAsia` / `cs` 四个都写对
- 内置一份中文友好默认模板：A4 纸、宋体 / 黑体、小四号 1.5 倍行距、首行缩进两字符
- 支持传入自定义 `reference.docx` 模板

## 快速试用

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

需要 Python 3.9 或更新版本。不建议用 Python 3.14——部分依赖的预编译包还没跟进，会装不上。Python 3.11 或 3.12 是目前最稳的选择。

### 2. 跑一个样例

把整个项目解压或克隆下来，`cd` 到项目根目录（里面应该能看到 `md2word`、`tests`、`scripts` 三个子文件夹）：

```bash
python -m md2word.cli tests/fixtures/sample.md -o out.docx
```

看到 `✓ 已生成 out.docx` 就成了。打开 `out.docx`，应该看到标题、列表、表格、代码块、引用块、超链接都正确套上了样式。

### 3. 转换你自己的文档

```bash
# 用内置中文模板
python -m md2word.cli your.md -o out.docx

# 用你自己的企业模板
python -m md2word.cli your.md -t your_template.docx -o out.docx
```

输出路径可以省略，默认会在输入文件旁边生成同名的 `.docx`。

### 4. 跑测试

```bash
python -m unittest discover tests -v
```

应当看到 20 个测试全部通过。

## 用自定义模板的注意事项

md2word 加载模板时会扫描所有段落样式，按下面这个优先级去找每类内容该用的样式：

| Markdown 元素 | 查找顺序 |
|---------------|----------|
| 一级标题 | `Heading 1` → `标题 1` |
| 二级标题 | `Heading 2` → `标题 2` |
| ...（类推到 6 级） | |
| 普通段落 | `Normal` → `正文` |
| 无序列表 | `List Bullet` → `无序列表` |
| 有序列表 | `List Number` → `有序列表` |
| 引用块 | `Quote` → `Intense Quote` → `引用` |
| 代码块 | `Code Block` → `Code` → `代码块` → `HTML Preformatted` |
| 图注 | `Caption` → `图注` |

只要你的模板里有这些样式之一，就会被自动识别。找不到时会降级到最接近的样式，不会报错。

想让模板完整发挥作用，最简单的做法是：把内置模板 `md2word/templates/default_zh.docx` 复制一份，在 Word 里按你的品牌要求调整颜色、字号、页边距，保存为自己的模板传给 `-t` 参数。

## 项目结构

```
md2word/
├── README.md                   项目说明（本文件）
├── requirements.txt            运行时依赖
├── md2word/                    核心包
│   ├── ast.py                  文档节点定义
│   ├── parser.py               Markdown → AST
│   ├── renderer.py             AST → docx
│   ├── fonts.py                CJK 字体槽位处理
│   ├── cli.py                  命令行入口
│   └── templates/
│       └── default_zh.docx     内置中文默认模板
├── scripts/
│   └── build_default_template.py   生成默认模板的脚本
└── tests/
    ├── fixtures/               测试用的 Markdown 片段
    └── test_basic.py           结构性测试
```

整个核心包大约 650 行 Python（含注释和空行），刻意保持精简，方便阅读和修改。

## 设计选择

几个值得说明的取舍：

**不走"MD → HTML → docx"**。中间的 HTML 层会丢掉表格对齐、代码语言、图片尺寸等元信息。md2word 直接用 `markdown-it-py` 产出 token 流，折叠成自己的 AST，再落到 docx，信息路径最短。

**AST 独立于输入输出**。未来要支持别的输入（HTML、reStructuredText）或别的输出（ODT、RTF），不用动 AST 层。

**样式映射按约定而非配置**。第一版不提供"MD 元素 → 模板样式名"的配置文件。约定 `Heading 1..6` / `List Bullet` / `Quote` 这些标准名字，覆盖 95% 的模板。真要自定义样式名，再考虑引入配置。**过早抽象是没必要的复杂度。**

**CJK 字体是一等公民**。不是后续扩展，是核心路径的一部分。每一个 run 在写入时都显式设置四个字体槽位。

## 路线图

短期（真实使用暴露的问题优先）：

- 中文正文段落的排版细节微调（比如标题 / 列表 / 代码块之间的默认段间距）
- 代码块可选的语法高亮（基于 `pygments` 着色）
- 图片自动图注编号（`图 1-1` 格式，跟随章节）
- `frontmatter` 支持：从 MD 头部的 YAML 注入标题、作者、日期到封面

中期（依赖真实模板和真实反馈）：

- 数学公式：LaTeX → OMML 原生公式；降级链到图片 / Unicode
- 目录 / 交叉引用：自动插入带 SEQ 和 REF 域的 Word 字段
- 模板脚手架：`md2word scaffold` 生成起始模板 + 配置 + 示例 MD
- 更多场景模板：技术方案、周报、学习笔记、学位论文等预置模板库

暂不做（除非出现明确需求）：

- 插件 / Hook 机制——第一版保持精简，用户真遇到扩展需求再考虑
- Pandoc 式的全格式矩阵——md2word 只做 MD → Word 这一条路，做深而不做宽
- GUI——命令行优先，有真实痛点时再套壳

## 开发与贡献

加新功能之前先在 `tests/fixtures/` 放一个最小 MD 片段，在 `tests/test_basic.py` 加对应的 TestCase，然后再实现。这不是繁文缛节——写不出测试就说明功能还没想清楚。

改了默认模板的任何样式参数，需要重新跑一次模板生成脚本：

```bash
python scripts/build_default_template.py
```

## 许可

MIT
