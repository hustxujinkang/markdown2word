# md2word

**把大模型生成的 Markdown,套上你单位的 Word 模板,一条命令生成可交付的文档。**

## 这个工具解决什么问题

现在越来越多人用豆包、DeepSeek、Claude 这些大模型写材料——提纲、周报、技术方案、分析报告。模型输出的是 Markdown,层次清晰、结构好。问题出在交付环节:

- 大模型自带的"导出 Word"功能,模板是写死的,样式单一,跟你单位的格式要求对不上
- 单位通常有规定的文档模板(公文红头、周报表头、技术方案书的章节样式、论文的字号行距),每份材料交付前都要花大量时间往模板里套格式
- 复制粘贴到模板里,标题级别、列表缩进、表格样式全部要重调,一份十几页的材料光排版就能耗掉一两个小时

md2word 的定位是解决这段"最后一公里":**给你一份 Markdown 原稿和一份你单位的 Word 模板,一条命令产出符合格式要求、可直接交付的文档。** 它假设你只接受你单位的格式,不做"自动布局",所有功能围绕"读模板 → 套样式 → 输出"展开。

## 没有模板?用我们的起步样本

如果你手头没有现成模板、只是想先把工具跑起来看看,我们捎带做了 5 套常见场景的起步样本——**不是这个项目的主打,是给"零成本上手"的人准备的样品**:

| 模板          | 适用场景                                | 一行命令                                    |
|--------------|----------------------------------------|---------------------------------------------|
| `@default`    | 通用中文文档                            | `md2word convert in.md`                     |
| `@gov_notice` | 党政机关公文                            | `md2word convert in.md -t @gov_notice`      |
| `@report`     | 周报 / 月报 / 季度总结 / 工作汇报       | `md2word convert in.md -t @report`          |
| `@tech_doc`   | 技术方案书 / PRD / 系统设计文档         | `md2word convert in.md -t @tech_doc`        |
| `@paper`      | 学术论文(通用样式,非特定高校)        | `md2word convert in.md -t @paper`           |

每个模板自带一份示例 markdown(`samples/` 目录),打开生成的 docx 立刻能看到样子。但**真正的预期用法是 `md2word convert your.md -t 你单位模板.docx`**——内置模板能覆盖到的格式总是有限的,你自己的模板才是要长期用的东西。

## 适用场景

- 机关单位的公文写作(通知、报告、请示)
- 公司内部的周报、月报、季度总结
- 技术方案书、产品需求文档、评审材料
- 学位论文、期刊投稿的格式适配
- 任何"内容用 AI 生成、格式按模板交付"的场景

## 快速试用

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

需要 Python 3.9 或更新版本。不建议用 Python 3.14——部分依赖的预编译包还没跟进,会装不上。Python 3.11 或 3.12 是目前最稳的选择。

### 2. 跑一个样例

`cd` 到项目根目录(里面应该能看到 `md2word`、`tests`、`scripts` 三个子文件夹):

```bash
python -m md2word.cli convert tests/fixtures/sample.md -o out.docx
```

看到 `✓ 已生成 out.docx` 就成了。打开 `out.docx` 应该看到标题、列表、表格、代码块、引用块、超链接都正确套上了样式。

> 省略 `convert` 子命令也可以(`python -m md2word.cli tests/fixtures/sample.md -o out.docx`),但推荐显式写,因为还有 `inspect` 和 `templates` 子命令。

### 3. 转换你自己的文档

```bash
# 用默认中文模板
python -m md2word.cli convert your.md -o out.docx

# 用内置模板
python -m md2word.cli convert your.md -t @report     -o 周报.docx
python -m md2word.cli convert your.md -t @tech_doc   -o 技术方案.docx
python -m md2word.cli convert your.md -t @gov_notice -o 公文.docx
python -m md2word.cli convert your.md -t @paper      -o 论文.docx

# 用你自己的单位模板
python -m md2word.cli convert your.md -t 单位模板.docx -o out.docx
```

输出路径可省略,默认在输入文件旁边生成同名 `.docx`。

## 内置模板

5 套常见中文办公场景的起步样本,通过 `-t @模板名` 引用,不需要本地有 .docx 文件。

### 模板一览

| `@` 名         | 文件                  | 主要参数                                                                          |
|---------------|----------------------|----------------------------------------------------------------------------------|
| `@default`    | `default_zh.docx`    | A4，宋体 12pt 正文 / 黑体阶梯标题 / 1.5 倍行距                                    |
| `@gov_notice` | `gov_notice_zh.docx` | A4,**边距 上33/下27/左27/右27 mm**,**固定 28 磅行距**,三号仿宋_GB2312 正文,标题层级方正小标宋简体 → 黑体 → 楷体_GB2312 → 仿宋_GB2312 → 仿宋_GB2312 |
| `@report`     | `report_zh.docx`     | A4，宋体 12pt 正文 / 黑体阶梯标题 / 小一黑体居中大标题 / 1.5 倍行距                |
| `@tech_doc`   | `tech_doc_zh.docx`   | A4，**20mm 窄边距**留空间放图表，微软雅黑做标题，代码块底色更明显，楷体图注       |
| `@paper`      | `paper_zh.docx`      | A4，**左边距 30mm** 留装订边，小二宋体加粗标题，三级黑体小标题，正文 1.5 倍行距    |

### 公文模板的标题层级映射

写公文时,Markdown 的标题级别按这套映射对应到公文格式:

| Markdown 写法           | 公文层级       | 字体              | 字号 | 加粗 | 对齐  |
|------------------------|----------------|-------------------|------|------|-------|
| `# 通知名`              | 文档主标题      | 方正小标宋简体   | 二号 (22pt) | 否   | 居中 |
| `## 一、xxx`            | 一级标题       | 黑体              | 三号 (16pt) | 否   | 左 |
| `### (一)xxx`         | 二级标题       | 楷体_GB2312       | 三号 (16pt) | 否   | 左 |
| `#### 1. xxx`           | 三级标题       | 仿宋_GB2312       | 三号 (16pt) | 否   | 左 |
| `##### (1)xxx`        | 四级标题       | 仿宋_GB2312       | 三号 (16pt) | 否   | 左 |
| 正文段落                | 正文           | 仿宋_GB2312       | 三号 (16pt) | —    | 两端对齐,首行缩进 2 字符 |

- 行距:正文/各级标题固定 28 磅,主标题固定 36 磅
- 段前段后:0
- 各级标题都按规范"不加粗",权重由字体本身的字形差异承担(黑体最重,楷体最轻)

> **关于"方正小标宋简体"**:这是商业授权字体,我们不能打包分发,但在模板的样式里写上了正确字体名。如果你的 Windows 装了这个字体(中国大陆机关单位常预装),Word 打开会用真正的方正小标宋;没装的话 Word 自动找最接近的中文标题字体替代,渲染依然规整,只是字面观感略有差异。

### 常用命令

```bash
# 列出所有内置模板（默认行为，也可以显式 list）
md2word templates

# 查看某个模板的详细信息（包含 inspect 报告）
md2word templates show gov_notice

# 把内置模板导出到当前目录，方便你在 Word 里二次定制
md2word templates eject gov_notice
md2word templates eject gov_notice -o 我的公文.docx

# 用模板渲染文档
md2word convert 我的通知.md -t @gov_notice -o 通知.docx
```

### 怎么二次定制

内置模板覆盖的是"格式骨架"——字号、行距、字体、边距、样式名都对得上。但每个单位还有自己特色的视觉元素：

- 公文的红头（单位徽标 + 单位名 + 红色反线）
- 公司模板的 logo、页眉页脚
- 学位论文的封面页、目录页
- 印章位置占位、骑缝章

这些**不在内置模板里**。我们的建议工作流是：

1. `md2word templates eject 模板名` 把内置模板导出
2. 在 Word 里打开它，加上单位特色元素（红头、logo、页眉等），另存为 `单位模板.docx`
3. 以后用 `md2word convert ... -t 单位模板.docx` 引用你自己的版本

这样你做一次定制，往后所有文档都用得上，且 md2word 不会改你模板里的格式（默认尊重模板字体，详见下文）。

### 关于字体

不同模板用的中文字体名是经过权衡的:

- **通用模板**(`@default` / `@report` / `@paper`)用最广泛可用的名字 —— 宋体 / 黑体 / 微软雅黑 / 楷体 —— 这些在 Windows + 中文版 Office 上几乎一定存在。
- **公文模板**(`@gov_notice`)用了精确的 GB2312 变体名 —— 仿宋_GB2312 / 楷体_GB2312 / 方正小标宋简体 —— 这是党政机关 Windows 上的标配,部分系统可能没装。

如果你的系统找不到对应字体,Word 会自动替换为最接近的中文字体。如果你要严格匹配带版权字体,请 `eject` 模板后在 Word 里手动替换为你单位的授权版本。

### 示例文档

`samples/` 目录下有每个模板对应的示例 Markdown，可以直接渲染对比：

```bash
md2word convert samples/gov_notice_demo.md -t @gov_notice
md2word convert samples/report_demo.md     -t @report
md2word convert samples/tech_doc_demo.md   -t @tech_doc
md2word convert samples/paper_demo.md      -t @paper
```

打开生成的 .docx 看效果——这是了解每个模板视觉差异最直接的方式。

## 工作机制

md2word 读你的 Word 模板,识别里面的段落样式,然后把 Markdown 的对应元素套上去。每类元素的样式查找顺序:

| Markdown 元素   | 按以下顺序找模板里的样式 |
|---------------|--------------------------|
| 一级到六级标题 | `Heading 1..6` → `标题 1..6` |
| 普通段落       | `Normal` → `正文`        |
| 无序列表       | `List Bullet` → `无序列表` |
| 有序列表       | `List Number` → `有序列表` |
| 引用块         | `Quote` → `Intense Quote` → `引用` |
| 代码块         | `Code Block` → `Code` → `代码块` → `HTML Preformatted` |
| 图注           | `Caption` → `图注`       |

模板里有这些名字之一就匹配。模板里没的(比如代码块样式),降级到 `Normal` 渲染。

**字体处理**:默认尊重模板的字体设置,只在样式缺 CJK 字体槽时帮补一个,避免中文回退到 Calibri。加 `--override-template-fonts` 可强制用内置字体覆盖。

**模板里的其他部分**(页眉页脚、嵌入图、自定义 XML、多级编号定义、theme 等)原样保留,不动。

## 检查你的模板能不能用

不确定模板里的样式是不是命对了名、字体是不是设全了?用 `inspect` 命令检查:

```bash
python -m md2word.cli inspect 单位模板.docx
```

会输出类似这样的报告:

```
模板文件: 单位模板.docx
页面:    210.0 × 297.0 mm  边距 上25.4 / 下25.4 / 左31.8 / 右31.8 mm
样式总数: 87

样式匹配
────────────────────────────────────────────────────────────
  ✓ 一级标题     → Heading 1
      字体: 西文 Arial, 中文 黑体, 16pt
  ✓ 正文        → Normal
      字体: 西文 Times New Roman, 中文 ⚠ 未设置, 12pt
  ✗ 代码块      未找到 (Code Block / Code / 代码块 / HTML Preformatted)
      → 将降级为 Normal/正文 样式

发现 2 个潜在问题:
  • 正文(Normal)未设置中文字体(eastAsia),中文可能回退到 Calibri
  • 代码块样式缺失,代码块将以普通段落渲染
```

加 `--json` 可以拿到结构化输出,便于做 CI 检查或集成到工作流里。

## Frontmatter:把元数据写进 Word 的"文件 → 信息"

Markdown 头部支持 YAML frontmatter,会被自动提取并写入 Word 文档的核心属性:

```markdown
---
title: 二〇二六年第一季度工作总结
author: 张三
date: 2026-04-01
version: v1.2
description: 季度工作回顾与下阶段计划
keywords: 季度总结, 工作回顾
---

# 第一季度总结

正文从这里开始……
```

支持的字段(其它字段会被静默忽略):`title` / `subject` / `author` / `keywords` / `description` / `category` / `version`。这些信息会出现在 Word 的"文件 → 信息"面板里,用于公文流转、文档检索、版本管理。

frontmatter 本身**不会**出现在正文里——它只更新文档属性,正文从 `---` 之后开始渲染。这样大模型可以稳定地把元数据写进 prompt 输出而不影响文档显示。

> **关于依赖**:完整 YAML 解析需要 PyYAML(已包含在 `requirements.txt` 里)。如果没装 PyYAML,会回退到一个简化的 `key: value` 解析器——只支持单行字符串,不支持列表 / 嵌套 / 多行。绝大多数公文场景的元数据用回退解析器也够。

## 路线图

### 已实现

- 6 级标题、段落、粗体/斜体/行内代码
- 真超链接(带 `w:hyperlink` relationship)
- 有序 / 无序列表 + 嵌套
- 引用块、代码块、表格(自动识别表头)
- 图片:本地路径 / URL / base64 data URI,自动按页面宽度缩放
- 水平分隔线
- YAML frontmatter → Word 核心属性
- 中英文字体槽位完整处理(ascii / hAnsi / eastAsia / cs)
- 自定义模板时默认尊重模板字体,只补缺失的 CJK 槽位
- `inspect` 命令:模板样式匹配、字体槽位、页面参数诊断(支持 `--json`)
- `templates` 子命令:list / show / eject 内置模板
- 5 套起步样本模板(`@default` / `@gov_notice` / `@report` / `@tech_doc` / `@paper`),每份内嵌示例内容
- 49 个端到端测试

### 待做

- 加载用户模板时自动跑健壮性检查(字体槽位、语言标记、theme 字段、basedOn 断链)
- 样式名模糊匹配:去空格、大小写归一、关键字包含
- `--style-map "Heading 1=我的一级标题"` 手动映射
- 列表多级编号优先复用模板的 `numbering.xml` 定义
- 表格样式按模板样式名查找,`Table Grid` 仅作兜底
- 失败时的可操作错误信息
- 真实 LLM 输出 fixtures 库
- 公文红头交互定制(eject 后填发文机关、文号、签发人)
- 图注 / 表注自动编号
- 代码块语法高亮
- 中英文标点自动空格
- 数学公式 LaTeX → OMML
- TOC 与交叉引用


## 项目结构

```
md2word/
├── README.md                   项目说明(本文件)
├── requirements.txt            运行时依赖(python-docx / markdown-it-py / PyYAML)
├── md2word/                    核心包
│   ├── ast.py                  文档节点定义
│   ├── parser.py               Markdown → AST(含 YAML frontmatter 提取)
│   ├── renderer.py             AST → docx
│   ├── fonts.py                中英文字体槽位处理
│   ├── inspect.py              模板诊断:样式匹配 / 字体槽位 / 页面布局检查
│   ├── builtin_templates.py    内置模板注册表(@name 到文件名的映射)
│   ├── cli.py                  命令行入口(convert / inspect / templates 子命令)
│   └── templates/              内置模板(开箱即用)
│       ├── default_zh.docx         通用中文文档
│       ├── gov_notice_zh.docx      党政机关公文(按用户单位规范定制)
│       ├── report_zh.docx          周报 / 月报 / 工作汇报
│       ├── tech_doc_zh.docx        技术方案 / PRD / 设计文档
│       └── paper_zh.docx           学术论文通用版
├── samples/                    每份模板对应的示例 Markdown
│   ├── gov_notice_demo.md
│   ├── report_demo.md
│   ├── tech_doc_demo.md
│   └── paper_demo.md
├── scripts/                    模板构建脚本
│   ├── _template_helpers.py        各 build 脚本共享的 OOXML 工具
│   ├── build_default_template.py
│   ├── build_gov_notice.py
│   ├── build_report.py
│   ├── build_tech_doc.py
│   ├── build_paper.py
│   └── build_all.py                一键全部生成
└── tests/
    ├── fixtures/               测试用的 Markdown 片段
    └── test_basic.py           结构性测试(49 个用例,覆盖所有特性)
```

整个核心包大约 2200 行 Python(其中相当一部分是注释和文档字符串),刻意保持精简,方便阅读和修改。

## 跑测试

```bash
python -m unittest discover tests -v
```

应当看到 49 个测试全部通过。

## 许可

MIT
