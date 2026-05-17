# md2word

**把大模型生成的 Markdown，一键转成符合你单位格式要求的 Word 文档。**

## 这个工具解决什么问题

现在越来越多人用豆包、DeepSeek、Claude 这些大模型写材料——提纲、周报、技术方案、分析报告。模型输出的是 Markdown，层次清晰、结构好。问题出在交付环节：

- 大模型自带的"导出 Word"功能，模板是写死的，样式单一，跟你单位的格式要求对不上
- 单位通常有规定的文档模板（公文红头、周报表头、技术方案书的章节样式、论文的字号行距），每份材料交付前都要花大量时间往模板里套格式
- 复制粘贴到模板里，标题级别、列表缩进、表格样式全部要重调，一份十几页的材料光排版就能耗掉一两个小时

md2word 的定位就是解决这段"最后一公里"：**只要你有一份 Markdown 原稿，再配一份单位的 Word 模板，一条命令生成符合格式要求、可直接交付的文档。**

或者，你完全可以不用自己的模板——**md2word 内置了 5 套常用文档模板，开箱即用**：

| 模板          | 适用场景                                | 一行命令                                    |
|--------------|----------------------------------------|---------------------------------------------|
| `@default`    | 通用中文文档                            | `md2word convert in.md`                     |
| `@gov_notice` | 党政机关公文（按 GB/T 9704）            | `md2word convert in.md -t @gov_notice`      |
| `@report`     | 周报 / 月报 / 季度总结 / 工作汇报       | `md2word convert in.md -t @report`          |
| `@tech_doc`   | 技术方案书 / PRD / 系统设计文档         | `md2word convert in.md -t @tech_doc`        |
| `@paper`      | 学术论文（通用样式，非特定高校）        | `md2word convert in.md -t @paper`           |

每套都设置好了符合场景的字号阶梯、行距、首行缩进、CJK 字体，且都通过 `md2word inspect` 的零警告检查。具体见下面的 [内置模板](#内置模板) 章节。

## 典型工作流

```
           ┌───────────────────┐
           │  大模型生成内容    │
           │  (Markdown 格式)   │
           └─────────┬─────────┘
                     │
           ┌─────────▼─────────┐
           │    md2word        │   (第一次使用某份模板前,
           │  + 你的模板(.docx) │    可先跑 md2word inspect
           └─────────┬─────────┘    检查兼容性)
                     │
           ┌─────────▼─────────┐
           │  符合格式要求的    │
           │  Word 文档,可交付  │
           └───────────────────┘
```

适用场景举例:

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

看到 `✓ 已生成 out.docx` 就成了。打开 `out.docx`,应该看到标题、列表、表格、代码块、引用块、超链接都正确套上了样式。

> **向后兼容**:省略 `convert` 子命令也可以,`python -m md2word.cli tests/fixtures/sample.md -o out.docx` 同样工作——这是为之前版本的用户保留的。新代码里建议显式写 `convert`,因为现在还有 `inspect` 子命令(见下文)。

### 3. 转换你自己的文档

```bash
# 用默认中文模板
python -m md2word.cli convert your.md -o out.docx

# 用内置模板(简写)
python -m md2word.cli convert your.md -t @report     -o 周报.docx
python -m md2word.cli convert your.md -t @tech_doc   -o 技术方案.docx
python -m md2word.cli convert your.md -t @gov_notice -o 公文.docx
python -m md2word.cli convert your.md -t @paper      -o 论文.docx

# 用你自己的单位模板
python -m md2word.cli convert your.md -t 单位模板.docx -o out.docx
```

输出路径可以省略,默认在输入文件旁边生成同名 `.docx`。

内置模板的完整清单和适用场景见下方[「内置模板」](#内置模板)一节。

### 4. 检查你的模板是否兼容(可选但强烈推荐)

第一次拿到一份单位模板,先 inspect 一下,看 md2word 能识别哪些样式、哪些会降级:

```bash
python -m md2word.cli inspect 单位模板.docx
```

具体输出和解读见下面 ["检查你的模板能不能用"](#检查你的模板能不能用) 一节。

## 内置模板

md2word 内置了 5 套针对常见中文办公场景的模板，全部存放在 `md2word/templates/` 下。每套模板都按对应规范精心调过字号、行距、首行缩进、CJK 字体槽位，可以直接通过 `-t @模板名` 引用，**不需要先解压、不需要本地有 .docx 文件**。

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
- **公文模板**(`@gov_notice`)用了**精确的 GB2312 变体名** —— 仿宋_GB2312 / 楷体_GB2312 / 方正小标宋简体 —— 这是党政机关 Windows 上的标配,但部分系统(macOS / Linux / 个人版 Windows)可能没装。

如果你的系统找不到对应字体,**Word 会自动替换为系统默认 CJK 字体**——视觉上会有差异,但不会乱码,**也不会变成日文字体**(下面这条修复保证)。

> **关于"标题怎么显示成 MS Gothic"的修复**:python-docx 创建的空文档默认 `w:lang w:eastAsia="en-US"`,且 `theme1.xml` 里的 CJK 字段是空字符串。这两件事叠加,会让 Word 把所有"未明确指定中文字体"的样式按日文(ja-JP)路径回退,最终落到 MS Gothic / MS Mincho。md2word 构建模板时三处都修了:(1) docDefaults 的 lang 改成 zh-CN;(2) 每个样式的 `w:eastAsiaTheme` 等 Theme 引用属性剥掉,让直接字体名生效;(3) theme1.xml 的 CJK 槽位填上正确字体。结果:即使遇到一个我们没主动配过的样式(比如 Word 自带的 Table Grid 系列),它走的回退链也是中文,不会出现日文字体。

如果你要严格匹配 GB/T 9704 要求的"方正小标宋"等带版权字体,请 `eject` 模板后在 Word 里手动替换为你单位的授权版本。

### 示例文档

`samples/` 目录下有每个模板对应的示例 Markdown，可以直接渲染对比：

```bash
md2word convert samples/gov_notice_demo.md -t @gov_notice
md2word convert samples/report_demo.md     -t @report
md2word convert samples/tech_doc_demo.md   -t @tech_doc
md2word convert samples/paper_demo.md      -t @paper
```

打开生成的 .docx 看效果——这是了解每个模板视觉差异最直接的方式。

## 怎么用好你的模板

md2word 的工作方式是:读取你的 Word 模板,识别里面的**段落样式**(比如 "Heading 1"、"正文"、"代码块"),然后把 Markdown 里对应的元素套上这些样式。

每类 Markdown 元素的样式查找顺序:

| Markdown 元素 | 按以下顺序找模板里的样式 |
|---------------|----------|
| 一级标题 | `Heading 1` → `标题 1` |
| 二级到六级标题 | `Heading 2..6` → `标题 2..6` |
| 普通段落 | `Normal` → `正文` |
| 无序列表 | `List Bullet` → `无序列表` |
| 有序列表 | `List Number` → `有序列表` |
| 引用块 | `Quote` → `Intense Quote` → `引用` |
| 代码块 | `Code Block` → `Code` → `代码块` |
| 图注 | `Caption` → `图注` |

**只要你的模板里有这些名字之一,就会自动匹配。** 所以准备模板时,只需要打开 Word 的"样式"面板,把你要的样式按这些名字命名好,md2word 就能认。

如果你没有自己的模板,可以直接用内置的中文默认模板(A4 纸、宋体正文、黑体标题、首行缩进两字符、1.5 倍行距),对大多数场景都够用。

### 检查你的模板能不能用

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

### 字体覆盖行为

- **使用内置模板时**(没传 `-t`):md2word 会按内置 FontProfile 把字体写到每个 run 上(西文 Times New Roman + 中文宋体)。
- **使用你自己的模板时**(传了 `-t`):**默认尊重模板里的字体设置**,只在样式没设 CJK 字体的时候帮你补一个,避免中文回退到 Calibri。如果你要强制覆盖,加 `--override-template-fonts`。

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

### 已经做好了

- ✅ Markdown 基础元素解析(标题 1-6 级、段落、粗体、斜体、行内代码)
- ✅ 真超链接(不是改颜色糊弄,是带 relationship 的 `w:hyperlink`)
- ✅ 无序 / 有序列表,含嵌套(深层级写入 `w:ilvl`,模板有多级编号会自动启用)
- ✅ 引用块(多段引用都正确套样式,不再只标记最后一段)
- ✅ 代码块(等宽字体 + 浅灰底)
- ✅ 表格(自动识别表头,表头加粗)
- ✅ 块级图片(本地 / URL / **base64 data URI**,自动按页面可用宽度缩放;找不到时降级为文字)
- ✅ URL 图片下载有 10 秒超时和 10 MB 上限,临时文件自动清理
- ✅ 水平分隔线
- ✅ 中英文字体槽位正确处理(ascii/hAnsi/eastAsia/cs 四个槽位都设置)
- ✅ **修复 Word 中文样式回退到日文字体的根因问题**(docDefaults lang、`*Theme` 属性剥离、theme1.xml CJK 槽位填充三处一并修)
- ✅ 自定义模板时**默认尊重模板字体**,只在缺失时补齐 CJK 槽位(`--override-template-fonts` 可强制覆盖)
- ✅ YAML frontmatter:自动写入 Word 文档核心属性
- ✅ **预置模板库:5 套常用文档模板,`-t @name` 一行命令调用**
  - ✅ 通用文档(`@default`)
  - ✅ 党政机关公文(`@gov_notice`,按用户单位规范定制字号/行距/字体/边距)
  - ✅ 周报 / 月报 / 工作汇报(`@report`)
  - ✅ 技术方案书 / PRD(`@tech_doc`)
  - ✅ 学术论文通用版(`@paper`)
- ✅ **每份模板内嵌示例内容**(打开就是样品,不需要先转换 markdown)
- ✅ `templates` 子命令:列出 / 查看详情 / 导出内置模板供二次定制
- ✅ `inspect` 命令:诊断模板样式匹配情况和字体设置
- ✅ 命令行工具(支持向后兼容的隐式 convert)
- ✅ 结构化自动测试(49 个测试用例)

### 近期计划

- ⬜ **更多场景模板**:学位论文(按具体高校)、合同协议、会议纪要、新闻稿
- ⬜ **公文红头一键定制**:`md2word templates eject @gov_notice` 后再加个交互流程,把发文机关、文号、签发人填进去就生成带红头的单位专属模板
- ⬜ **图注 / 表注自动编号**:"图 1-1"、"表 2-3" 格式,跟随章节自动递增
- ⬜ **代码块语法高亮**:基于 pygments 着色
- ⬜ **中英文标点自动规范**:中英文之间自动加空格(pangu 风格,可配置开关)
- ⬜ **真实 LLM 输出 fixtures 库**:收集豆包/DeepSeek/Claude/GPT 的真实输出做回归测试

### 更远的规划

- ⬜ **数学公式支持**:LaTeX → 原生 Word 公式(OMML);降级到图片或 Unicode
- ⬜ **目录与交叉引用**:自动插入 TOC 和章节、图表的交叉引用
- ⬜ **模板脚手架**:`md2word scaffold` 命令生成规范的起始模板 + 配置 + 示例
- ⬜ **更简单的界面**:考虑做一个拖拽式桌面工具,让不会命令行的文档工作者也能用

### 明确不做

- ❌ 插件 / 扩展机制——第一版保持精简,真出现扩展需求再说
- ❌ 支持 Markdown 以外的输入格式或 Word 以外的输出格式——只做 MD → Word 这一条路,做深不做宽
- ❌ 和 Pandoc 比拼全格式转换能力——我们只服务中文文档交付这一个场景

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
