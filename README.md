# Markdown转Word工具

这是一个简单的Python工具，用于将Markdown文件转换为Word文档，保留原始文档的结构和格式。

## 功能特点

- 支持Markdown基本语法（标题、段落、列表等）
- 支持表格转换
- 支持代码块转换
- 支持基本的文本格式（粗体、斜体等）

## 安装

1. 确保已安装Python 3.6或更高版本
2. 安装所需依赖：

```bash
pip install -r requirements.txt
```

## 使用方法

基本用法：

```bash
python markdown2word.py 输入的Markdown文件.md
```

指定输出文件：

```bash
python markdown2word.py 输入的Markdown文件.md -o 输出的Word文件.docx
```

## 示例

将示例Markdown文件转换为Word：

```bash
python markdown2word.py example.md
```

## 注意事项

- 目前不支持图片直接嵌入，只会显示图片描述
- 不支持复杂的HTML元素
- 超链接只会显示为蓝色下划线文本，不会保留实际链接功能

## 许可

本项目采用MIT许可证。