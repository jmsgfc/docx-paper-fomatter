# DOCX 论文格式规范 Skill

[English](../README.md) | [简体中文](./README.zh-CN.md) | [繁體中文](./README.zh-TW.md) | [日本語](./README.ja.md)

![Skill](https://img.shields.io/badge/Codex-Skill-0A7B83)
![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB)
![DOCX](https://img.shields.io/badge/Input-DOCX-2B579A)
![Output](https://img.shields.io/badge/Output-DOCX%20%7C%20Markdown%20%7C%20JSON-4E7A27)

根据教师格式要求文档，把学术论文 `.docx` 规范化为可复用的 Word 样式、多级标题编号、页面设置和格式检查输出。

## 概述

这个 Codex skill 适合这样的场景：你有教师要求或模板 `.docx`，也有待修改论文 `.docx`，希望用结构化、可复用的方式整理格式，而不是手工逐段调格式。

## 主要功能

- 识别标题、摘要、关键词、各级标题、正文和参考文献
- 应用可复用的 Word 段落样式
- 创建可复用的多级标题自动编号
- 创建可复用的参考文献自动编号
- 应用页面版式设置
- 生成 Markdown 格式检查报告
- 按需导出可复用 JSON 配置和 Schema 文件

## 安装

把这个仓库放到本地 Codex 的 skills 目录中，并保持目录名为 `docx-paper-formatter`。

```text
$CODEX_HOME/skills/docx-paper-formatter/
|-- SKILL.md
|-- agents/
|-- evals/
`-- scripts/
```

运行要求：

- Python 3.10+
- `lxml`

```powershell
pip install lxml
```

## 如何使用

你可以这样向 Codex 发出请求：

- “按老师要求整理这篇论文 DOCX”
- “帮我把论文套用成可复用的 Word 样式”
- “不要手写标题编号，改成 Word 多级自动编号”
- “从老师的 DOCX 里提取格式规则并导出可复用 JSON 配置”

典型输入：

- 一份教师格式要求 `.docx`
- 一份待修改论文 `.docx`

典型输出：

- 一份规范化后的 `.docx`
- 一份 Markdown 格式检查报告
- 可选的 JSON 配置文件
- 可选的 JSON Schema 文件

## 脚本直接运行

格式化论文：

```powershell
python .\scripts\format_paper_docx.py `
  --teacher ".\teacher-format.docx" `
  --paper ".\paper.docx" `
  --out ".\paper-formatted.docx" `
  --report ".\format-report.md"
```

只导出可复用配置：

```powershell
python .\scripts\format_paper_docx.py `
  --teacher ".\teacher-format.docx" `
  --paper ".\paper.docx" `
  --config-out ".\teacher-rules.json" `
  --schema-out ".\teacher-rules-schema.json" `
  --config-only
```

复用已有配置：

```powershell
python .\scripts\format_paper_docx.py `
  --paper ".\paper.docx" `
  --config ".\teacher-rules.json" `
  --out ".\paper-formatted.docx" `
  --report ".\format-report.md"
```

## 当前限制

这个 skill 主要聚焦于论文正文格式规范化和可复用编号体系。

它不打算完全自动化以下内容：

- 封面内容自动生成
- 猜测缺失的个人信息
- 直接改写或代写论文正文
