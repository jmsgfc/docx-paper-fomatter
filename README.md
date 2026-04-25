# DOCX Paper Formatter

[English](./README.md) | [简体中文](./README.zh-CN.md) | [繁體中文](./README.zh-TW.md) | [日本語](./README.ja.md)

![Skill](https://img.shields.io/badge/Codex-Skill-0A7B83)
![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB)
![DOCX](https://img.shields.io/badge/Input-DOCX-2B579A)
![Output](https://img.shields.io/badge/Output-DOCX%20%7C%20Markdown%20%7C%20JSON-4E7A27)

Format academic `.docx` papers from teacher requirements into reusable Word styles, multilevel heading numbering, page setup, and format-check outputs.

## Overview

This Codex skill is for users who need to normalize a paper `.docx` against a teacher or template `.docx` without relying on one-off manual formatting.

## What It Does

- identifies title, abstract, keywords, headings, body text, and references
- applies reusable Word paragraph styles
- creates reusable multilevel heading numbering
- creates reusable automatic numbering for references
- applies page layout settings
- generates a Markdown format-check report
- optionally exports reusable JSON config and schema files

## Install

Place this repository in your local Codex skills directory and keep the folder name as `docx-paper-formatter`.

```text
$CODEX_HOME/skills/docx-paper-formatter/
|-- SKILL.md
|-- agents/
|-- evals/
`-- scripts/
```

Requirements:

- Python 3.10+
- `lxml`

```powershell
pip install lxml
```

## Use the Skill

Ask Codex with prompts like:

- "Format this paper DOCX according to the teacher requirements"
- "Apply reusable Word styles to this paper"
- "Use Word multilevel numbering instead of manual heading numbers"
- "Extract formatting rules from the teacher DOCX and export reusable JSON config"

Typical input:

- one teacher requirement `.docx`
- one paper `.docx`

Typical output:

- one formatted `.docx`
- one Markdown format-check report
- optional JSON config file
- optional JSON schema file

## Run the Script Directly

Format a paper:

```powershell
python .\scripts\format_paper_docx.py `
  --teacher ".\teacher-format.docx" `
  --paper ".\paper.docx" `
  --out ".\paper-formatted.docx" `
  --report ".\format-report.md"
```

Export reusable config only:

```powershell
python .\scripts\format_paper_docx.py `
  --teacher ".\teacher-format.docx" `
  --paper ".\paper.docx" `
  --config-out ".\teacher-rules.json" `
  --schema-out ".\teacher-rules-schema.json" `
  --config-only
```

Reuse an existing config:

```powershell
python .\scripts\format_paper_docx.py `
  --paper ".\paper.docx" `
  --config ".\teacher-rules.json" `
  --out ".\paper-formatted.docx" `
  --report ".\format-report.md"
```

## Limitations

This skill focuses on paper body formatting and reusable numbering.

It does not aim to fully automate:

- cover-page content generation
- guessing missing personal information
- writing or rewriting paper content
