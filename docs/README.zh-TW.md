# DOCX 論文格式規範 Skill

[English](../README.md) | [简体中文](./README.zh-CN.md) | [繁體中文](./README.zh-TW.md) | [日本語](./README.ja.md) | [한국어](./README.ko.md)

![Skill](https://img.shields.io/badge/Codex-Skill-0A7B83)
![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB)
![DOCX](https://img.shields.io/badge/Input-DOCX-2B579A)
![Output](https://img.shields.io/badge/Output-DOCX%20%7C%20Markdown%20%7C%20JSON-4E7A27)

根據教師格式要求文件，把學術論文 `.docx` 規範化為可重用的 Word 樣式、多層級標題編號、頁面設定與格式檢查輸出。

## 概述

這個 Codex skill 適合這樣的場景：你有教師要求或範本 `.docx`，也有待修改論文 `.docx`，希望用結構化、可重用的方式整理格式，而不是手動逐段調整格式。

## 主要功能

- 識別標題、摘要、關鍵詞、各級標題、正文與參考文獻
- 套用可重用的 Word 段落樣式
- 建立可重用的多層級標題自動編號
- 建立可重用的參考文獻自動編號
- 套用頁面版式設定
- 產生 Markdown 格式檢查報告
- 視需要匯出可重用 JSON 設定與 Schema 文件

## 安裝

將這個 skill 放入 `skills` 目錄中，並保持目錄名為 `docx-paper-formatter`。

```text
$CODEX_HOME/skills/docx-paper-formatter/
|-- SKILL.md
|-- agents/
|-- evals/
`-- scripts/
```

執行要求：

- Python 3.10+
- `lxml`

```powershell
pip install lxml
```

## 如何使用

你可以這樣向 Codex 發出請求：

- 「按老師要求整理這篇論文 DOCX」
- 「幫我把論文套用成可重用的 Word 樣式」
- 「不要手寫標題編號，改成 Word 多層級自動編號」
- 「從老師的 DOCX 裡提取格式規則並匯出可重用 JSON 設定」

典型輸入：

- 一份教師格式要求 `.docx`
- 一份待修改論文 `.docx`

典型輸出：

- 一份規範化後的 `.docx`
- 一份 Markdown 格式檢查報告
- 可選的 JSON 設定檔
- 可選的 JSON Schema 檔

## 腳本直接執行

格式化論文：

```powershell
python .\scripts\format_paper_docx.py `
  --teacher ".\teacher-format.docx" `
  --paper ".\paper.docx" `
  --out ".\paper-formatted.docx" `
  --report ".\format-report.md"
```

只匯出可重用設定：

```powershell
python .\scripts\format_paper_docx.py `
  --teacher ".\teacher-format.docx" `
  --paper ".\paper.docx" `
  --config-out ".\teacher-rules.json" `
  --schema-out ".\teacher-rules-schema.json" `
  --config-only
```

重用既有設定：

```powershell
python .\scripts\format_paper_docx.py `
  --paper ".\paper.docx" `
  --config ".\teacher-rules.json" `
  --out ".\paper-formatted.docx" `
  --report ".\format-report.md"
```

## 目前限制

這個 skill 主要聚焦於論文正文格式規範化與可重用編號體系。

它不打算完全自動化以下內容：

- 封面內容自動生成
- 猜測缺失的個人資訊
- 直接改寫或代寫論文正文
