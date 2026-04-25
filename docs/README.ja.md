# DOCX Paper Formatter

[English](./README.md) | [简体中文](./README.zh-CN.md) | [繁體中文](./README.zh-TW.md) | [日本語](./README.ja.md)

![Skill](https://img.shields.io/badge/Codex-Skill-0A7B83)
![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB)
![DOCX](https://img.shields.io/badge/Input-DOCX-2B579A)
![Output](https://img.shields.io/badge/Output-DOCX%20%7C%20Markdown%20%7C%20JSON-4E7A27)

教員の書式要件文書に基づいて、学術論文 `.docx` を再利用可能な Word スタイル、マルチレベル見出し番号、ページ設定、書式チェック出力へ正規化する Skill です。

## 概要

この Codex skill は、教員要件またはテンプレート `.docx` と修正対象の論文 `.docx` をもとに、手作業ではなく再利用可能な構造化書式で論文を整えたい場合に向いています。

## 主な機能

- タイトル、要旨、キーワード、各レベルの見出し、本文、参考文献を識別する
- 再利用可能な Word 段落スタイルを適用する
- 再利用可能なマルチレベル見出し自動番号を作成する
- 再利用可能な参考文献自動番号を作成する
- ページレイアウト設定を適用する
- Markdown 形式の書式チェックレポートを生成する
- 必要に応じて再利用可能な JSON 設定と Schema ファイルを出力する

## インストール

このリポジトリをローカルの Codex skills ディレクトリに配置し、フォルダ名を `docx-paper-formatter` のままにしてください。

```text
$CODEX_HOME/skills/docx-paper-formatter/
|-- SKILL.md
|-- agents/
|-- evals/
`-- scripts/
```

動作要件：

- Python 3.10+
- `lxml`

```powershell
pip install lxml
```

## 使い方

Codex には次のように依頼できます。

- 「先生の要件に従ってこの論文 DOCX を整形して」
- 「この論文に再利用可能な Word スタイルを適用して」
- 「見出し番号を手書きにせず、Word のマルチレベル自動番号にして」
- 「先生の DOCX から書式ルールを抽出して、再利用可能な JSON 設定を出力して」

典型的な入力：

- 教員の書式要件 `.docx` 1 件
- 修正対象の論文 `.docx` 1 件

典型的な出力：

- 整形済み `.docx` 1 件
- Markdown 形式の書式チェックレポート 1 件
- 任意の JSON 設定ファイル
- 任意の JSON Schema ファイル

## スクリプトを直接実行する

論文を整形する：

```powershell
python .\scripts\format_paper_docx.py `
  --teacher ".\teacher-format.docx" `
  --paper ".\paper.docx" `
  --out ".\paper-formatted.docx" `
  --report ".\format-report.md"
```

再利用可能な設定のみを出力する：

```powershell
python .\scripts\format_paper_docx.py `
  --teacher ".\teacher-format.docx" `
  --paper ".\paper.docx" `
  --config-out ".\teacher-rules.json" `
  --schema-out ".\teacher-rules-schema.json" `
  --config-only
```

既存の設定を再利用する：

```powershell
python .\scripts\format_paper_docx.py `
  --paper ".\paper.docx" `
  --config ".\teacher-rules.json" `
  --out ".\paper-formatted.docx" `
  --report ".\format-report.md"
```

## 制限事項

この skill は、論文本文の書式正規化と再利用可能な番号体系に主眼を置いています。

次の内容を完全自動化することは目的としていません。

- 表紙内容の自動生成
- 不足している個人情報の推測
- 論文本文の執筆や書き換え
