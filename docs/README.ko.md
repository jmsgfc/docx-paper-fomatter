# DOCX 논문 포맷터 Skill

[English](../README.md) | [简体中文](./README.zh-CN.md) | [繁體中文](./README.zh-TW.md) | [日本語](./README.ja.md) | [한국어](./README.ko.md)

![Skill](https://img.shields.io/badge/Codex-Skill-0A7B83)
![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB)
![DOCX](https://img.shields.io/badge/Input-DOCX-2B579A)
![Output](https://img.shields.io/badge/Output-DOCX%20%7C%20Markdown%20%7C%20JSON-4E7A27)

교수님의 서식 요구 사항 문서에 따라 학술 논문 `.docx`를 재사용 가능한 Word 스타일, 다단계 제목 번호 매기기, 페이지 설정 및 서식 검사 출력으로 표준화합니다.

## 개요

이 Codex skill은 교수님의 요구 사항 또는 템플릿 `.docx`와 수정할 논문 `.docx`가 있고, 문단별로 서식을 수동으로 조정하는 대신 구조화되고 재사용 가능한 방식으로 서식을 정리하려는 시나리오에 적합합니다.

## 주요 기능

- 제목, 요약, 키워드, 각 수준 제목, 본문 및 참고 문헌 식별
- 재사용 가능한 Word 단락 스타일 적용
- 재사용 가능한 다단계 제목 자동 번호 매기기 생성
- 재사용 가능한 참고 문헌 자동 번호 매기기 생성
- 페이지 레이아웃 설정 적용
- Markdown 서식 검사 보고서 생성
- 필요에 따라 재사용 가능한 JSON 구성 및 Schema 파일 내보내기

## 설치

이 skill을 `skills` 디렉토리에 넣고 디렉토리 이름을 `docx-paper-formatter`로 유지합니다.

```text
$CODEX_HOME/skills/docx-paper-formatter/
|-- SKILL.md
|-- agents/
|-- evals/
`-- scripts/
```

실행 요구 사항:

- Python 3.10+
- `lxml`

```powershell
pip install lxml
```

## 사용 방법

Codex에 다음과 같이 요청할 수 있습니다:

- "교수님 요구 사항에 따라 이 논문 DOCX를 정리해 줘"
- "논문에 재사용 가능한 Word 스타일을 적용해 줘"
- "제목 번호를 수동으로 작성하지 말고 Word 다단계 자동 번호 매기기로 변경해 줘"
- "교수님의 DOCX에서 서식 규칙을 추출하고 재사용 가능한 JSON 구성을 내보내 줘"

일반적인 입력:

- 교수님 서식 요구 사항 `.docx`
- 수정할 논문 `.docx`

일반적인 출력:

- 표준화된 `.docx`
- Markdown 서식 검사 보고서
- 선택적인 JSON 구성 파일
- 선택적인 JSON Schema 파일

## 스크립트 직접 실행

논문 서식 지정:

```powershell
python .\scripts\format_paper_docx.py `
  --teacher ".\teacher-format.docx" `
  --paper ".\paper.docx" `
  --out ".\paper-formatted.docx" `
  --report ".\format-report.md"
```

재사용 가능한 구성만 내보내기:

```powershell
python .\scripts\format_paper_docx.py `
  --teacher ".\teacher-format.docx" `
  --paper ".\paper.docx" `
  --config-out ".\teacher-rules.json" `
  --schema-out ".\teacher-rules-schema.json" `
  --config-only
```

기존 구성 재사용:

```powershell
python .\scripts\format_paper_docx.py `
  --paper ".\paper.docx" `
  --config ".\teacher-rules.json" `
  --out ".\paper-formatted.docx" `
  --report ".\format-report.md"
```

## 현재 제한 사항

이 skill은 주로 논문 본문 서식 표준화 및 재사용 가능한 번호 체계에 중점을 둡니다.

다음 내용은 완전히 자동화하지 않습니다:

- 표지 내용 자동 생성
- 누락된 개인 정보 추측
- 논문 본문 직접 재작성 또는 대필
