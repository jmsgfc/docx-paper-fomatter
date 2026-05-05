#!/usr/bin/env python3
"""Verify that a formatted paper DOCX contains reusable Word numbering.

This script checks the structural pieces that weaker agents often miss:
numbering.xml, style-level numPr, numbering pStyle bindings, space suffixes,
and independent reference numbering. It does not judge the visual quality of
the whole paper.
"""

from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
HEADING_STYLES = ["PaperHeading1", "PaperHeading2", "PaperHeading3", "PaperHeading4"]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Verify formatted DOCX OOXML bindings.")
    parser.add_argument("--docx", required=True, help="Formatted DOCX path.")
    parser.add_argument("--report", help="Optional Markdown format-check report path.")
    return parser.parse_args()


def read_xml(package: zipfile.ZipFile, name: str, failures: list[str]) -> ET.Element | None:
    if name not in package.namelist():
        failures.append(f"Missing {name}")
        return None
    try:
        return ET.fromstring(package.read(name))
    except ET.ParseError as exc:
        failures.append(f"Invalid XML in {name}: {exc}")
        return None


def style_has_numpr(styles_root: ET.Element, style_id: str) -> bool:
    for style in styles_root.iter(W + "style"):
        if style.get(W + "styleId") == style_id:
            return style.find(".//" + W + "numPr") is not None
    return False


def has_style(styles_root: ET.Element, style_id: str) -> bool:
    return any(style.get(W + "styleId") == style_id for style in styles_root.iter(W + "style"))


def pstyle_bound(numbering_root: ET.Element, style_id: str) -> bool:
    for lvl in numbering_root.iter(W + "lvl"):
        for pstyle in lvl.findall(W + "pStyle"):
            if pstyle.get(W + "val") == style_id:
                return True
    return False


def count_space_suffixes(numbering_root: ET.Element) -> int:
    return sum(1 for suff in numbering_root.iter(W + "suff") if suff.get(W + "val") == "space")


def count_paragraph_numpr(document_root: ET.Element) -> int:
    return sum(1 for _ in document_root.iter(W + "numPr"))


def report_contains_numbering_check(report_path: Path) -> bool:
    if not report_path.exists():
        return False
    text = report_path.read_text(encoding="utf-8", errors="ignore")
    return "Numbering Check" in text or "编号" in text


def main() -> int:
    args = parse_args()
    docx_path = Path(args.docx)
    failures: list[str] = []
    notes: list[str] = []

    if not docx_path.exists():
        print(f"FAIL: DOCX does not exist: {docx_path}", file=sys.stderr)
        return 2

    try:
        package = zipfile.ZipFile(docx_path)
    except zipfile.BadZipFile:
        print(f"FAIL: Not a valid DOCX zip package: {docx_path}", file=sys.stderr)
        return 2

    with package:
        numbering = read_xml(package, "word/numbering.xml", failures)
        styles = read_xml(package, "word/styles.xml", failures)
        document = read_xml(package, "word/document.xml", failures)

    if numbering is None or styles is None or document is None:
        for failure in failures:
            print(f"FAIL: {failure}", file=sys.stderr)
        return 1

    for style_id in HEADING_STYLES:
        if not has_style(styles, style_id):
            failures.append(f"Missing heading style {style_id}")
        if not style_has_numpr(styles, style_id):
            failures.append(f"Heading style {style_id} does not contain numPr")
        if not pstyle_bound(numbering, style_id):
            failures.append(f"numbering.xml does not bind pStyle {style_id}")

    if count_space_suffixes(numbering) < len(HEADING_STYLES):
        failures.append("Expected at least four heading levels with w:suff=\"space\"")

    if not has_style(styles, "ReferenceBody"):
        failures.append("Missing ReferenceBody style")
    elif not style_has_numpr(styles, "ReferenceBody"):
        failures.append("ReferenceBody style does not contain numPr")

    paragraph_numpr_count = count_paragraph_numpr(document)
    if paragraph_numpr_count == 0:
        failures.append("No paragraph-level numPr found in document.xml")
    else:
        notes.append(f"paragraph-level numPr count: {paragraph_numpr_count}")

    if args.report:
        report_path = Path(args.report)
        if not report_path.exists():
            failures.append(f"Report file does not exist: {report_path}")
        elif not report_contains_numbering_check(report_path):
            failures.append("Report file does not contain a numbering check section")

    if failures:
        for failure in failures:
            print(f"FAIL: {failure}", file=sys.stderr)
        for note in notes:
            print(f"NOTE: {note}")
        return 1

    print(f"PASS: {docx_path}")
    print("PASS: heading styles PaperHeading1/2/3/4 are bound to reusable numbering")
    print("PASS: ReferenceBody is bound to automatic numbering")
    print(f"PASS: heading space suffix count = {count_space_suffixes(numbering)}")
    for note in notes:
        print(f"PASS: {note}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
