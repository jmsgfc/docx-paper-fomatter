import argparse
import json
import re
from collections import Counter
from copy import deepcopy
from dataclasses import asdict, dataclass, field, replace
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
XML = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W, "r": R, "rel": REL, "ct": CT}

DEFAULT_TEACHER_DOCX = Path("2026.04作业格式要求.docx")
DEFAULT_PAPER_DOCX = Path("待修改论文文档.docx")
DEFAULT_OUTPUT_DOCX = Path("待修改论文文档_格式规范化_可复用多级标题.docx")
DEFAULT_REPORT_MD = Path("格式检查报告_可复用多级标题.md")
CONFIG_SCHEMA_VERSION = 1
ALLOWED_ALIGNMENTS = {"left", "right", "center", "both"}
ALLOWED_LINE_RULES = {"auto", "exact"}
ALLOWED_NUMBERING_SCHEMES = {"numeric", "chinese"}

SIZE_NAME_TO_HALF_POINTS = {
    "初号": "84",
    "小初": "72",
    "一号": "52",
    "小一": "48",
    "二号": "44",
    "小二": "36",
    "三号": "32",
    "小三": "30",
    "四号": "28",
    "小四": "24",
    "五号": "21",
    "小五": "18",
    "六号": "15",
    "小六": "13",
}


@dataclass
class StyleSpec:
    east_asia: Optional[str] = None
    ascii_font: Optional[str] = None
    size_half_points: Optional[str] = None
    bold: Optional[bool] = None
    jc: Optional[str] = None
    first_line_chars: Optional[str] = None
    outline_level: Optional[int] = None
    line_rule: Optional[str] = None
    line: Optional[str] = None
    before: Optional[str] = None
    after: Optional[str] = None


@dataclass
class PageSpec:
    width: str = "11906"
    height: str = "16838"
    top: str = "1134"
    right: str = "1304"
    bottom: str = "1134"
    left: str = "1304"
    header: str = "851"
    footer: str = "992"
    gutter: str = "0"


@dataclass
class TeacherConfig:
    title_style: StyleSpec = field(
        default_factory=lambda: StyleSpec(
            east_asia="黑体",
            ascii_font="Times New Roman",
            size_half_points="44",
            bold=True,
            jc="center",
            line_rule="auto",
            line="360",
        )
    )
    body_style: StyleSpec = field(
        default_factory=lambda: StyleSpec(
            east_asia="宋体",
            ascii_font="Times New Roman",
            size_half_points="24",
            bold=False,
            jc=None,
            first_line_chars="200",
            line_rule="exact",
            line="440",
        )
    )
    heading1_style: StyleSpec = field(default_factory=lambda: StyleSpec(outline_level=0))
    heading2_style: StyleSpec = field(default_factory=lambda: StyleSpec(outline_level=1))
    heading3_style: StyleSpec = field(default_factory=lambda: StyleSpec(outline_level=2))
    abstract_style: StyleSpec = field(default_factory=lambda: StyleSpec(first_line_chars="200"))
    keywords_style: StyleSpec = field(default_factory=lambda: StyleSpec(first_line_chars="200"))
    reference_title_style: StyleSpec = field(
        default_factory=lambda: StyleSpec(
            east_asia="黑体",
            ascii_font="Times New Roman",
            size_half_points="24",
            bold=True,
            jc="center",
            line_rule="exact",
            line="440",
        )
    )
    reference_body_style: StyleSpec = field(
        default_factory=lambda: StyleSpec(
            east_asia="宋体",
            ascii_font="Times New Roman",
            size_half_points="24",
            bold=False,
            first_line_chars="0",
            line_rule="exact",
            line="440",
        )
    )
    page: PageSpec = field(default_factory=PageSpec)
    numbering_scheme: str = "numeric"
    min_chars: int = 1500
    sources: Dict[str, str] = field(
        default_factory=lambda: {
            "title_style": "默认值",
            "body_style": "默认值",
            "heading1_style": "派生默认值",
            "heading2_style": "派生默认值",
            "heading3_style": "派生默认值",
            "page": "默认值",
            "numbering_scheme": "默认值",
            "min_chars": "默认值",
        }
    )
    warnings: List[str] = field(default_factory=list)


def qn(ns, tag):
    return f"{{{ns}}}{tag}"


def w(tag):
    return qn(W, tag)


def set_w_attr(el, name, value):
    el.set(w(name), str(value))


def child(parent, tag, **attrs):
    el = etree.SubElement(parent, w(tag))
    for key, value in attrs.items():
        set_w_attr(el, key, value)
    return el


def para_text(p):
    return "".join(p.xpath(".//w:t/text()", namespaces=NS)).strip()


def normalize_text(text):
    return re.sub(r"\s+", "", text or "")


def set_para_text(p, text):
    runs = p.findall("w:r", NS)
    if not runs:
        r = child(p, "r")
        t = child(r, "t")
        t.text = text
        return
    first_text_done = False
    for r in runs:
        for t in r.findall("w:t", NS):
            if not first_text_done:
                t.text = text
                if text.startswith(" ") or text.endswith(" "):
                    t.set(qn(XML, "space"), "preserve")
                first_text_done = True
            else:
                t.text = ""


def ensure_ppr(p):
    ppr = p.find("w:pPr", NS)
    if ppr is None:
        ppr = etree.Element(w("pPr"))
        p.insert(0, ppr)
    return ppr


def clean_paragraph_formatting(p):
    old_ppr = p.find("w:pPr", NS)
    if old_ppr is not None:
        p.remove(old_ppr)
    ppr = etree.Element(w("pPr"))
    p.insert(0, ppr)
    for r in p.xpath(".//w:r", namespaces=NS):
        rpr = r.find("w:rPr", NS)
        if rpr is not None:
            r.remove(rpr)
    return ppr


def apply_style(p, style_id):
    ppr = clean_paragraph_formatting(p)
    child(ppr, "pStyle", val=style_id)
    return ppr


def apply_style_and_numbering(p, style_id, ilvl, num_id):
    ppr = apply_style(p, style_id)
    num_pr = child(ppr, "numPr")
    child(num_pr, "ilvl", val=str(ilvl))
    child(num_pr, "numId", val=str(num_id))
    return ppr


def make_spacing(line_rule="exact", line="440"):
    spacing = etree.Element(w("spacing"))
    set_w_attr(spacing, "before", "0")
    set_w_attr(spacing, "after", "0")
    if line is not None:
        set_w_attr(spacing, "line", line)
    if line_rule is not None:
        set_w_attr(spacing, "lineRule", line_rule)
    return spacing


def make_fonts(east_asia="宋体", ascii_font="Times New Roman"):
    fonts = etree.Element(w("rFonts"))
    set_w_attr(fonts, "ascii", ascii_font)
    set_w_attr(fonts, "hAnsi", ascii_font)
    set_w_attr(fonts, "eastAsia", east_asia)
    set_w_attr(fonts, "cs", ascii_font)
    return fonts


def remove_style(styles_root, style_id):
    for style in styles_root.xpath(f"./w:style[@w:styleId='{style_id}']", namespaces=NS):
        styles_root.remove(style)


def add_paragraph_style(styles_root, style_id, name, spec, num_id=None, ilvl=None):
    remove_style(styles_root, style_id)
    style = etree.SubElement(styles_root, w("style"))
    set_w_attr(style, "type", "paragraph")
    set_w_attr(style, "styleId", style_id)
    set_w_attr(style, "customStyle", "1")
    child(style, "name", val=name)
    child(style, "basedOn", val="1")
    child(style, "qFormat")

    ppr = child(style, "pPr")
    spacing = make_spacing(line_rule=spec.line_rule, line=spec.line)
    if spec.before is not None:
        set_w_attr(spacing, "before", spec.before)
    if spec.after is not None:
        set_w_attr(spacing, "after", spec.after)
    ppr.append(spacing)
    if spec.jc:
        child(ppr, "jc", val=spec.jc)
    if spec.first_line_chars is not None:
        child(ppr, "ind", firstLineChars=str(spec.first_line_chars))
    if spec.outline_level is not None:
        child(ppr, "outlineLvl", val=str(spec.outline_level))
    if num_id is not None and ilvl is not None:
        num_pr = child(ppr, "numPr")
        child(num_pr, "ilvl", val=str(ilvl))
        child(num_pr, "numId", val=str(num_id))

    rpr = child(style, "rPr")
    rpr.append(make_fonts(east_asia=spec.east_asia, ascii_font=spec.ascii_font))
    child(rpr, "sz", val=spec.size_half_points)
    child(rpr, "szCs", val=spec.size_half_points)
    if spec.bold:
        child(rpr, "b")
        child(rpr, "bCs")


def ensure_content_type(parts, part_name, content_type):
    ct = etree.fromstring(parts["[Content_Types].xml"])
    if not ct.xpath(f"./ct:Override[@PartName='/{part_name}']", namespaces=NS):
        override = etree.SubElement(ct, qn(CT, "Override"))
        override.set("PartName", f"/{part_name}")
        override.set("ContentType", content_type)
    parts["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)


def ensure_document_relationship(parts, rel_type, target):
    rels_path = "word/_rels/document.xml.rels"
    rels = etree.fromstring(parts[rels_path])
    for rel in rels.findall(f"{{{REL}}}Relationship"):
        if rel.get("Type") == rel_type and rel.get("Target") == target:
            return rel.get("Id")
    nums = []
    for rel in rels.findall(f"{{{REL}}}Relationship"):
        rid = rel.get("Id", "")
        if rid.startswith("rId") and rid[3:].isdigit():
            nums.append(int(rid[3:]))
    rid = f"rId{max(nums, default=0) + 1}"
    rel = etree.SubElement(rels, qn(REL, "Relationship"))
    rel.set("Id", rid)
    rel.set("Type", rel_type)
    rel.set("Target", target)
    parts[rels_path] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
    return rid


def build_level_specs(numbering_scheme):
    if numbering_scheme == "chinese":
        return [
            (0, "PaperHeading1", "chineseCounting", "%1、", "left", "0", "420"),
            (1, "PaperHeading2", "chineseCounting", "（%2）", "left", "0", "420"),
            (2, "PaperHeading3", "decimal", "%3.", "left", "0", "420"),
        ]
    return [
        (0, "PaperHeading1", "decimal", "%1", "left", "0", "420"),
        (1, "PaperHeading2", "decimal", "%1.%2", "left", "0", "420"),
        (2, "PaperHeading3", "decimal", "%1.%2.%3", "left", "0", "420"),
    ]


def ensure_numbering(parts, numbering_scheme):
    if "word/numbering.xml" in parts:
        numbering = etree.fromstring(parts["word/numbering.xml"])
    else:
        numbering = etree.Element(w("numbering"), nsmap={"w": W})
        ensure_content_type(
            parts,
            "word/numbering.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
        )
        ensure_document_relationship(
            parts,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
            "numbering.xml",
        )

    used_abs = [
        int(x)
        for x in numbering.xpath("//w:abstractNum/@w:abstractNumId", namespaces=NS)
        if str(x).isdigit()
    ]
    used_num = [
        int(x)
        for x in numbering.xpath("//w:num/@w:numId", namespaces=NS)
        if str(x).isdigit()
    ]
    abstract_id = max(used_abs, default=0) + 1
    num_id = max(used_num, default=0) + 1

    abstract = etree.SubElement(numbering, w("abstractNum"))
    set_w_attr(abstract, "abstractNumId", str(abstract_id))
    child(abstract, "multiLevelType", val="multilevel")
    child(abstract, "name", val="PaperReusableHeadingNumbering")

    for ilvl, style_id, num_fmt, lvl_text, jc, left, hanging in build_level_specs(numbering_scheme):
        lvl = child(abstract, "lvl", ilvl=str(ilvl))
        child(lvl, "start", val="1")
        child(lvl, "numFmt", val=num_fmt)
        child(lvl, "pStyle", val=style_id)
        child(lvl, "lvlText", val=lvl_text)
        child(lvl, "suff", val="space")
        child(lvl, "lvlJc", val=jc)
        if ilvl > 0:
            child(lvl, "lvlRestart", val="1")
        ppr = child(lvl, "pPr")
        child(ppr, "ind", left=left, hanging=hanging)
        rpr = child(lvl, "rPr")
        rpr.append(make_fonts(east_asia="黑体"))

    num = etree.SubElement(numbering, w("num"))
    set_w_attr(num, "numId", str(num_id))
    child(num, "abstractNumId", val=str(abstract_id))
    parts["word/numbering.xml"] = etree.tostring(numbering, xml_declaration=True, encoding="UTF-8", standalone=True)
    return num_id, abstract_id


def ensure_reference_numbering(parts):
    numbering = etree.fromstring(parts["word/numbering.xml"])
    used_abs = [
        int(x)
        for x in numbering.xpath("//w:abstractNum/@w:abstractNumId", namespaces=NS)
        if str(x).isdigit()
    ]
    used_num = [
        int(x)
        for x in numbering.xpath("//w:num/@w:numId", namespaces=NS)
        if str(x).isdigit()
    ]
    abstract_id = max(used_abs, default=0) + 1
    num_id = max(used_num, default=0) + 1

    abstract = etree.SubElement(numbering, w("abstractNum"))
    set_w_attr(abstract, "abstractNumId", str(abstract_id))
    child(abstract, "multiLevelType", val="singleLevel")
    child(abstract, "name", val="PaperReferenceNumbering")

    lvl = child(abstract, "lvl", ilvl="0")
    child(lvl, "start", val="1")
    child(lvl, "numFmt", val="decimal")
    child(lvl, "pStyle", val="ReferenceBody")
    child(lvl, "lvlText", val="[%1]")
    child(lvl, "lvlJc", val="left")
    ppr = child(lvl, "pPr")
    child(ppr, "ind", left="0", hanging="420")
    rpr = child(lvl, "rPr")
    rpr.append(make_fonts(east_asia="宋体"))

    num = etree.SubElement(numbering, w("num"))
    set_w_attr(num, "numId", str(num_id))
    child(num, "abstractNumId", val=str(abstract_id))
    parts["word/numbering.xml"] = etree.tostring(numbering, xml_declaration=True, encoding="UTF-8", standalone=True)
    return num_id, abstract_id


def bind_reference_style_numbering(styles_root, reference_num_id):
    style = styles_root.xpath("./w:style[@w:styleId='ReferenceBody']", namespaces=NS)
    if not style:
        return
    ppr = style[0].find("w:pPr", NS)
    if ppr is None:
        ppr = child(style[0], "pPr")
    old_num_pr = ppr.find("w:numPr", NS)
    if old_num_pr is not None:
        ppr.remove(old_num_pr)
    num_pr = child(ppr, "numPr")
    child(num_pr, "ilvl", val="0")
    child(num_pr, "numId", val=str(reference_num_id))


def ensure_styles(styles_root, num_id, config):
    heading_font = config.title_style.east_asia or "黑体"
    body = config.body_style
    heading_base = replace(
        deepcopy(body),
        east_asia=heading_font,
        bold=True,
        first_line_chars=None,
    )
    heading1 = resolve_style(replace(deepcopy(heading_base), outline_level=0), config.heading1_style)
    heading2 = resolve_style(replace(deepcopy(heading_base), outline_level=1), config.heading2_style)
    heading3 = resolve_style(replace(deepcopy(heading_base), outline_level=2), config.heading3_style)
    abstract_style = replace(config.body_style, first_line_chars=config.body_style.first_line_chars or "200")
    keywords_style = replace(config.body_style, first_line_chars=config.body_style.first_line_chars or "200")
    reference_body_style = replace(config.body_style, first_line_chars="0")
    reference_title_style = replace(
        heading_base,
        jc=config.reference_title_style.jc or "center",
        outline_level=0,
    )

    add_paragraph_style(styles_root, "PaperTitle", "论文标题", config.title_style)
    add_paragraph_style(styles_root, "PaperHeading1", "论文一级标题", heading1, num_id=num_id, ilvl=0)
    add_paragraph_style(styles_root, "PaperHeading2", "论文二级标题", heading2, num_id=num_id, ilvl=1)
    add_paragraph_style(styles_root, "PaperHeading3", "论文三级标题", heading3, num_id=num_id, ilvl=2)
    add_paragraph_style(styles_root, "PaperBody", "论文正文", body)
    add_paragraph_style(styles_root, "AbstractStyle", "摘要", abstract_style)
    add_paragraph_style(styles_root, "KeywordsStyle", "关键词", keywords_style)
    add_paragraph_style(styles_root, "ReferenceTitle", "参考文献标题", reference_title_style)
    add_paragraph_style(styles_root, "ReferenceBody", "参考文献正文", reference_body_style)


def make_paragraph(text="", style_id="PaperBody"):
    p = etree.Element(w("p"))
    ppr = child(p, "pPr")
    child(ppr, "pStyle", val=style_id)
    if text:
        r = child(p, "r")
        t = child(r, "t")
        if text.startswith(" ") or text.endswith(" "):
            t.set(qn(XML, "space"), "preserve")
        t.text = text
    return p


def set_section_page(section, page_spec):
    for tag in ["pgSz", "pgMar"]:
        old = section.find(f"w:{tag}", NS)
        if old is not None:
            section.remove(old)
    pg_sz = etree.Element(w("pgSz"))
    set_w_attr(pg_sz, "w", page_spec.width)
    set_w_attr(pg_sz, "h", page_spec.height)
    section.insert(0, pg_sz)

    pg_mar = etree.Element(w("pgMar"))
    for key in ["top", "right", "bottom", "left", "header", "footer", "gutter"]:
        set_w_attr(pg_mar, key, getattr(page_spec, key))
    section.insert(1, pg_mar)


def make_section(page_spec, with_next_page=False):
    sect = etree.Element(w("sectPr"))
    if with_next_page:
        child(sect, "type", val="nextPage")
    set_section_page(sect, page_spec)
    return sect


def classify(text, index, in_references, previous_heading_ilvl=None):
    if index == 0:
        return "PaperTitle", "论文标题", None, text
    if text.startswith(("摘要：", "摘要:")):
        return "AbstractStyle", "摘要", None, text
    if text.startswith(("关键词：", "关键词:")):
        return "KeywordsStyle", "关键词", None, text
    if re.fullmatch(r"参考文献[:：]?", text):
        return "ReferenceTitle", "参考文献标题", None, "参考文献"
    if in_references:
        return "ReferenceBody", "参考文献正文", None, text
    if re.match(r"^\d+\.\d+\.\d+(?:\s|　|$)", text):
        clean = re.sub(r"^\d+\.\d+\.\d+(?:\s|　)*", "", text)
        return "PaperHeading3", "三级标题", 2, clean
    if re.match(r"^\d+\.\d+(?:\s|　|$)", text):
        clean = re.sub(r"^\d+\.\d+(?:\s|　)*", "", text)
        return "PaperHeading2", "二级标题", 1, clean
    if re.match(r"^\d+(?:\.|．)?(?:\s|　)+", text):
        clean = re.sub(r"^\d+(?:\.|．)?(?:\s|　)+", "", text)
        return "PaperHeading1", "一级标题", 0, clean
    if re.match(r"^[一二三四五六七八九十]+、", text):
        clean = re.sub(r"^[一二三四五六七八九十]+、\s*", "", text)
        return "PaperHeading1", "一级标题", 0, clean
    if re.match(r"^第[一二三四五六七八九十\d]+章", text):
        clean = re.sub(r"^第[一二三四五六七八九十\d]+章\s*", "", text)
        return "PaperHeading1", "一级标题", 0, clean
    if re.match(r"^（[一二三四五六七八九十]+）", text):
        clean = re.sub(r"^（[一二三四五六七八九十]+）\s*", "", text)
        return "PaperHeading2", "二级标题", 1, clean
    if re.match(r"^\d+、", text):
        clean = re.sub(r"^\d+、\s*", "", text)
        if previous_heading_ilvl == 1:
            return "PaperHeading3", "三级标题", 2, clean
        return "PaperHeading1", "一级标题", 0, clean
    return "PaperBody", "正文", None, text


def strip_reference_marker(text):
    return re.sub(r"^\s*(?:\[\d+\]|\(\d+\)|（\d+）|\d+[\.、])\s*", "", text)


def add_footer(parts, final_sect):
    existing = [name for name in parts if re.match(r"word/footer\d+\.xml$", name)]
    nums = [int(re.search(r"footer(\d+)\.xml", name).group(1)) for name in existing]
    footer_target = f"footer{max(nums, default=0) + 1}.xml"
    footer_name = f"word/{footer_target}"
    parts[footer_name] = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="{W}" xmlns:r="{R}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>'''.encode("utf-8")
    rid = ensure_document_relationship(
        parts,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
        footer_target,
    )
    ensure_content_type(
        parts,
        footer_name,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
    )
    for old in final_sect.findall("w:footerReference", NS):
        final_sect.remove(old)
    ref = etree.Element(w("footerReference"))
    set_w_attr(ref, "type", "default")
    ref.set(qn(R, "id"), rid)
    final_sect.insert(0, ref)
    old_pg = final_sect.find("w:pgNumType", NS)
    if old_pg is not None:
        final_sect.remove(old_pg)
    pg = etree.Element(w("pgNumType"))
    set_w_attr(pg, "start", "1")
    final_sect.append(pg)


def count_chars(texts):
    return len(re.findall(r"[\u4e00-\u9fffA-Za-z0-9]", "".join(texts)))


def remove_children(parent, tag):
    for item in list(parent.findall(f"w:{tag}", NS)):
        parent.remove(item)


def set_num_pr(parent_ppr, ilvl, num_id):
    remove_children(parent_ppr, "numPr")
    num_pr = etree.SubElement(parent_ppr, w("numPr"))
    child(num_pr, "ilvl", val=str(ilvl))
    child(num_pr, "numId", val=str(num_id))


def enforce_heading_numbering_continuity(document, styles_root, parts, num_id, abstract_id, numbering_scheme):
    numbering = etree.fromstring(parts["word/numbering.xml"])
    level_specs = {
        ilvl: (style_id, lvl_text, num_fmt)
        for ilvl, style_id, num_fmt, lvl_text, _, _, _ in build_level_specs(numbering_scheme)
    }

    abstract = numbering.find(f"./w:abstractNum[@w:abstractNumId='{abstract_id}']", NS)
    if abstract is not None:
        for ilvl, (style_id, lvl_text, num_fmt) in level_specs.items():
            lvl = abstract.find(f"./w:lvl[@w:ilvl='{ilvl}']", NS)
            if lvl is None:
                lvl = etree.SubElement(abstract, w("lvl"))
                set_w_attr(lvl, "ilvl", str(ilvl))
            remove_children(lvl, "numFmt")
            remove_children(lvl, "pStyle")
            remove_children(lvl, "lvlText")
            remove_children(lvl, "suff")
            child(lvl, "numFmt", val=num_fmt)
            child(lvl, "pStyle", val=style_id)
            child(lvl, "lvlText", val=lvl_text)
            child(lvl, "suff", val="space")
            if ilvl > 0 and lvl.find("w:lvlRestart", NS) is None:
                child(lvl, "lvlRestart", val="1")

    style_fixes = 0
    for ilvl, (style_id, _, _) in level_specs.items():
        style = styles_root.find(f"./w:style[@w:styleId='{style_id}']", NS)
        if style is None:
            continue
        ppr = style.find("w:pPr", NS)
        if ppr is None:
            ppr = child(style, "pPr")
        old_num_id = ppr.xpath("./w:numPr/w:numId/@w:val", namespaces=NS)
        old_ilvl = ppr.xpath("./w:numPr/w:ilvl/@w:val", namespaces=NS)
        if old_num_id != [str(num_id)] or old_ilvl != [str(ilvl)]:
            style_fixes += 1
        set_num_pr(ppr, ilvl, num_id)

    heading_count = 0
    paragraph_fixes = 0
    bad_after = 0
    for p in document.xpath("//w:body/w:p", namespaces=NS):
        ppr = p.find("w:pPr", NS)
        if ppr is None:
            continue
        style_id = ppr.xpath("./w:pStyle/@w:val", namespaces=NS)
        style_id = style_id[0] if style_id else ""
        if style_id not in {"PaperHeading1", "PaperHeading2", "PaperHeading3"}:
            continue
        heading_count += 1
        ilvl = {"PaperHeading1": 0, "PaperHeading2": 1, "PaperHeading3": 2}[style_id]
        old_num_id = ppr.xpath("./w:numPr/w:numId/@w:val", namespaces=NS)
        old_ilvl = ppr.xpath("./w:numPr/w:ilvl/@w:val", namespaces=NS)
        if old_num_id != [str(num_id)] or old_ilvl != [str(ilvl)]:
            paragraph_fixes += 1
        set_num_pr(ppr, ilvl, num_id)
        new_num_id = ppr.xpath("./w:numPr/w:numId/@w:val", namespaces=NS)
        if new_num_id != [str(num_id)]:
            bad_after += 1

    parts["word/numbering.xml"] = etree.tostring(numbering, xml_declaration=True, encoding="UTF-8", standalone=True)
    return {
        "heading_count": heading_count,
        "paragraph_fixes": paragraph_fixes,
        "style_fixes": style_fixes,
        "bad_heading_num_id_after": bad_after,
        "continuous_num_id": num_id,
    }


def read_docx_parts(docx_path):
    with ZipFile(docx_path, "r") as zin:
        return {name: zin.read(name) for name in zin.namelist()}


def paragraphs_from_body(document_root):
    body = document_root.find("w:body", NS)
    return [deepcopy(el) for el in body if el.tag == w("p")]


def extract_table_rows(document_root):
    rows = []
    for tbl in document_root.xpath("//w:body/w:tbl", namespaces=NS):
        table_rows = []
        for tr in tbl.xpath("./w:tr", namespaces=NS):
            cells = []
            for tc in tr.xpath("./w:tc", namespaces=NS):
                cell_texts = ["".join(p.xpath(".//w:t/text()", namespaces=NS)).strip() for p in tc.xpath(".//w:p", namespaces=NS)]
                cell_text = "\n".join(t for t in cell_texts if t)
                cells.append(cell_text)
            if any(cell.strip() for cell in cells):
                table_rows.append(cells)
        if table_rows:
            rows.append(table_rows)
    return rows


def parse_docx_root(docx_path):
    parts = read_docx_parts(docx_path)
    document = etree.fromstring(parts["word/document.xml"])
    return parts, document


def get_first(values):
    values = [v for v in values if v]
    return values[0] if values else None


def get_common(values):
    values = [v for v in values if v]
    if not values:
        return None
    return Counter(values).most_common(1)[0][0]


def get_max_numeric(values):
    numbers = [int(v) for v in values if str(v).isdigit()]
    return str(max(numbers)) if numbers else None


def merge_style(base, update):
    result = deepcopy(base)
    for field_name in StyleSpec.__dataclass_fields__:
        value = getattr(update, field_name)
        if value is not None:
            setattr(result, field_name, value)
    return result


def resolve_style(base, overlay):
    result = deepcopy(base)
    for field_name in StyleSpec.__dataclass_fields__:
        value = getattr(overlay, field_name)
        if value is not None:
            setattr(result, field_name, value)
    return result


def style_from_paragraph(p):
    ppr = p.find("w:pPr", NS)
    line = None
    line_rule = None
    jc = None
    first_line_chars = None
    before = None
    after = None
    if ppr is not None:
        line = get_first(ppr.xpath("./w:spacing/@w:line", namespaces=NS))
        line_rule = get_first(ppr.xpath("./w:spacing/@w:lineRule", namespaces=NS))
        before = get_first(ppr.xpath("./w:spacing/@w:before", namespaces=NS))
        after = get_first(ppr.xpath("./w:spacing/@w:after", namespaces=NS))
        jc = get_first(ppr.xpath("./w:jc/@w:val", namespaces=NS))
        first_line_chars = get_first(ppr.xpath("./w:ind/@w:firstLineChars", namespaces=NS))

    east_asia = get_common(p.xpath(".//w:rPr/w:rFonts/@w:eastAsia", namespaces=NS))
    ascii_font = get_common(p.xpath(".//w:rPr/w:rFonts/@w:ascii", namespaces=NS))
    size_half_points = get_max_numeric(p.xpath(".//w:rPr/w:sz/@w:val", namespaces=NS))
    bold = True if p.xpath(".//w:rPr/w:b", namespaces=NS) else None

    return StyleSpec(
        east_asia=east_asia,
        ascii_font=ascii_font,
        size_half_points=size_half_points,
        bold=bold,
        jc=jc,
        first_line_chars=first_line_chars,
        outline_level=None,
        line_rule=line_rule,
        line=line,
        before=before,
        after=after,
    )


def cm_to_twips(value):
    return str(int(round(float(value) * 1440 / 2.54)))


def twips_to_cm(value):
    return round(int(value) * 2.54 / 1440, 2)


def size_name_to_half_points(size_name):
    return SIZE_NAME_TO_HALF_POINTS.get(size_name)


def add_source(config, key, label):
    current = config.sources.get(key)
    if not current or current in {"默认值", "派生默认值"}:
        config.sources[key] = label
        return
    parts = [item.strip() for item in current.split("+")]
    if label not in parts:
        config.sources[key] = f"{current} + {label}"


def apply_body_rules_from_text(texts, config):
    joined = "\n".join(texts)

    body_line = next((t for t in texts if "内容用" in t and "宋体" in t), "")
    title_line = next((t for t in texts if "标题内容用" in t or ("标题" in t and "字体" in t)), "")

    if body_line:
        body_cn = re.search(r"内容用(黑体|宋体|仿宋|楷体|楷体_GB2312)(小初|初号|小一|一号|小二|二号|小三|三号|小四|四号|小五|五号|小六|六号)", body_line)
        body_en = re.search(r"(?:英文字用|英文用)([A-Za-z ]+?)(小初|初号|小一|一号|小二|二号|小三|三号|小四|四号|小五|五号|小六|六号)", body_line)
        config.body_style.bold = False
        config.body_style.jc = None
        if body_cn:
            config.body_style.east_asia = body_cn.group(1)
            config.body_style.size_half_points = size_name_to_half_points(body_cn.group(2)) or config.body_style.size_half_points
            add_source(config, "body_style", "教师要求文本")
        if body_en:
            config.body_style.ascii_font = body_en.group(1).strip()
            config.body_style.size_half_points = size_name_to_half_points(body_en.group(2)) or config.body_style.size_half_points
            add_source(config, "body_style", "教师要求文本")

        line_match = re.search(r"固定值\s*(\d+(?:\.\d+)?)\s*磅", body_line)
        if line_match:
            config.body_style.line_rule = "exact"
            config.body_style.line = str(int(round(float(line_match.group(1)) * 20)))
            add_source(config, "body_style", "教师要求文本")

        indent_match = re.search(r"缩进\s*(\d+)\s*字符", body_line)
        if indent_match:
            config.body_style.first_line_chars = str(int(indent_match.group(1)) * 100)
            add_source(config, "body_style", "教师要求文本")

    if title_line:
        title_font = re.search(r"标题[^。；\n]*?用(黑体|宋体|仿宋|楷体|楷体_GB2312)(小初|初号|小一|一号|小二|二号|小三|三号|小四|四号|小五|五号|小六|六号)", title_line)
        if title_font:
            config.title_style.east_asia = title_font.group(1)
            config.title_style.ascii_font = config.body_style.ascii_font
            config.title_style.size_half_points = size_name_to_half_points(title_font.group(2)) or config.title_style.size_half_points
            add_source(config, "title_style", "教师要求文本")
        if "加粗" in title_line:
            config.title_style.bold = True
            add_source(config, "title_style", "教师要求文本")
        if re.search(r"1\.5\s*倍行距", title_line):
            config.title_style.line_rule = "auto"
            config.title_style.line = "360"
            add_source(config, "title_style", "教师要求文本")

    margin_matches = {}
    for key, label in [("top", "上"), ("bottom", "下"), ("left", "左"), ("right", "右")]:
        match = re.search(rf"{label}[：:]\s*(\d+(?:\.\d+)?)\s*厘米", joined)
        if match:
            margin_matches[key] = cm_to_twips(match.group(1))
    if margin_matches:
        for key, value in margin_matches.items():
            setattr(config.page, key, value)
        add_source(config, "page", "教师要求文本")

    chars_match = re.search(r"(\d+)\s*字以上", joined)
    if chars_match:
        config.min_chars = int(chars_match.group(1))
        add_source(config, "min_chars", "教师要求文本")


def parse_alignment(text):
    if "居中" in text:
        return "center"
    if "左对齐" in text:
        return "left"
    if "右对齐" in text:
        return "right"
    if "两端对齐" in text:
        return "both"
    return None


def detect_heading_level(text):
    normalized = normalize_text(text)
    if "一级标题" in normalized or normalized in {"一级", "1级", "标题1"}:
        return 1
    if "二级标题" in normalized or normalized in {"二级", "2级", "标题2"}:
        return 2
    if "三级标题" in normalized or normalized in {"三级", "3级", "标题3"}:
        return 3
    return None


def parse_spacing_twips(text, label):
    pattern = rf"{label}\s*(\d+(?:\.\d+)?)\s*(磅|pt|行)" if label else r"(\d+(?:\.\d+)?)\s*(磅|pt|行)"
    match = re.search(pattern, text, re.IGNORECASE)
    if not match:
        return None
    value = float(match.group(1))
    unit = match.group(2).lower()
    if unit in {"磅", "pt"}:
        return str(int(round(value * 20)))
    return str(int(round(value * 240)))


def parse_font_and_size(text):
    match = re.search(
        r"(黑体|宋体|仿宋|楷体|楷体_GB2312|Arial|Calibri|Times New Roman)"
        r"\s*(小初|初号|小一|一号|小二|二号|小三|三号|小四|四号|小五|五号|小六|六号)?",
        text,
        re.IGNORECASE,
    )
    if not match:
        return None, None
    font = match.group(1)
    size_name = match.group(2)
    size = size_name_to_half_points(size_name) if size_name else None
    return font, size


def parse_size_only(text):
    match = re.search(r"(小初|初号|小一|一号|小二|二号|小三|三号|小四|四号|小五|五号|小六|六号)", text)
    if not match:
        return None
    return size_name_to_half_points(match.group(1))


def parse_line_value(text):
    exact = re.search(r"(固定值|行距)\s*(\d+(?:\.\d+)?)\s*磅", text)
    if exact:
        return "exact", str(int(round(float(exact.group(2)) * 20)))
    if re.search(r"1\.5\s*倍行距", text):
        return "auto", "360"
    if re.search(r"2\s*倍行距", text):
        return "auto", "480"
    return None, None


def parse_heading_rule_line(text):
    for level, label in [(1, "一级标题"), (2, "二级标题"), (3, "三级标题")]:
        if label not in text:
            continue
        font, size = parse_font_and_size(text)
        spec = StyleSpec(outline_level=level - 1)
        if font:
            spec.east_asia = font
            if re.search(r"[A-Za-z]", font):
                spec.ascii_font = font
        if size:
            spec.size_half_points = size
        if "加粗" in text:
            spec.bold = True
        elif "不加粗" in text:
            spec.bold = False
        align = parse_alignment(text)
        if align:
            spec.jc = align
        before = parse_spacing_twips(text, "段前")
        after = parse_spacing_twips(text, "段后")
        if before is not None:
            spec.before = before
        if after is not None:
            spec.after = after
        line_rule, line = parse_line_value(text)
        if line_rule and line:
            spec.line_rule = line_rule
            spec.line = line
        return level, spec
    return None, None


def apply_heading_rules_from_text(texts, config):
    found = False
    for text in texts:
        level, spec = parse_heading_rule_line(text)
        if not level or spec is None:
            continue
        current = getattr(config, f"heading{level}_style")
        setattr(config, f"heading{level}_style", merge_style(current, spec))
        add_source(config, f"heading{level}_style", "教师要求文本")
        found = True
    return found


def infer_table_header_map(header_cells):
    header_map = {}
    for idx, cell in enumerate(header_cells):
        normalized = normalize_text(cell).lower()
        if not normalized:
            continue
        if any(token in normalized for token in ["级别", "层级", "标题级", "标题层级"]):
            header_map[idx] = "level"
        elif "字体" in normalized:
            header_map[idx] = "font"
        elif "字号" in normalized:
            header_map[idx] = "size"
        elif "对齐" in normalized:
            header_map[idx] = "alignment"
        elif "段前" in normalized:
            header_map[idx] = "before"
        elif "段后" in normalized:
            header_map[idx] = "after"
        elif "加粗" in normalized or "字重" in normalized:
            header_map[idx] = "bold"
        elif "行距" in normalized:
            header_map[idx] = "line"
    return header_map


def parse_heading_rule_cells(cells, header_map=None):
    header_map = header_map or {}
    joined = " ".join(cell for cell in cells if cell)
    level = None
    spec = StyleSpec()
    if header_map:
        for idx, cell in enumerate(cells):
            text = cell.strip()
            if not text:
                continue
            field_name = header_map.get(idx)
            if field_name == "level":
                level = detect_heading_level(text) or level
                continue
            if field_name == "font":
                font, size = parse_font_and_size(text)
                if font:
                    spec.east_asia = font
                    if re.search(r"[A-Za-z]", font):
                        spec.ascii_font = font
                if size:
                    spec.size_half_points = size
                continue
            if field_name == "size":
                spec.size_half_points = parse_size_only(text) or spec.size_half_points
                continue
            if field_name == "alignment":
                spec.jc = parse_alignment(text) or spec.jc
                continue
            if field_name == "before":
                spec.before = parse_spacing_twips(text, "") or spec.before
                continue
            if field_name == "after":
                spec.after = parse_spacing_twips(text, "") or spec.after
                continue
            if field_name == "bold":
                normalized = normalize_text(text)
                if normalized in {"是", "加粗", "bold", "true", "1"}:
                    spec.bold = True
                elif normalized in {"否", "不加粗", "normal", "false", "0"}:
                    spec.bold = False
                elif "不" in text and "粗" in text:
                    spec.bold = False
                elif "粗" in text:
                    spec.bold = True
                continue
            if field_name == "line":
                line_rule, line = parse_line_value(text)
                if line_rule and line:
                    spec.line_rule = line_rule
                    spec.line = line
                continue
            level = detect_heading_level(text) or level
    else:
        fallback_level, fallback_spec = parse_heading_rule_line(joined)
        if fallback_level:
            return fallback_level, fallback_spec

    if not level:
        level = detect_heading_level(joined)
    if not level:
        return None, None

    if spec.east_asia is None or spec.size_half_points is None:
        font, size = parse_font_and_size(joined)
        if font and spec.east_asia is None:
            spec.east_asia = font
            if re.search(r"[A-Za-z]", font) and spec.ascii_font is None:
                spec.ascii_font = font
        if size and spec.size_half_points is None:
            spec.size_half_points = size
    if spec.jc is None:
        spec.jc = parse_alignment(joined)
    if spec.before is None:
        spec.before = parse_spacing_twips(joined, "段前")
    if spec.after is None:
        spec.after = parse_spacing_twips(joined, "段后")
    if spec.bold is None:
        if "不加粗" in joined:
            spec.bold = False
        elif "加粗" in joined:
            spec.bold = True
    if spec.line_rule is None or spec.line is None:
        line_rule, line = parse_line_value(joined)
        if line_rule and line:
            spec.line_rule = line_rule
            spec.line = line
    spec.outline_level = level - 1
    return level, spec


def apply_heading_rules_from_tables(teacher_tables, config):
    found = False
    for table in teacher_tables:
        header_map = {}
        data_rows = table
        if table:
            candidate_map = infer_table_header_map(table[0])
            if candidate_map:
                header_map = candidate_map
                data_rows = table[1:]
        for row in data_rows:
            level, spec = parse_heading_rule_cells(row, header_map)
            if not level or spec is None:
                continue
            current = getattr(config, f"heading{level}_style")
            setattr(config, f"heading{level}_style", merge_style(current, spec))
            add_source(config, f"heading{level}_style", "教师要求表格")
            found = True
    return found


def apply_teacher_example_page(document_root, config):
    sect = document_root.find(".//w:body/w:sectPr", NS)
    if sect is None:
        return
    page_changed = False
    pg_sz = sect.find("w:pgSz", NS)
    if pg_sz is not None:
        width = pg_sz.get(w("w"))
        height = pg_sz.get(w("h"))
        if width and height:
            config.page.width = width
            config.page.height = height
            page_changed = True
    pg_mar = sect.find("w:pgMar", NS)
    if pg_mar is not None:
        for key in ["top", "right", "bottom", "left", "header", "footer", "gutter"]:
            value = pg_mar.get(w(key))
            if value:
                if key in {"top", "right", "bottom", "left"} and config.sources["page"] == "教师要求文本":
                    continue
                setattr(config.page, key, value)
                page_changed = True
    if page_changed and config.sources["page"] == "默认值":
        add_source(config, "page", "教师示例文档节属性")

def apply_teacher_example_styles(teacher_paragraphs, config):
    sample_map = {}
    for p in teacher_paragraphs:
        text = para_text(p)
        if not text:
            continue
        if text.startswith(("摘要：", "摘要:")) and "abstract" not in sample_map:
            sample_map["abstract"] = p
        elif text.startswith(("关键词：", "关键词:")) and "keywords" not in sample_map:
            sample_map["keywords"] = p
        elif text.startswith("正文：") and "body" not in sample_map:
            sample_map["body"] = p
        elif re.fullmatch(r"参考文献[:：]?", text) and "reference_title" not in sample_map:
            sample_map["reference_title"] = p

    if "body" in sample_map:
        config.body_style = merge_style(config.body_style, style_from_paragraph(sample_map["body"]))
        add_source(config, "body_style", "教师示例正文")
    if "abstract" in sample_map:
        config.abstract_style = merge_style(config.body_style, style_from_paragraph(sample_map["abstract"]))
    if "keywords" in sample_map:
        config.keywords_style = merge_style(config.body_style, style_from_paragraph(sample_map["keywords"]))
    if "reference_title" in sample_map:
        config.reference_title_style = merge_style(config.reference_title_style, style_from_paragraph(sample_map["reference_title"]))


def apply_teacher_heading_examples(teacher_paragraphs, config):
    for p in teacher_paragraphs:
        text = para_text(p)
        if not text:
            continue
        normalized = text.replace(" ", "")
        for level, label in [(1, "一级标题"), (2, "二级标题"), (3, "三级标题")]:
            if label not in normalized:
                continue
            if any(skip in normalized for skip in ["格式要求", "字体", "字号", "段前", "段后", "行距"]):
                continue
            current = getattr(config, f"heading{level}_style")
            extracted = style_from_paragraph(p)
            setattr(config, f"heading{level}_style", merge_style(current, extracted))
            add_source(config, f"heading{level}_style", "教师示例标题")


def infer_numbering_scheme(teacher_texts, paper_texts, config):
    teacher_joined = "\n".join(teacher_texts)
    if re.search(r"(中文编号|中文章节|一、|（一）)", teacher_joined) and "1.1.1" not in teacher_joined:
        config.numbering_scheme = "chinese"
        add_source(config, "numbering_scheme", "教师要求文本")
        return
    if re.search(r"1\.1(?:\.1)?", teacher_joined):
        config.numbering_scheme = "numeric"
        add_source(config, "numbering_scheme", "教师要求文本")
        return

    numeric_count = 0
    chinese_count = 0
    for text in paper_texts:
        if re.match(r"^\d+\.\d+\.\d+(?:\s|　|$)", text) or re.match(r"^\d+\.\d+(?:\s|　|$)", text) or re.match(r"^\d+(?:\.|．|、)(?:\s|　)+", text):
            numeric_count += 1
        if re.match(r"^[一二三四五六七八九十]+、", text) or re.match(r"^（[一二三四五六七八九十]+）", text) or re.match(r"^第[一二三四五六七八九十\d]+章", text):
            chinese_count += 1

    if chinese_count > numeric_count and chinese_count >= 2:
        config.numbering_scheme = "chinese"
    else:
        config.numbering_scheme = "numeric"
    add_source(config, "numbering_scheme", "论文现有主导编号")


def finalize_heading_styles(config):
    heading_font = config.title_style.east_asia or "黑体"
    base = replace(
        deepcopy(config.body_style),
        east_asia=heading_font,
        bold=True,
        first_line_chars=None,
    )
    defaults = {
        1: replace(deepcopy(base), outline_level=0),
        2: replace(deepcopy(base), outline_level=1),
        3: replace(deepcopy(base), outline_level=2),
    }
    for level in [1, 2, 3]:
        key = f"heading{level}_style"
        resolved = resolve_style(defaults[level], getattr(config, key))
        setattr(config, key, resolved)


def extract_teacher_config(teacher_docx, paper_texts):
    config = TeacherConfig()
    _, teacher_document = parse_docx_root(teacher_docx)
    teacher_paragraphs = paragraphs_from_body(teacher_document)
    teacher_tables = extract_table_rows(teacher_document)
    teacher_texts = [para_text(p) for p in teacher_paragraphs if para_text(p)]

    apply_teacher_example_page(teacher_document, config)
    apply_teacher_example_styles(teacher_paragraphs, config)
    apply_teacher_heading_examples(teacher_paragraphs, config)
    apply_body_rules_from_text(teacher_texts, config)
    apply_heading_rules_from_text(teacher_texts, config)
    apply_heading_rules_from_tables(teacher_tables, config)
    infer_numbering_scheme(teacher_texts, paper_texts, config)
    finalize_heading_styles(config)
    return config

def style_from_dict(data):
    return StyleSpec(**data)


def page_from_dict(data):
    return PageSpec(**data)


def config_from_dict(data):
    config = TeacherConfig()
    for key in [
        "title_style",
        "body_style",
        "heading1_style",
        "heading2_style",
        "heading3_style",
        "abstract_style",
        "keywords_style",
        "reference_title_style",
        "reference_body_style",
    ]:
        if key in data:
            setattr(config, key, style_from_dict(data[key]))
    if "page" in data:
        config.page = page_from_dict(data["page"])
    if "numbering_scheme" in data:
        config.numbering_scheme = data["numbering_scheme"]
    if "min_chars" in data:
        config.min_chars = data["min_chars"]
    if "sources" in data:
        config.sources.update(data["sources"])
    if "warnings" in data:
        config.warnings = list(data["warnings"])
    finalize_heading_styles(config)
    return config


def get_config_schema():
    return {
        "schema_name": "docx-paper-formatter-config",
        "schema_version": CONFIG_SCHEMA_VERSION,
        "required_keys": [
            "config_version",
            "title_style",
            "body_style",
            "heading1_style",
            "heading2_style",
            "heading3_style",
            "page",
            "numbering_scheme",
            "min_chars",
            "sources",
        ],
        "style_fields": {
            "east_asia": "string|null",
            "ascii_font": "string|null",
            "size_half_points": "digit-string|null",
            "bold": "boolean|null",
            "jc": f"enum({sorted(ALLOWED_ALIGNMENTS)})|null",
            "first_line_chars": "digit-string|null",
            "outline_level": "integer|null",
            "line_rule": f"enum({sorted(ALLOWED_LINE_RULES)})|null",
            "line": "digit-string|null",
            "before": "digit-string|null",
            "after": "digit-string|null",
        },
        "page_fields": {
            "width": "digit-string",
            "height": "digit-string",
            "top": "digit-string",
            "right": "digit-string",
            "bottom": "digit-string",
            "left": "digit-string",
            "header": "digit-string",
            "footer": "digit-string",
            "gutter": "digit-string",
        },
        "numbering_scheme": sorted(ALLOWED_NUMBERING_SCHEMES),
    }


def ensure_string_or_none(value, path):
    if value is None or isinstance(value, str):
        return
    raise ValueError(f"{path} 必须是字符串或 null")


def ensure_digit_string_or_none(value, path):
    if value is None:
        return
    if isinstance(value, str) and value.isdigit():
        return
    raise ValueError(f"{path} 必须是数字字符串或 null")


def validate_style_dict(data, path):
    if not isinstance(data, dict):
        raise ValueError(f"{path} 必须是对象")
    allowed_fields = set(StyleSpec.__dataclass_fields__.keys())
    extra = set(data.keys()) - allowed_fields
    if extra:
        raise ValueError(f"{path} 包含未知字段: {sorted(extra)}")
    ensure_string_or_none(data.get("east_asia"), f"{path}.east_asia")
    ensure_string_or_none(data.get("ascii_font"), f"{path}.ascii_font")
    ensure_digit_string_or_none(data.get("size_half_points"), f"{path}.size_half_points")
    if data.get("bold") is not None and not isinstance(data.get("bold"), bool):
        raise ValueError(f"{path}.bold 必须是布尔值或 null")
    if data.get("jc") is not None and data.get("jc") not in ALLOWED_ALIGNMENTS:
        raise ValueError(f"{path}.jc 必须是 {sorted(ALLOWED_ALIGNMENTS)} 之一")
    ensure_digit_string_or_none(data.get("first_line_chars"), f"{path}.first_line_chars")
    if data.get("outline_level") is not None and not isinstance(data.get("outline_level"), int):
        raise ValueError(f"{path}.outline_level 必须是整数或 null")
    if data.get("line_rule") is not None and data.get("line_rule") not in ALLOWED_LINE_RULES:
        raise ValueError(f"{path}.line_rule 必须是 {sorted(ALLOWED_LINE_RULES)} 之一")
    ensure_digit_string_or_none(data.get("line"), f"{path}.line")
    ensure_digit_string_or_none(data.get("before"), f"{path}.before")
    ensure_digit_string_or_none(data.get("after"), f"{path}.after")


def validate_page_dict(data, path):
    if not isinstance(data, dict):
        raise ValueError(f"{path} 必须是对象")
    required = set(PageSpec.__dataclass_fields__.keys())
    missing = required - set(data.keys())
    if missing:
        raise ValueError(f"{path} 缺少字段: {sorted(missing)}")
    extra = set(data.keys()) - required
    if extra:
        raise ValueError(f"{path} 包含未知字段: {sorted(extra)}")
    for key in required:
        ensure_digit_string_or_none(data.get(key), f"{path}.{key}")
        if data.get(key) is None:
            raise ValueError(f"{path}.{key} 不能为空")


def validate_config_dict(data):
    if not isinstance(data, dict):
        raise ValueError("配置文件根节点必须是对象")
    schema = get_config_schema()
    missing = [key for key in schema["required_keys"] if key not in data]
    if missing:
        raise ValueError(f"配置文件缺少字段: {missing}")
    if data.get("config_version") != CONFIG_SCHEMA_VERSION:
        raise ValueError(f"config_version 必须为 {CONFIG_SCHEMA_VERSION}")
    for key in [
        "title_style",
        "body_style",
        "heading1_style",
        "heading2_style",
        "heading3_style",
        "abstract_style",
        "keywords_style",
        "reference_title_style",
        "reference_body_style",
    ]:
        if key in data:
            validate_style_dict(data[key], key)
    validate_page_dict(data["page"], "page")
    if data["numbering_scheme"] not in ALLOWED_NUMBERING_SCHEMES:
        raise ValueError(f"numbering_scheme 必须是 {sorted(ALLOWED_NUMBERING_SCHEMES)} 之一")
    if not isinstance(data["min_chars"], int) or data["min_chars"] <= 0:
        raise ValueError("min_chars 必须是正整数")
    if not isinstance(data["sources"], dict):
        raise ValueError("sources 必须是对象")
    for key, value in data["sources"].items():
        if not isinstance(key, str) or not isinstance(value, str):
            raise ValueError("sources 的键和值都必须是字符串")
    if "warnings" in data and (
        not isinstance(data["warnings"], list) or any(not isinstance(item, str) for item in data["warnings"])
    ):
        raise ValueError("warnings 必须是字符串数组")


def config_to_dict(config):
    data = asdict(config)
    data["config_version"] = CONFIG_SCHEMA_VERSION
    return data


def save_config_json(config, path):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    data = config_to_dict(config)
    validate_config_dict(data)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def save_schema_json(path):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(get_config_schema(), ensure_ascii=False, indent=2), encoding="utf-8")


def load_config_json(path):
    path = Path(path)
    data = json.loads(path.read_text(encoding="utf-8"))
    validate_config_dict(data)
    return config_from_dict(data)


def numbering_scheme_label(numbering_scheme):
    return "中文章节编号" if numbering_scheme == "chinese" else "阿拉伯数字编号"


def numbering_pattern_label(numbering_scheme):
    if numbering_scheme == "chinese":
        return "一级 `一、`，二级 `（一）`，三级 `1.`"
    return "一级 `1`，二级 `1.1`，三级 `1.1.1`"


def describe_style(spec):
    font = f"中文 {spec.east_asia} / 英文 {spec.ascii_font}"
    size_pt = round(int(spec.size_half_points) / 2, 1) if spec.size_half_points else "未知"
    if spec.line_rule == "exact" and spec.line is not None:
        line_text = f"固定值 {round(int(spec.line) / 20, 1)} 磅"
    elif spec.line is not None:
        line_text = f"{round(int(spec.line) / 240, 2)} 倍行距"
    else:
        line_text = "未指定"
    indent = f"{round(int(spec.first_line_chars) / 100, 2)} 字符" if spec.first_line_chars else "无"
    bold = "加粗" if spec.bold else "常规"
    align = spec.jc or "默认"
    before = f"{round(int(spec.before) / 20, 1)} 磅" if spec.before is not None else "0"
    after = f"{round(int(spec.after) / 20, 1)} 磅" if spec.after is not None else "0"
    return f"{font}，{size_pt} pt，{bold}，对齐={align}，行距={line_text}，段前={before}，段后={after}，首行缩进={indent}"


def humanize_font_name(name):
    mapping = {
        "???": "HeiTi",
        "???": "SongTi",
    }
    if name is None:
        return ""
    return mapping.get(name, name)


def build_report(config, num_id, abstract_id, reference_num_id, numbering_check, classified, config_origin=None):
    headings1 = [(orig, clean) for orig, clean, style_id, _, _ in classified if style_id == "PaperHeading1"]
    headings2 = [(orig, clean) for orig, clean, style_id, _, _ in classified if style_id == "PaperHeading2"]
    headings3 = [(orig, clean) for orig, clean, style_id, _, _ in classified if style_id == "PaperHeading3"]
    refs = [orig for orig, _, style_id, _, _ in classified if style_id == "ReferenceBody"]
    missing = []
    if not any(style_id == "AbstractStyle" for _, _, style_id, _, _ in classified):
        missing.append("Missing abstract")
    if not any(style_id == "KeywordsStyle" for _, _, style_id, _, _ in classified):
        missing.append("Missing keywords")
    if not any(style_id == "ReferenceTitle" for _, _, style_id, _, _ in classified):
        missing.append("Missing references section")

    approx_chars = count_chars(
        [clean for _, clean, style_id, _, _ in classified if style_id in {"PaperBody", "PaperHeading1", "PaperHeading2", "PaperHeading3"}]
    )
    if approx_chars < config.min_chars:
        missing.append(f"Body length is about {approx_chars}, below minimum {config.min_chars}")

    report = [
        "# Formatting Check Report",
        "",
        "This report describes the extracted rules, the Word style bindings, and the automatic numbering result for headings and references.",
        "",
        "## Rules",
        "",
        "| Item | Value | Source |",
        "|---|---|---|",
        f"| Title style | {describe_style(config.title_style)} | {config.sources['title_style']} |",
        f"| Body style | {describe_style(config.body_style)} | {config.sources['body_style']} |",
        f"| Heading 1 style | {describe_style(config.heading1_style)} | {config.sources['heading1_style']} |",
        f"| Heading 2 style | {describe_style(config.heading2_style)} | {config.sources['heading2_style']} |",
        f"| Heading 3 style | {describe_style(config.heading3_style)} | {config.sources['heading3_style']} |",
        f"| Reference title style | {describe_style(config.reference_title_style)} | {config.sources.get('reference_title_style', 'derived default')} |",
        f"| Reference body style | {describe_style(config.reference_body_style)} | {config.sources.get('reference_body_style', 'derived default')} |",
        *([f"| Config replay | {config_origin} | external config |"] if config_origin else []),
        f"| Page | A4, top {twips_to_cm(config.page.top)} / bottom {twips_to_cm(config.page.bottom)} / left {twips_to_cm(config.page.left)} / right {twips_to_cm(config.page.right)} cm | {config.sources['page']} |",
        f"| Numbering scheme | {numbering_scheme_label(config.numbering_scheme)} | {config.sources['numbering_scheme']} |",
        f"| Minimum chars | {config.min_chars} | {config.sources['min_chars']} |",
        "",
        "## Style Mapping",
        "",
        "| Content | Style ID | Numbering |",
        "|---|---|---|",
        "| Paper title | PaperTitle | none |",
        f"| Heading 1 | PaperHeading1 | numId={num_id}, ilvl=0 |",
        f"| Heading 2 | PaperHeading2 | numId={num_id}, ilvl=1 |",
        f"| Heading 3 | PaperHeading3 | numId={num_id}, ilvl=2 |",
        "| Body | PaperBody | none |",
        "| Abstract | AbstractStyle | none |",
        "| Keywords | KeywordsStyle | none |",
        "| Reference title | ReferenceTitle | none |",
        f"| Reference body | ReferenceBody | numId={reference_num_id}, ilvl=0 |",
        "",
        "## Numbering Check",
        "",
        f"- Heading multilevel numbering created: abstractNumId={abstract_id}, numId={num_id}.",
        f"- Heading numbering pattern: {numbering_pattern_label(config.numbering_scheme)}.",
        f"- Recognized heading paragraphs: {numbering_check['heading_count']}.",
        f"- Unified heading numId: {numbering_check['continuous_num_id']}.",
        f"- Paragraph-level fixes: {numbering_check['paragraph_fixes']}; style-level fixes: {numbering_check['style_fixes']}.",
        f"- Remaining bad heading numId count: {numbering_check['bad_heading_num_id_after']}.",
        f"- Reference paragraphs using automatic numbering: {len(refs)}, numId={reference_num_id}.",
        "",
        "## Detected Headings",
        "",
        "### Heading 1",
        *( ["- none"] if not headings1 else [f"- {orig} -> {clean}" for orig, clean in headings1] ),
        "",
        "### Heading 2",
        *( ["- none"] if not headings2 else [f"- {orig} -> {clean}" for orig, clean in headings2] ),
        "",
        "### Heading 3",
        *( ["- none"] if not headings3 else [f"- {orig} -> {clean}" for orig, clean in headings3] ),
        "",
        "## Risks",
        "",
        *( ["- No obvious missing abstract/keywords/references issues detected."] if not missing else [f"- {x}" for x in missing] ),
        *( [f"- {warning}" for warning in config.warnings] ),
        "",
        "## Key Decisions",
        "",
        "- Cover handling has been fully removed from this skill.",
        "- Only body content, headings, abstract, keywords, references, and page layout are normalized.",
        "- Headings use a reusable multilevel Word list.",
        "- References use a separate automatic numbering list.",
    ]
    return "\n".join(report)


def process(
    teacher_docx=DEFAULT_TEACHER_DOCX,
    paper_docx=DEFAULT_PAPER_DOCX,
    output_docx=DEFAULT_OUTPUT_DOCX,
    report_md=DEFAULT_REPORT_MD,
    config_json=None,
    config_out=None,
    config_only=False,
    schema_out=None,
):
    teacher_docx = Path(teacher_docx)
    paper_docx = Path(paper_docx)
    output_docx = Path(output_docx)
    report_md = Path(report_md)
    if schema_out:
        save_schema_json(schema_out)

    if not paper_docx.exists():
        raise FileNotFoundError("缺少待修改论文文档")
    if config_json is None and not teacher_docx.exists():
        raise FileNotFoundError("缺少教师格式要求文档")

    parts, document = parse_docx_root(paper_docx)
    styles = etree.fromstring(parts["word/styles.xml"])
    body = document.find("w:body", NS)

    original_paras = [deepcopy(el) for el in body if el.tag == w("p")]
    non_empty = [p for p in original_paras if para_text(p)]
    if not non_empty:
        raise ValueError("待修改论文文档未识别到正文段落")

    paper_texts = [para_text(p) for p in non_empty]
    config_origin = None
    if config_json:
        config = load_config_json(config_json)
        config_origin = Path(config_json).name
    else:
        config = extract_teacher_config(teacher_docx, paper_texts)

    if config_out:
        save_config_json(config, config_out)
    if config_only:
        return

    num_id, abstract_id = ensure_numbering(parts, config.numbering_scheme)
    reference_num_id, _ = ensure_reference_numbering(parts)
    ensure_styles(styles, num_id, config)
    bind_reference_style_numbering(styles, reference_num_id)

    final_sect = body.find("w:sectPr", NS)
    final_sect = deepcopy(final_sect) if final_sect is not None else make_section(config.page)
    set_section_page(final_sect, config.page)

    add_footer(parts, final_sect)

    classified = []
    in_refs = False
    previous_heading_ilvl = None
    for idx, p in enumerate(non_empty):
        text = para_text(p)
        style_id, label, ilvl, clean_text = classify(text, idx, in_refs, previous_heading_ilvl)
        if label == "参考文献标题":
            in_refs = True
        if style_id == "ReferenceBody":
            clean_text = strip_reference_marker(clean_text)
        if clean_text != text:
            set_para_text(p, clean_text)
        if style_id == "ReferenceBody":
            apply_style_and_numbering(p, style_id, 0, reference_num_id)
        elif ilvl is None:
            apply_style(p, style_id)
        else:
            apply_style_and_numbering(p, style_id, ilvl, num_id)
            previous_heading_ilvl = ilvl
        classified.append((text, clean_text, style_id, label, ilvl))

    for el in list(body):
        body.remove(el)
    for p in non_empty:
        body.append(p)
    body.append(final_sect)

    numbering_check = enforce_heading_numbering_continuity(
        document,
        styles,
        parts,
        num_id,
        abstract_id,
        config.numbering_scheme,
    )

    parts["word/document.xml"] = etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)
    parts["word/styles.xml"] = etree.tostring(styles, xml_declaration=True, encoding="UTF-8", standalone=True)

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    report_md.parent.mkdir(parents=True, exist_ok=True)

    with ZipFile(output_docx, "w", ZIP_DEFLATED) as zout:
        for name, data in parts.items():
            zout.writestr(name, data)

    report_text = build_report(config, num_id, abstract_id, reference_num_id, numbering_check, classified, config_origin)
    report_md.write_text(report_text, encoding="utf-8")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Normalize an academic paper DOCX by extracting teacher rules and applying reusable Word heading numbering."
    )
    parser.add_argument("--teacher", type=Path, default=DEFAULT_TEACHER_DOCX)
    parser.add_argument("--paper", type=Path, default=DEFAULT_PAPER_DOCX)
    parser.add_argument("--out", type=Path, default=DEFAULT_OUTPUT_DOCX)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT_MD)
    parser.add_argument("--config", type=Path, default=None, help="Load extracted formatting config from JSON.")
    parser.add_argument("--config-out", type=Path, default=None, help="Write extracted formatting config to JSON.")
    parser.add_argument("--config-only", action="store_true", help="Only extract/save config, do not write output docx.")
    parser.add_argument("--schema-out", type=Path, default=None, help="Write stable config schema description to JSON.")
    args = parser.parse_args()
    process(args.teacher, args.paper, args.out, args.report, args.config, args.config_out, args.config_only, args.schema_out)
