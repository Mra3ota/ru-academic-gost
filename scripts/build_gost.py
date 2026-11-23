#!/usr/bin/env python3
"""Assemble GOST-compliant Word (.docx) from Markdown drafts via Pandoc + in-place python-docx + Word TOC update.

v2.0 features:
- Symbolic citation keys [@key, с. X] auto-resolved against GOST-sorted bibliography (НПА -> А-Я/Ё=Е -> A-Z)
- Native Word OMML numbered equations: $$...$$ {#eq:1} or $$...$$ (1) in borderless 3-column GOST layout
- Custom university title page template import (--title-template custom_title.docx) with image & {{placeholder}} support
- Automatic Pandoc --resource-path for embedded GOST charts/figures (data/fig_N.png)
- Native Word footnotes styled directly in word/footnotes.xml with strict superscript marks
- Live Microsoft Word COM TOC update with language-independent WdBuiltinStyle (-20/-21/-22)
"""

from __future__ import annotations

import argparse
import copy
import io
import json
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Mm, Pt, RGBColor
from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _common import (
    clean_hidden_unicode,
    ensure_utf8_io,
    find_pandoc,
    load_contract,
    ru_sort_key,
    update_word_toc_and_export_pdf,
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

EQ_MARKER_PREFIX = "@@GOST_EQ_NUM:"
EQ_TBL_DESC = "GOST_EQUATION_LAYOUT"

# Symbolic citation bracket pattern: [@key], [@key, с. 45], [@k1, с. 10; @k2, с. 12]
SYMBOLIC_CITE_GROUP_RE = re.compile(
    r"\[(@[A-Za-zА-Яа-яЁё0-9_\-]+(?:\s*,\s*[^;\]]+)?(?:\s*;\s*@[A-Za-zА-Яа-яЁё0-9_\-]+(?:\s*,\s*[^;\]]+)?)*)\]"
)
SYMBOLIC_BIB_LINE_RE = re.compile(r"^\s*\[@([A-Za-zА-Яа-яЁё0-9_\-]+)\]\s+(.+)$")


def set_run_font(
    run: Any,
    font_name: str,
    size_pt: float,
    bold: bool | None = None,
    italic: bool | None = None,
) -> None:
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.color.rgb = RGBColor(0, 0, 0)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        rFonts.set(qn(f"w:{attr}"), font_name)


def configure_style(
    style: Any,
    font_name: str,
    size_pt: float,
    line_spacing: float = 1.5,
    first_line_cm: float = 1.25,
    align: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.JUSTIFY,
    bold: bool = False,
    space_before_pt: float = 0.0,
    space_after_pt: float = 0.0,
    keep_with_next: bool = False,
    page_break_before: bool = False,
) -> None:
    style.font.name = font_name
    style.font.size = Pt(size_pt)
    style.font.bold = bold
    style.font.color.rgb = RGBColor(0, 0, 0)
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        rFonts.set(qn(f"w:{attr}"), font_name)

    pf = style.paragraph_format
    pf.alignment = align
    pf.line_spacing = line_spacing
    pf.first_line_indent = Cm(first_line_cm)
    pf.left_indent = Cm(0)
    pf.right_indent = Cm(0)
    pf.space_before = Pt(space_before_pt)
    pf.space_after = Pt(space_after_pt)
    pf.keep_with_next = keep_with_next
    pf.page_break_before = page_break_before
    pf.widow_control = True


def ensure_toc_and_footnote_styles(doc: Document, contract: dict[str, Any]) -> None:
    layout = contract["layout"]
    font_name = layout["font_name"]
    font_size = float(layout["font_size_pt"])

    for level, left_cm in ((1, 0.0), (2, 0.6), (3, 1.2)):
        sname = f"TOC {level}"
        st = doc.styles[sname] if sname in doc.styles else doc.styles.add_style(sname, WD_STYLE_TYPE.PARAGRAPH)
        configure_style(
            st,
            font_name,
            font_size,
            line_spacing=1.15,
            first_line_cm=0.0,
            align=WD_ALIGN_PARAGRAPH.LEFT,
            bold=False,
            space_before_pt=2.0,
            space_after_pt=2.0,
        )
        st.paragraph_format.left_indent = Cm(left_cm)

    for ref_sname in ("Footnote Reference", "FootnoteReference"):
        rst = doc.styles[ref_sname] if ref_sname in doc.styles else doc.styles.add_style(ref_sname, WD_STYLE_TYPE.CHARACTER)
        rst.font.name = font_name
        rst.font.color.rgb = RGBColor(0, 0, 0)
        rPr = rst.element.get_or_add_rPr()
        if rPr.find(qn("w:vertAlign")) is None:
            va = OxmlElement("w:vertAlign")
            va.set(qn("w:val"), "superscript")
            rPr.append(va)


def build_reference_docx(ref_path: Path, contract: dict[str, Any]) -> None:
    layout = contract["layout"]
    font_name = layout["font_name"]
    font_size = float(layout["font_size_pt"])
    line_sp = float(layout["line_spacing"])
    indent_cm = float(layout["first_line_indent_cm"])
    fn_size = float(layout["footnote_font_size_pt"])
    margins = layout["margins_mm"]

    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Mm(210)
    sec.page_height = Mm(297)
    sec.left_margin = Mm(float(margins["left"]))
    sec.right_margin = Mm(float(margins["right"]))
    sec.top_margin = Mm(float(margins["top"]))
    sec.bottom_margin = Mm(float(margins["bottom"]))

    configure_style(doc.styles["Normal"], font_name, font_size, line_sp, indent_cm)
    for sname in ("Body Text", "First Paragraph", "Compact", "List Paragraph"):
        if sname in doc.styles:
            configure_style(doc.styles[sname], font_name, font_size, line_sp, indent_cm)

    configure_style(
        doc.styles["Heading 1"],
        font_name,
        font_size,
        line_spacing=1.5,
        first_line_cm=0.0,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        bold=True,
        space_before_pt=0.0,
        space_after_pt=12.0,
        keep_with_next=True,
        page_break_before=True,
    )
    configure_style(
        doc.styles["Heading 2"],
        font_name,
        font_size,
        line_spacing=1.5,
        first_line_cm=indent_cm,
        align=WD_ALIGN_PARAGRAPH.LEFT,
        bold=True,
        space_before_pt=12.0,
        space_after_pt=6.0,
        keep_with_next=True,
    )
    configure_style(
        doc.styles["Heading 3"],
        font_name,
        font_size,
        line_spacing=1.5,
        first_line_cm=indent_cm,
        align=WD_ALIGN_PARAGRAPH.LEFT,
        bold=True,
        space_before_pt=8.0,
        space_after_pt=4.0,
        keep_with_next=True,
    )
    for fn_style in ("Footnote Text", "FootnoteText"):
        if fn_style in doc.styles:
            configure_style(
                doc.styles[fn_style],
                font_name,
                fn_size,
                line_spacing=1.0,
                first_line_cm=0.0,
                align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                bold=False,
            )

    ensure_toc_and_footnote_styles(doc, contract)
    doc.save(str(ref_path))


def collect_draft_files(drafts_dir: Path) -> list[Path]:
    return sorted(
        [
            p
            for p in drafts_dir.iterdir()
            if p.is_file()
            and p.suffix.lower() == ".md"
            and not p.name.startswith(("00_", "sources_"))
            and not p.name.endswith(".before.md")
        ],
        key=lambda p: p.name.lower(),
    )


def is_biblio_file(fpath: Path, text: str) -> bool:
    return "biblio" in fpath.name.lower() or bool(
        re.search(r"^#\s*список\s+(?:использованн\w+\s+)?(?:источников|литературы)", text, re.I | re.M)
    )


def parse_and_sort_symbolic_biblio(text: str, npa_first: bool) -> tuple[str, dict[str, int]]:
    """If bibliography uses [@key] entries, sort them strictly by GOST and return (rewritten_md, key_to_num)."""
    lines = text.splitlines()
    entries: list[tuple[str, str]] = []
    header_lines: list[str] = []
    seen_keys: set[str] = set()

    for ln in lines:
        m = SYMBOLIC_BIB_LINE_RE.match(ln)
        if m:
            key, body = m.group(1).strip(), m.group(2).strip()
            if key in seen_keys:
                raise ValueError(f"Duplicate bibliography key [@{key}]")
            seen_keys.add(key)
            entries.append((key, body))
        elif not entries:
            header_lines.append(ln)

    if not entries:
        return text, {}

    sorted_entries = sorted(entries, key=lambda pair: ru_sort_key(pair[1], npa_first=npa_first))
    key_to_num: dict[str, int] = {}
    out_lines = list(header_lines)
    if out_lines and out_lines[-1].strip():
        out_lines.append("")
    for idx, (key, body) in enumerate(sorted_entries, start=1):
        key_to_num[key] = idx
        out_lines.append(f"{idx}\\. {body}")
        out_lines.append("")

    return "\n".join(out_lines), key_to_num


def resolve_symbolic_citations(
    text: str,
    key_to_num: dict[str, int],
    used_keys: set[str],
    filename: str,
) -> str:
    """Replace [@key, с. X] and [@k1, с. 5; @k2] with exact GOST-sorted [N, с. X] numbers."""
    if not key_to_num:
        # Still catch dangling [@key] if someone used symbolic citations without tagged bibliography
        dangling = SYMBOLIC_CITE_GROUP_RE.findall(text)
        if dangling:
            raise ValueError(
                f"Symbolic citations {dangling[:3]} found in {filename}, "
                "but bibliography does not define [@key] entries."
            )
        return text

    def _replace_group(m: re.Match[str]) -> str:
        raw_inner = m.group(1)
        items = [part.strip() for part in raw_inner.split(";")]
        resolved_parts: list[str] = []
        for item in items:
            m_it = re.match(r"^@([A-Za-zА-Яа-яЁё0-9_\-]+)(.*)$", item)
            if not m_it:
                resolved_parts.append(item)
                continue
            k = m_it.group(1)
            suffix = m_it.group(2)
            if k not in key_to_num:
                raise ValueError(f"Undefined symbolic citation key [@{k}] in {filename}")
            used_keys.add(k)
            resolved_parts.append(f"{key_to_num[k]}{suffix}")
        return "[" + "; ".join(resolved_parts) + "]"

    return SYMBOLIC_CITE_GROUP_RE.sub(_replace_group, text)


def extract_numbered_equations(text: str) -> tuple[str, int]:
    """Convert display equations tagged with {#eq:N} or (N) into pure Pandoc $$...$$ + marker paragraph."""
    eq_count = 0

    # Match multi-line or single-line $$...$$ followed by {#eq:1.1} or (1.1) on the closing line
    eq_re = re.compile(
        r"(\$\$\s*[\s\S]*?\$\$)[ \t]*(?:\{#eq:([A-Za-z0-9._\-]+)\}|\((\d+(?:\.\d+)?)\))[ \t]*(?=\n|$)"
    )

    def _repl(m: re.Match[str]) -> str:
        nonlocal eq_count
        eq_count += 1
        math_blk = m.group(1).strip()
        num = (m.group(2) or m.group(3)).strip()
        return f"\n\n{math_blk}\n\n{EQ_MARKER_PREFIX}{num}@@\n\n"

    return eq_re.sub(_repl, text), eq_count


def normalize_markdown_for_pandoc(
    files: list[Path],
    contract: dict[str, Any],
    strict_bib: bool = False,
) -> tuple[str, list[str], dict[str, int], int]:
    """Sanitize drafts, resolve symbolic [@key] citations, convert numbered equations, and namespace footnotes."""
    warnings: list[str] = []
    npa_first = bool(contract.get("bibliography", {}).get("npa_first", True))

    raw_texts: list[tuple[Path, str]] = []
    key_to_num: dict[str, int] = {}
    bib_rewritten: dict[Path, str] = {}

    for fpath in files:
        txt = clean_hidden_unicode(fpath.read_text(encoding="utf-8"))
        txt = txt.replace("\r\n", "\n").replace("\r", "\n")
        raw_texts.append((fpath, txt))
        if is_biblio_file(fpath, txt):
            rewritten, mapping = parse_and_sort_symbolic_biblio(txt, npa_first=npa_first)
            if mapping:
                key_to_num = mapping
                bib_rewritten[fpath] = rewritten

    used_keys: set[str] = set()
    total_equations = 0
    combined_parts: list[str] = []

    for fidx, (fpath, text) in enumerate(raw_texts, start=1):
        is_bib = is_biblio_file(fpath, text)
        if fpath in bib_rewritten:
            text = bib_rewritten[fpath]
        else:
            text = resolve_symbolic_citations(text, key_to_num, used_keys, fpath.name)

        # Validate internal footnote definitions within this file
        def_ids = re.findall(r"^\[\^([^\]]+)\]:", text, re.M)
        if len(def_ids) != len(set(def_ids)):
            dups = sorted({k for k in def_ids if def_ids.count(k) > 1})
            raise ValueError(f"Duplicate footnote definitions inside {fpath.name}: {dups}")

        ref_ids = set(re.findall(r"(?<!^)\[\^([^\]]+)\](?!:)", text, re.M))
        missing = sorted(ref_ids - set(def_ids))
        if missing:
            raise ValueError(f"Undefined footnote reference(s) in {fpath.name}: {missing}")

        # Auto-namespace footnote IDs per file to eliminate cross-file key collisions
        prefix = f"f{fidx:02d}_"
        text = re.sub(r"\[\^([^\]]+)\]", lambda m: f"[^{prefix}{m.group(1)}]", text)

        # Normalize footnote punctuation spacing (GOST: «слово[^id].»)
        text = re.sub(r"(?<=[^\s\n])[ \t]+(\[\^[^\]]+\])(?!:)", r"\1", text)
        text = re.sub(r"([,;:])(\[\^[^\]]+\])(?!:)", r"\2\1", text)
        text = re.sub(
            r"(?<!\bруб)(?<!\bтыс)(?<!\bмлн)(?<!\bмлрд)(?<!\bст)(?<!\bдр)(?<!\bсм)(?<!\bпр)(?<!\bгг)(?<!\bвв)(?<!\s[гвс])(?<![а-яё]\.[а-яё])\.(\[\^[^\]]+\])(?!:|\.)",
            r"\1.",
            text,
        )
        text = re.sub(r"(\[\^[^\]]+\])[ \t]+([.,;:])", r"\1\2", text)

        # Numbered display equations -> pure Pandoc OMML + marker paragraph
        text, eq_n = extract_numbered_equations(text)
        total_equations += eq_n

        # Strip trailing periods from Markdown headings (# / ## / ###) per GOST 7.32
        def _clean_heading(m: re.Match[str]) -> str:
            hashes, title = m.group(1), m.group(2).strip()
            title = re.sub(r"\.$", "", title)
            return f"{hashes} {title}"

        text = re.sub(r"^(#{1,3})\s+(.+)$", _clean_heading, text, flags=re.M)

        # Plain numbered bibliography (when NOT using [@key] symbolic entries) -> separate GOST paragraphs
        if is_bib and fpath not in bib_rewritten:
            bib_lines: list[str] = []
            for ln in text.splitlines():
                m_item = re.match(r"^(\d+)\.\s+(.*)$", ln.strip())
                if m_item:
                    if bib_lines and bib_lines[-1] != "":
                        bib_lines.append("")
                    bib_lines.append(f"{m_item.group(1)}\\. {m_item.group(2)}")
                    bib_lines.append("")
                else:
                    bib_lines.append(ln)
            text = "\n".join(bib_lines)

        combined_parts.append(text.strip())

    if key_to_num:
        unused = sorted(set(key_to_num.keys()) - used_keys)
        if unused:
            msg = f"Uncited bibliography keys in [@key] list: {unused}"
            if strict_bib:
                raise ValueError(msg)
            warnings.append(msg)

    return "\n\n".join(combined_parts) + "\n", warnings, key_to_num, total_equations


def make_xml_paragraph(
    text: str,
    font_name: str,
    size_pt: float,
    align: str = "center",
    bold: bool = False,
    italic: bool = False,
    space_before_pt: float = 0.0,
    space_after_pt: float = 0.0,
    line_spacing: float = 1.0,
    left_indent_cm: float = 0.0,
    keep_with_next: bool = False,
    page_break_before: bool = False,
) -> etree._Element:
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    if page_break_before:
        pPr.append(OxmlElement("w:pageBreakBefore"))
    if keep_with_next:
        pPr.append(OxmlElement("w:keepNext"))

    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), align)
    pPr.append(jc)

    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), str(int(round(space_before_pt * 20))))
    spacing.set(qn("w:after"), str(int(round(space_after_pt * 20))))
    spacing.set(qn("w:line"), str(int(round(line_spacing * 240))))
    spacing.set(qn("w:lineRule"), "auto")
    pPr.append(spacing)

    ind = OxmlElement("w:ind")
    ind.set(qn("w:firstLine"), "0")
    ind.set(qn("w:left"), str(int(round(left_indent_cm * 567))))
    ind.set(qn("w:right"), "0")
    pPr.append(ind)
    p.append(pPr)

    if text:
        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        rFonts = OxmlElement("w:rFonts")
        for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
            rFonts.set(qn(f"w:{attr}"), font_name)
        rPr.append(rFonts)
        if bold:
            rPr.append(OxmlElement("w:b"))
        if italic:
            rPr.append(OxmlElement("w:i"))
        col = OxmlElement("w:color")
        col.set(qn("w:val"), "000000")
        rPr.append(col)
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(int(round(size_pt * 2))))
        rPr.append(sz)
        szCs = OxmlElement("w:szCs")
        szCs.set(qn("w:val"), str(int(round(size_pt * 2))))
        rPr.append(szCs)
        r.append(rPr)

        t = OxmlElement("w:t")
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = text
        r.append(t)
        p.append(r)

    return p


def build_title_page_elements(meta: dict[str, Any], contract: dict[str, Any]) -> list[etree._Element]:
    font = contract["layout"]["font_name"]
    work_type = contract.get("work_type", "kursovaya")
    elems: list[etree._Element] = []

    def add_p(txt: str, **kw: Any) -> None:
        elems.append(make_xml_paragraph(txt, font_name=font, **kw))

    if meta.get("ministry"):
        add_p(str(meta["ministry"]).upper(), size_pt=12, align="center", bold=False, space_after_pt=4, line_spacing=1.0)
    if meta.get("university"):
        add_p(str(meta["university"]), size_pt=12, align="center", bold=True, space_after_pt=4, line_spacing=1.05)
    if meta.get("institute"):
        add_p(str(meta["institute"]), size_pt=12, align="center", bold=False, space_after_pt=3, line_spacing=1.05)
    if meta.get("department"):
        add_p(str(meta["department"]), size_pt=12, align="center", bold=False, space_after_pt=6, line_spacing=1.05)

    work_label = str(meta.get("work_label") or (
        "ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА" if work_type == "vkr"
        else ("ДОКЛАД" if work_type == "report" else ("РЕФЕРАТ" if work_type == "referat" else "КУРСОВАЯ РАБОТА"))
    )).upper()
    add_p(work_label, size_pt=16, align="center", bold=True, space_before_pt=72, space_after_pt=8, line_spacing=1.15)

    if meta.get("discipline"):
        add_p(str(meta["discipline"]), size_pt=14, align="center", bold=False, space_after_pt=10, line_spacing=1.15)

    topic = str(meta.get("topic") or contract.get("topic") or "").rstrip(".")
    if topic:
        add_p(f"на тему: «{topic}»", size_pt=14, align="center", bold=True, space_after_pt=10, line_spacing=1.2)

    if meta.get("direction"):
        add_p(str(meta["direction"]), size_pt=12, align="center", bold=False, space_after_pt=4, line_spacing=1.1)
    if meta.get("profile"):
        add_p(str(meta["profile"]), size_pt=12, align="center", bold=False, space_after_pt=4, line_spacing=1.1)

    left_ind = 8.2
    student_header_before = 64.0 if work_type != "vkr" else 48.0
    add_p("Выполнил(а):", size_pt=14, align="left", bold=False, left_indent_cm=left_ind, space_before_pt=student_header_before, space_after_pt=2, line_spacing=1.15)
    student_parts = [p for p in [meta.get("course"), meta.get("student_group")] if p]
    if student_parts:
        add_p(f"студент(ка) {', '.join(str(x) for x in student_parts)}", size_pt=14, align="left", left_indent_cm=left_ind, space_after_pt=2, line_spacing=1.15)
    if meta.get("student_name"):
        add_p(str(meta["student_name"]), size_pt=14, align="left", bold=False, left_indent_cm=left_ind, space_after_pt=12, line_spacing=1.15)

    sup_role = str(meta.get("supervisor_role") or "Научный руководитель:")
    add_p(sup_role, size_pt=14, align="left", left_indent_cm=left_ind, space_after_pt=2, line_spacing=1.15)
    if meta.get("supervisor_title"):
        add_p(str(meta["supervisor_title"]), size_pt=14, align="left", left_indent_cm=left_ind, space_after_pt=2, line_spacing=1.15)
    if meta.get("supervisor_name"):
        add_p(str(meta["supervisor_name"]), size_pt=14, align="left", left_indent_cm=left_ind, space_after_pt=10, line_spacing=1.15)

    if meta.get("reviewer_name"):
        add_p("Рецензент:", size_pt=14, align="left", left_indent_cm=left_ind, space_before_pt=4, space_after_pt=2, line_spacing=1.15)
        rev_line = " ".join(str(x) for x in [meta.get("reviewer_title"), meta.get("reviewer_name")] if x)
        add_p(rev_line, size_pt=14, align="left", left_indent_cm=left_ind, space_after_pt=8, line_spacing=1.15)

    city = str(meta.get("city") or "Москва")
    year = str(meta.get("year") or "2026")
    bottom_gap = 135.0 if not meta.get("reviewer_name") else 80.0
    add_p(f"{city} {year}", size_pt=14, align="center", bold=False, space_before_pt=bottom_gap, space_after_pt=0, line_spacing=1.0)

    return elems


def _replace_placeholders_in_xml_paragraph(p_el: etree._Element, replacements: dict[str, str]) -> None:
    """Replace {{key}} placeholders inside a paragraph even if split across runs, preserving <w:br/> line breaks."""
    runs = p_el.findall(f"{{{W_NS}}}r")
    if not runs:
        return
    parts: list[str] = []
    for r in runs:
        for child in r:
            if child.tag == f"{{{W_NS}}}t" and child.text:
                parts.append(child.text)
            elif child.tag == f"{{{W_NS}}}br":
                parts.append("\n")
            elif child.tag == f"{{{W_NS}}}tab":
                parts.append("\t")
    full = "".join(parts)
    if "{{" not in full:
        return
    new_full = full
    for k, v in replacements.items():
        new_full = new_full.replace(f"{{{{{k}}}}}", v)
    if new_full == full:
        return

    first_r = runs[0]
    for extra_r in runs[1:]:
        p_el.remove(extra_r)
    for child in list(first_r):
        if child.tag in (f"{{{W_NS}}}t", f"{{{W_NS}}}br", f"{{{W_NS}}}tab"):
            first_r.remove(child)

    lines = new_full.split("\n")
    for l_idx, seg in enumerate(lines):
        if l_idx > 0:
            first_r.append(OxmlElement("w:br"))
        if seg:
            t = OxmlElement("w:t")
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = seg
            first_r.append(t)


def _ensure_custom_template_p_defaults(p_el: etree._Element, in_table: bool = False) -> None:
    """Prevent imported custom title paragraphs/cells from inheriting target Normal's 1.25cm indent or justify."""
    pPr = p_el.find(f"{{{W_NS}}}pPr")
    if pPr is None:
        pPr = etree.Element(f"{{{W_NS}}}pPr")
        p_el.insert(0, pPr)
    if pPr.find(f"{{{W_NS}}}ind") is None:
        ind = etree.SubElement(pPr, f"{{{W_NS}}}ind")
        ind.set(f"{{{W_NS}}}firstLine", "0")
        ind.set(f"{{{W_NS}}}left", "0")
    if in_table and pPr.find(f"{{{W_NS}}}jc") is None:
        jc = etree.SubElement(pPr, f"{{{W_NS}}}jc")
        jc.set(f"{{{W_NS}}}val", "left")
    if in_table and pPr.find(f"{{{W_NS}}}spacing") is None:
        sp = etree.SubElement(pPr, f"{{{W_NS}}}spacing")
        sp.set(f"{{{W_NS}}}before", "0")
        sp.set(f"{{{W_NS}}}after", "40")
        sp.set(f"{{{W_NS}}}line", "260")
        sp.set(f"{{{W_NS}}}lineRule", "auto")


def import_custom_title_page_elements(
    target_doc: Document,
    template_path: Path,
    meta: dict[str, Any],
    contract: dict[str, Any],
) -> list[etree._Element]:
    """Import paragraphs, tables, and embedded images (e.g., university logo) from a custom .docx title page."""
    tpl = Document(str(template_path))
    replacements = {k: str(v) for k, v in meta.items() if v is not None}
    replacements.setdefault("topic", str(contract.get("topic") or ""))

    # Remap any embedded images (logos/crests) cleanly into target_doc's package
    rid_map: dict[str, str] = {}
    for rel_id, rel in tpl.part.rels.items():
        if "image" in rel.reltype:
            blob = rel.target_part.blob
            new_rid, _ = target_doc.part.get_or_add_image(io.BytesIO(blob))
            rid_map[rel_id] = new_rid

    elems: list[etree._Element] = []
    for child in tpl.element.body:
        if child.tag == f"{{{W_NS}}}sectPr":
            continue
        cloned = copy.deepcopy(child)
        # Remap image r:embed / r:id references
        if rid_map:
            for node in cloned.iter():
                for attr in (f"{{{R_NS}}}embed", f"{{{R_NS}}}id"):
                    old_val = node.get(attr)
                    if old_val and old_val in rid_map:
                        node.set(attr, rid_map[old_val])
        # Substitute {{placeholders}} and prevent Normal's justify/1.25cm indent leak inside template
        is_tbl = cloned.tag == f"{{{W_NS}}}tbl"
        for p_el in ([cloned] if cloned.tag == f"{{{W_NS}}}p" else cloned.findall(f".//{{{W_NS}}}p")):
            _replace_placeholders_in_xml_paragraph(p_el, replacements)
            _ensure_custom_template_p_defaults(p_el, in_table=is_tbl)
        # Tag custom title tables so format_tables_gost() preserves the university's exact template layout
        if cloned.tag == f"{{{W_NS}}}tbl":
            tblPr = cloned.find(f"{{{W_NS}}}tblPr")
            if tblPr is None:
                tblPr = etree.Element(f"{{{W_NS}}}tblPr")
                cloned.insert(0, tblPr)
            desc = tblPr.find(f"{{{W_NS}}}tblDescription")
            if desc is None:
                desc = etree.SubElement(tblPr, f"{{{W_NS}}}tblDescription")
            desc.set(f"{{{W_NS}}}val", "CUSTOM_TITLE_TEMPLATE_TABLE")
        elems.append(cloned)

    return elems


def build_toc_elements(contract: dict[str, Any]) -> list[etree._Element]:
    font = contract["layout"]["font_name"]
    size = float(contract["layout"]["font_size_pt"])
    elems: list[etree._Element] = []

    elems.append(
        make_xml_paragraph(
            "СОДЕРЖАНИЕ",
            font_name=font,
            size_pt=size,
            align="center",
            bold=True,
            space_before_pt=0.0,
            space_after_pt=12.0,
            line_spacing=1.5,
            keep_with_next=True,
            page_break_before=True,
        )
    )

    p_toc = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    ind = OxmlElement("w:ind")
    ind.set(qn("w:firstLine"), "0")
    pPr.append(ind)
    p_toc.append(pPr)

    def _run_with_rpr() -> etree._Element:
        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        rf = OxmlElement("w:rFonts")
        for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
            rf.set(qn(f"w:{attr}"), font)
        rPr.append(rf)
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(int(round(size * 2))))
        rPr.append(sz)
        r.append(rPr)
        return r

    r1 = _run_with_rpr()
    fc_begin = OxmlElement("w:fldChar")
    fc_begin.set(qn("w:fldCharType"), "begin")
    fc_begin.set(qn("w:dirty"), "true")
    r1.append(fc_begin)
    p_toc.append(r1)

    r2 = _run_with_rpr()
    instr = OxmlElement("w:instrText")
    instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr.text = ' TOC \\o "1-3" \\h \\z \\u '
    r2.append(instr)
    p_toc.append(r2)

    r3 = _run_with_rpr()
    fc_sep = OxmlElement("w:fldChar")
    fc_sep.set(qn("w:fldCharType"), "separate")
    r3.append(fc_sep)
    p_toc.append(r3)

    r4 = _run_with_rpr()
    t_ph = OxmlElement("w:t")
    t_ph.text = "Обновите поле оглавления в Word (Ctrl+A → F9)."
    r4.append(t_ph)
    p_toc.append(r4)

    r5 = _run_with_rpr()
    fc_end = OxmlElement("w:fldChar")
    fc_end.set(qn("w:fldCharType"), "end")
    r5.append(fc_end)
    p_toc.append(r5)

    elems.append(p_toc)
    return elems


def format_numbered_equations(doc: Document, contract: dict[str, Any]) -> int:
    """Wrap marked Pandoc OMML display equations into GOST 3-column borderless tables with right-aligned (N)."""
    font = contract["layout"]["font_name"]
    size_pt = float(contract["layout"]["font_size_pt"])
    body = doc.element.body
    children = list(body)
    converted = 0

    for idx, elem in enumerate(children):
        if elem.tag != f"{{{W_NS}}}p":
            continue
        txt = "".join(t.text or "" for t in elem.findall(f".//{{{W_NS}}}t")).strip()
        m = re.fullmatch(re.escape(EQ_MARKER_PREFIX) + r"([^@]+)@@", txt)
        if not m or idx == 0:
            continue

        eq_num = m.group(1).strip()
        math_p = children[idx - 1]
        if math_p.tag != f"{{{W_NS}}}p":
            body.remove(elem)
            continue

        # Normalize equation paragraph inside center cell
        mpPr = math_p.find(f"{{{W_NS}}}pPr")
        if mpPr is None:
            mpPr = etree.Element(f"{{{W_NS}}}pPr")
            math_p.insert(0, mpPr)
        for tag in ("jc", "ind", "spacing"):
            for old in mpPr.findall(f"{{{W_NS}}}{tag}"):
                mpPr.remove(old)
        jc = etree.SubElement(mpPr, f"{{{W_NS}}}jc")
        jc.set(f"{{{W_NS}}}val", "center")
        ind = etree.SubElement(mpPr, f"{{{W_NS}}}ind")
        ind.set(f"{{{W_NS}}}firstLine", "0")
        ind.set(f"{{{W_NS}}}left", "0")
        ind.set(f"{{{W_NS}}}right", "0")
        sp = etree.SubElement(mpPr, f"{{{W_NS}}}spacing")
        sp.set(f"{{{W_NS}}}before", "120")
        sp.set(f"{{{W_NS}}}after", "120")
        sp.set(f"{{{W_NS}}}line", "240")
        sp.set(f"{{{W_NS}}}lineRule", "auto")

        # Build borderless 3-column GOST equation layout table: 12% | 76% | 12%
        tbl = OxmlElement("w:tbl")
        tblPr = parse_xml(
            f'<w:tblPr {nsdecls("w")}>'
            f'  <w:tblW w:w="5000" w:type="pct"/>'
            f'  <w:jc w:val="center"/>'
            f'  <w:tblBorders>'
            f'    <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'    <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'    <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'    <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'    <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'    <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'  </w:tblBorders>'
            f'  <w:tblCellMar>'
            f'    <w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>'
            f'    <w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>'
            f'  </w:tblCellMar>'
            f'  <w:tblDescription w:val="{EQ_TBL_DESC}"/>'
            f"</w:tblPr>"
        )
        tbl.append(tblPr)

        tr = OxmlElement("w:tr")
        trPr = OxmlElement("w:trPr")
        trPr.append(OxmlElement("w:cantSplit"))
        tr.append(trPr)

        def _make_cell(pct_w: int, content_p: etree._Element) -> etree._Element:
            tc = OxmlElement("w:tc")
            tcPr = parse_xml(
                f'<w:tcPr {nsdecls("w")}>'
                f'  <w:tcW w:w="{pct_w}" w:type="pct"/>'
                f'  <w:vAlign w:val="center"/>'
                f"</w:tcPr>"
            )
            tc.append(tcPr)
            tc.append(content_p)
            return tc

        empty_p = make_xml_paragraph("", font_name=font, size_pt=size_pt, align="left")
        num_p = make_xml_paragraph(f"({eq_num})", font_name=font, size_pt=size_pt, align="right", line_spacing=1.0)
        eq_clone = copy.deepcopy(math_p)

        tr.append(_make_cell(600, empty_p))
        tr.append(_make_cell(3800, eq_clone))
        tr.append(_make_cell(600, num_p))
        tbl.append(tr)

        pos = list(body).index(math_p)
        body.insert(pos, tbl)
        body.remove(math_p)
        body.remove(elem)
        converted += 1

    return converted


def enable_update_fields_on_open(doc: Document) -> None:
    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is None:
        uf = OxmlElement("w:updateFields")
        uf.set(qn("w:val"), "true")
        settings.append(uf)


def add_page_number_field(paragraph: Any, font_name: str, size_pt: float = 12.0) -> None:
    def _new_run() -> Any:
        r = paragraph.add_run()
        set_run_font(r, font_name, size_pt, bold=False)
        return r

    r1 = _new_run()
    fc1 = OxmlElement("w:fldChar")
    fc1.set(qn("w:fldCharType"), "begin")
    r1._r.append(fc1)

    r2 = _new_run()
    instr = OxmlElement("w:instrText")
    instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr.text = " PAGE "
    r2._r.append(instr)

    r3 = _new_run()
    fc2 = OxmlElement("w:fldChar")
    fc2.set(qn("w:fldCharType"), "separate")
    r3._r.append(fc2)

    r4 = _new_run()
    fc3 = OxmlElement("w:fldChar")
    fc3.set(qn("w:fldCharType"), "end")
    r4._r.append(fc3)


def style_footnotes_xml(doc: Document, font_name: str, fn_size_pt: float) -> int:
    fn_part = None
    for part in doc.part.package.parts:
        if str(part.partname).endswith("footnotes.xml"):
            fn_part = part
            break
    if fn_part is None:
        return 0

    root = etree.fromstring(fn_part.blob)
    sz_val = str(int(round(fn_size_pt * 2)))
    count = 0
    for fn in root.findall(f"{{{W_NS}}}footnote"):
        ftype = fn.get(f"{{{W_NS}}}type")
        if ftype in ("separator", "continuationSeparator"):
            continue
        count += 1
        for p in fn.findall(f"{{{W_NS}}}p"):
            pPr = p.find(f"{{{W_NS}}}pPr")
            if pPr is None:
                pPr = etree.Element(f"{{{W_NS}}}pPr")
                p.insert(0, pPr)
            for old in pPr.findall(f"{{{W_NS}}}jc"):
                pPr.remove(old)
            jc = etree.SubElement(pPr, f"{{{W_NS}}}jc")
            jc.set(f"{{{W_NS}}}val", "both")

            for old in pPr.findall(f"{{{W_NS}}}spacing"):
                pPr.remove(old)
            sp = etree.SubElement(pPr, f"{{{W_NS}}}spacing")
            sp.set(f"{{{W_NS}}}before", "0")
            sp.set(f"{{{W_NS}}}after", "0")
            sp.set(f"{{{W_NS}}}line", "240")
            sp.set(f"{{{W_NS}}}lineRule", "auto")

            for old in pPr.findall(f"{{{W_NS}}}ind"):
                pPr.remove(old)
            ind = etree.SubElement(pPr, f"{{{W_NS}}}ind")
            ind.set(f"{{{W_NS}}}firstLine", "0")
            ind.set(f"{{{W_NS}}}left", "0")
            ind.set(f"{{{W_NS}}}right", "0")

            for r in p.findall(f"{{{W_NS}}}r"):
                rPr = r.find(f"{{{W_NS}}}rPr")
                if rPr is None:
                    rPr = etree.Element(f"{{{W_NS}}}rPr")
                    r.insert(0, rPr)
                for tag in ("rFonts", "sz", "szCs", "color"):
                    for old in rPr.findall(f"{{{W_NS}}}{tag}"):
                        rPr.remove(old)
                rf = etree.SubElement(rPr, f"{{{W_NS}}}rFonts")
                for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
                    rf.set(f"{{{W_NS}}}{attr}", font_name)
                col = etree.SubElement(rPr, f"{{{W_NS}}}color")
                col.set(f"{{{W_NS}}}val", "000000")
                sz = etree.SubElement(rPr, f"{{{W_NS}}}sz")
                sz.set(f"{{{W_NS}}}val", sz_val)
                szc = etree.SubElement(rPr, f"{{{W_NS}}}szCs")
                szc.set(f"{{{W_NS}}}val", sz_val)
                if r.find(f"{{{W_NS}}}footnoteRef") is not None:
                    if rPr.find(f"{{{W_NS}}}vertAlign") is None:
                        va = etree.SubElement(rPr, f"{{{W_NS}}}vertAlign")
                        va.set(f"{{{W_NS}}}val", "superscript")

    fn_part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return count


def _table_desc(tbl: Any) -> str:
    tblPr = tbl._tbl.tblPr
    if tblPr is None:
        return ""
    desc = tblPr.find(qn("w:tblDescription"))
    return desc.get(qn("w:val"), "") if desc is not None else ""


def format_tables_gost(doc: Document, contract: dict[str, Any]) -> int:
    layout = contract["layout"]
    font = layout["font_name"]
    tbl_size = float(layout["table_font_size_pt"])
    short_max = int(layout.get("short_table_max_rows", 5))

    borders_xml = (
        f'<w:tblBorders {nsdecls("w")}>'
        '  <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '  <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '  <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        "</w:tblBorders>"
    )
    cell_mar_xml = (
        f'<w:tblCellMar {nsdecls("w")}>'
        '  <w:top w:w="60" w:type="dxa"/>'
        '  <w:left w:w="100" w:type="dxa"/>'
        '  <w:bottom w:w="60" w:type="dxa"/>'
        '  <w:right w:w="100" w:type="dxa"/>'
        "</w:tblCellMar>"
    )

    data_tables = 0
    for tbl in doc.tables:
        if _table_desc(tbl) in (EQ_TBL_DESC, "CUSTOM_TITLE_TEMPLATE_TABLE"):
            continue
        data_tables += 1
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tblPr = tbl._tbl.tblPr
        for tag in ("w:tblBorders", "w:tblCellMar", "w:tblW"):
            old = tblPr.find(qn(tag))
            if old is not None:
                tblPr.remove(old)
        tblW = OxmlElement("w:tblW")
        tblW.set(qn("w:w"), "5000")
        tblW.set(qn("w:type"), "pct")
        tblPr.append(tblW)
        tblPr.append(parse_xml(borders_xml))
        tblPr.append(parse_xml(cell_mar_xml))

        row_cnt = len(tbl.rows)
        keep_whole = row_cnt <= short_max

        for ridx, row in enumerate(tbl.rows):
            trPr = row._tr.get_or_add_trPr()
            if trPr.find(qn("w:cantSplit")) is None:
                trPr.append(OxmlElement("w:cantSplit"))
            if ridx == 0 and trPr.find(qn("w:tblHeader")) is None:
                trPr.append(OxmlElement("w:tblHeader"))

            for cell in row.cells:
                for p in cell.paragraphs:
                    pf = p.paragraph_format
                    pf.first_line_indent = Cm(0)
                    pf.left_indent = Cm(0)
                    pf.right_indent = Cm(0)
                    pf.space_before = Pt(0)
                    pf.space_after = Pt(0)
                    pf.line_spacing = 1.05
                    pf.keep_with_next = keep_whole and (ridx < row_cnt - 1)
                    if ridx == 0:
                        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in p.runs:
                        set_run_font(run, font, tbl_size, bold=(True if ridx == 0 else None))
    return data_tables


def format_body_paragraphs(doc: Document, contract: dict[str, Any]) -> None:
    layout = contract["layout"]
    font = layout["font_name"]
    body_size = float(layout["font_size_pt"])
    line_sp = float(layout["line_spacing"])
    indent_cm = float(layout["first_line_indent_cm"])

    paras = doc.paragraphs
    for idx, p in enumerate(paras):
        sname = p.style.name if p.style else ""
        txt = p.text.strip()
        pf = p.paragraph_format

        # Do not alter standalone OMML display equation paragraphs before format_numbered_equations()
        if p._p.findall(f".//{{{M_NS}}}oMathPara"):
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.first_line_indent = Cm(0)
            pf.space_before = Pt(6)
            pf.space_after = Pt(6)
            continue

        if sname == "Heading 1":
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.first_line_indent = Cm(0)
            pf.left_indent = Cm(0)
            pf.space_before = Pt(0)
            pf.space_after = Pt(12)
            pf.line_spacing = 1.5
            pf.keep_with_next = True
            pf.page_break_before = True
            for r in p.runs:
                set_run_font(r, font, body_size, bold=True)
            continue

        if sname in ("Heading 2", "Heading 3"):
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf.first_line_indent = Cm(indent_cm)
            pf.left_indent = Cm(0)
            pf.space_before = Pt(12 if sname == "Heading 2" else 8)
            pf.space_after = Pt(6 if sname == "Heading 2" else 4)
            pf.line_spacing = 1.5
            pf.keep_with_next = True
            for r in p.runs:
                set_run_font(r, font, body_size, bold=True)
            continue

        if re.fullmatch(r"Таблица\s+\d+(?:\.\d+)?", txt):
            pf.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            pf.first_line_indent = Cm(0)
            pf.space_before = Pt(6)
            pf.space_after = Pt(0)
            pf.line_spacing = 1.15
            pf.keep_with_next = True
            for r in p.runs:
                set_run_font(r, font, body_size, bold=False)
            if idx + 1 < len(paras):
                nxt = paras[idx + 1]
                nxt_txt = nxt.text.strip()
                if nxt_txt and not nxt_txt.startswith("Таблица"):
                    npf = nxt.paragraph_format
                    npf.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    npf.first_line_indent = Cm(0)
                    npf.space_before = Pt(0)
                    npf.space_after = Pt(4)
                    npf.line_spacing = 1.15
                    npf.keep_with_next = True
                    for r in nxt.runs:
                        set_run_font(r, font, body_size, bold=True)
            continue

        if re.match(r"^Таблица\s+\d+(?:\.\d+)?\s*[—–-]\s+\S", txt):
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf.first_line_indent = Cm(0)
            pf.space_before = Pt(6)
            pf.space_after = Pt(4)
            pf.line_spacing = 1.15
            pf.keep_with_next = True
            for r in p.runs:
                set_run_font(r, font, body_size, bold=False)
            continue

        if idx > 0 and re.fullmatch(r"Таблица\s+\d+(?:\.\d+)?", paras[idx - 1].text.strip()):
            continue

        if re.match(r"^Рисунок\s+\d+(?:\.\d+)?\s*[—–-]", txt):
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.first_line_indent = Cm(0)
            pf.space_before = Pt(4)
            next_is_source = (idx + 1 < len(paras) and bool(re.match(r"^Источник:", paras[idx + 1].text.strip())))
            pf.space_after = Pt(2 if next_is_source else 8)
            pf.line_spacing = 1.15
            pf.keep_with_next = next_is_source
            for r in p.runs:
                set_run_font(r, font, 12.0, bold=False)
            continue

        if re.match(r"^Источник:", txt):
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf.first_line_indent = Cm(0)
            pf.space_before = Pt(2)
            pf.space_after = Pt(8)
            pf.line_spacing = 1.15
            for r in p.runs:
                set_run_font(r, font, 12.0, bold=False)
            continue

        if p._p.findall(f".//{{{W_NS}}}drawing"):
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.first_line_indent = Cm(0)
            pf.space_before = Pt(6)
            pf.space_after = Pt(2)
            pf.keep_with_next = True
            continue

        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.first_line_indent = Cm(indent_cm)
        pf.left_indent = Cm(0)
        pf.right_indent = Cm(0)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = line_sp
        pf.widow_control = True
        for r in p.runs:
            if r._r.find(qn("w:footnoteReference")) is not None:
                set_run_font(r, font, body_size, bold=False, italic=False)
                rPr = r._r.get_or_add_rPr()
                for old_va in rPr.findall(qn("w:vertAlign")):
                    rPr.remove(old_va)
                va = OxmlElement("w:vertAlign")
                va.set(qn("w:val"), "superscript")
                rPr.append(va)
                continue
            set_run_font(r, font, body_size)


def configure_page_setup_and_numbers(doc: Document, contract: dict[str, Any]) -> None:
    layout = contract["layout"]
    margins = layout["margins_mm"]
    font = layout["font_name"]
    pos = layout.get("page_number_position", "bottom_center")
    include_title = bool(layout.get("include_title_page", True))

    sec = doc.sections[0]
    sec.page_width = Mm(210)
    sec.page_height = Mm(297)
    sec.left_margin = Mm(float(margins["left"]))
    sec.right_margin = Mm(float(margins["right"]))
    sec.top_margin = Mm(float(margins["top"]))
    sec.bottom_margin = Mm(float(margins["bottom"]))
    sec.different_first_page_header_footer = include_title

    if include_title:
        for p in sec.first_page_footer.paragraphs:
            p.text = ""
        for p in sec.first_page_header.paragraphs:
            p.text = ""

    for p in sec.footer.paragraphs:
        p.text = ""
    for p in sec.header.paragraphs:
        p.text = ""

    target_hf = sec.header if pos == "top_center" else sec.footer
    p_num = target_hf.paragraphs[0]
    p_num.paragraph_format.first_line_indent = Cm(0)
    p_num.paragraph_format.space_before = Pt(0)
    p_num.paragraph_format.space_after = Pt(0)
    p_num.paragraph_format.alignment = (
        WD_ALIGN_PARAGRAPH.RIGHT if pos == "bottom_right" else WD_ALIGN_PARAGRAPH.CENTER
    )
    add_page_number_field(p_num, font_name=font, size_pt=12.0)


def main() -> int:
    ensure_utf8_io()
    parser = argparse.ArgumentParser(
        description="Build GOST-compliant Word (.docx) with native footnotes, [@key] auto-citations, OMML equations, and live TOC."
    )
    parser.add_argument("--drafts", default="drafts", help="Directory containing ordered Markdown draft files")
    parser.add_argument("--meta", default="meta.json", help="Path to title page meta.json")
    parser.add_argument("--contract", default="project_contract.json", help="Path to project_contract.json")
    parser.add_argument("--out", default="output/final.docx", help="Output .docx path")
    parser.add_argument(
        "--title-template",
        default="",
        help="Optional custom university title page .docx template (supports {{meta_key}} placeholders & embedded logos)",
    )
    parser.add_argument("--strict-bib", action="store_true", help="Fail build if any [@key] bibliography entry is uncited")
    parser.add_argument("--no-word-toc", action="store_true", help="Skip live Microsoft Word COM TOC update")
    args = parser.parse_args()

    pandoc_exe = find_pandoc()
    if not pandoc_exe:
        print("ERROR: pandoc not found. Install via: python -m pip install pypandoc-binary", file=sys.stderr)
        return 2

    drafts_dir = Path(args.drafts).resolve()
    if not drafts_dir.exists():
        print(f"ERROR: drafts directory not found: {drafts_dir}", file=sys.stderr)
        return 2

    files = collect_draft_files(drafts_dir)
    if not files:
        print(f"ERROR: no draft .md files found in {drafts_dir}", file=sys.stderr)
        return 2

    contract = load_contract(args.contract)
    meta: dict[str, Any] = {}
    meta_path = Path(args.meta)
    if meta_path.exists():
        meta = json.loads(meta_path.read_text(encoding="utf-8"))

    title_tpl_arg = args.title_template or str(contract.get("layout", {}).get("title_template_path") or "")
    title_tpl_path = Path(title_tpl_arg).resolve() if title_tpl_arg else None
    if title_tpl_path and not title_tpl_path.exists():
        print(f"ERROR: --title-template file not found: {title_tpl_path}", file=sys.stderr)
        return 2

    try:
        combined_md, warnings, key_to_num, eq_planned = normalize_markdown_for_pandoc(
            files, contract, strict_bib=args.strict_bib
        )
    except ValueError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    out_path = Path(args.out).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Resource search paths so Markdown figures like ![](data/fig_1.png) resolve reliably from any CWD
    res_dirs = [str(Path.cwd().resolve()), str(drafts_dir), str(drafts_dir.parent), str(drafts_dir.parent / "data")]
    resource_path = os.pathsep.join(dict.fromkeys(res_dirs))

    with tempfile.TemporaryDirectory(prefix="gost_build_") as tmpdir:
        tmp = Path(tmpdir)
        ref_docx = tmp / "reference.docx"
        combined_md_path = tmp / "combined.md"
        pandoc_docx = tmp / "pandoc_body.docx"

        build_reference_docx(ref_docx, contract)
        combined_md_path.write_text(combined_md, encoding="utf-8")

        cmd = [
            pandoc_exe,
            str(combined_md_path),
            "-f",
            "markdown+footnotes+pipe_tables+tex_math_dollars+raw_attribute+link_attributes",
            "-t",
            "docx",
            f"--reference-doc={ref_docx}",
            f"--resource-path={resource_path}",
            "-o",
            str(pandoc_docx),
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=90)
        if res.returncode != 0 or not pandoc_docx.exists():
            print(f"ERROR: pandoc failed (rc={res.returncode}): {res.stderr}", file=sys.stderr)
            return 1

        # Open Pandoc DOCX IN PLACE so native footnotes, OMML equations, images & relationships stay 100% intact
        doc = Document(str(pandoc_docx))
        ensure_toc_and_footnote_styles(doc, contract)
        configure_page_setup_and_numbers(doc, contract)
        format_body_paragraphs(doc, contract)
        eq_formatted = format_numbered_equations(doc, contract)
        data_tables = format_tables_gost(doc, contract)
        fn_count = style_footnotes_xml(
            doc,
            font_name=contract["layout"]["font_name"],
            fn_size_pt=float(contract["layout"]["footnote_font_size_pt"]),
        )

        prepend_elems: list[etree._Element] = []
        title_mode = "none"
        if contract["layout"].get("include_title_page", True):
            if title_tpl_path:
                prepend_elems.extend(import_custom_title_page_elements(doc, title_tpl_path, meta, contract))
                title_mode = f"custom template ({title_tpl_path.name})"
            elif meta:
                prepend_elems.extend(build_title_page_elements(meta, contract))
                title_mode = "built-in GOST"

        if contract["layout"].get("include_toc", True):
            prepend_elems.extend(build_toc_elements(contract))
            enable_update_fields_on_open(doc)

        body = doc.element.body
        for idx, elem in enumerate(prepend_elems):
            body.insert(idx, elem)

        if not prepend_elems and doc.paragraphs:
            doc.paragraphs[0].paragraph_format.page_break_before = False

        doc.save(str(out_path))

    toc_updated_msg = "skipped (press F9 in Word)"
    if not args.no_word_toc and contract["layout"].get("include_toc", True) and sys.platform == "win32":
        ok, engine = update_word_toc_and_export_pdf(out_path, pdf_path=None, update_toc=True)
        if ok:
            toc_updated_msg = f"LIVE UPDATED via {engine}"
        else:
            toc_updated_msg = f"deferred ({engine}; press F9 in Word)"

    print(f"BUILD OK: {out_path}")
    print(f"  Draft files merged ({len(files)}): {', '.join(f.name for f in files)}")
    print(f"  Title page: {title_mode}")
    print(f"  Native Word footnotes styled: {fn_count}")
    if key_to_num:
        print(f"  Symbolic [@key] bibliography auto-sorted & resolved: {len(key_to_num)} entries")
    print(f"  GOST data tables: {data_tables} | Numbered OMML equations: {eq_formatted}")
    print(f"  TOC update: {toc_updated_msg}")
    for w in warnings:
        print(f"  [WARN] {w}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
