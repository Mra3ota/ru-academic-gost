#!/usr/bin/env python3
"""Assemble GOST-compliant Word (.docx) from Markdown drafts via Pandoc + in-place python-docx + Word TOC update."""

from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

from docx import Document
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
    update_word_toc_and_export_pdf,
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def set_run_font(run: Any, font_name: str, size_pt: float, bold: bool | None = None, italic: bool | None = None) -> None:
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

    # Configure TOC 1/2/3 styles so Word TOC never inherits Normal's 1.25cm first-line indent
    from docx.enum.style import WD_STYLE_TYPE

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

    # Ensure Footnote Reference character style is strictly superscript
    for ref_sname in ("Footnote Reference", "FootnoteReference"):
        if ref_sname in doc.styles:
            rst = doc.styles[ref_sname]
        else:
            rst = doc.styles.add_style(ref_sname, WD_STYLE_TYPE.CHARACTER)
        rst.font.name = font_name
        rst.font.color.rgb = RGBColor(0, 0, 0)
        rPr = rst.element.get_or_add_rPr()
        if rPr.find(qn("w:vertAlign")) is None:
            va = OxmlElement("w:vertAlign")
            va.set(qn("w:val"), "superscript")
            rPr.append(va)

    doc.save(str(ref_path))


def collect_draft_files(drafts_dir: Path) -> list[Path]:
    files = sorted(
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
    return files


def normalize_markdown_for_pandoc(files: list[Path]) -> tuple[str, list[str]]:
    """Sanitize drafts, auto-namespace footnotes across files, fix spacing, and escape biblio list dots."""
    warnings: list[str] = []
    combined_parts: list[str] = []

    for fidx, fpath in enumerate(files, start=1):
        text = clean_hidden_unicode(fpath.read_text(encoding="utf-8"))
        # Strip BOM/trailing spaces
        text = text.replace("\r\n", "\n").replace("\r", "\n")

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

        # Normalize footnote punctuation spacing (GOST: «слово[^id].», never «слово [^id] .» or «слово.[^id]»)
        # 1) Remove spaces before footnote reference (except at start of line for definitions)
        text = re.sub(r"(?<=[^\s\n])[ \t]+(\[\^[^\]]+\])(?!:)", r"\1", text)
        # 2) Move comma/semicolon/period AFTER footnote reference ONLY when not a Russian abbreviation
        #    (e.g. keep 'руб.[^1].', 'г.[^2].', 'др.[^3].' intact) and no period already follows [^id]
        text = re.sub(r"([,;:])(\[\^[^\]]+\])(?!:)", r"\2\1", text)
        text = re.sub(
            r"(?<!\bруб)(?<!\bтыс)(?<!\bмлн)(?<!\bмлрд)(?<!\bст)(?<!\bдр)(?<!\bсм)(?<!\bпр)(?<!\bгг)(?<!\bвв)(?<!\s[гвс])(?<![а-яё]\.[а-яё])\.(\[\^[^\]]+\])(?!:|\.)",
            r"\1.",
            text,
        )
        # 3) Remove spaces between footnote reference and following punctuation
        text = re.sub(r"(\[\^[^\]]+\])[ \t]+([.,;:])", r"\1\2", text)

        # Strip trailing periods from Markdown headings (# / ## / ###) per GOST 7.32
        def _clean_heading(m: re.Match[str]) -> str:
            hashes, title = m.group(1), m.group(2).strip()
            title = re.sub(r"\.$", "", title)
            return f"{hashes} {title}"

        text = re.sub(r"^(#{1,3})\s+(.+)$", _clean_heading, text, flags=re.M)

        # In bibliography file, escape "1. " -> "1\. " AND separate items with blank lines
        # so Pandoc renders each numbered entry as its own crisp GOST paragraph (not glued together or hanging list)
        is_bib = "biblio" in fpath.name.lower() or bool(
            re.search(r"^#\s*список\s+(?:использованн\w+\s+)?(?:источников|литературы)", text, re.I | re.M)
        )
        if is_bib:
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

    return "\n\n".join(combined_parts) + "\n", warnings


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

    # Top university header block
    if meta.get("ministry"):
        add_p(str(meta["ministry"]).upper(), size_pt=12, align="center", bold=False, space_after_pt=4, line_spacing=1.0)
    if meta.get("university"):
        add_p(str(meta["university"]), size_pt=12, align="center", bold=True, space_after_pt=4, line_spacing=1.05)
    if meta.get("institute"):
        add_p(str(meta["institute"]), size_pt=12, align="center", bold=False, space_after_pt=3, line_spacing=1.05)
    if meta.get("department"):
        add_p(str(meta["department"]), size_pt=12, align="center", bold=False, space_after_pt=6, line_spacing=1.05)

    # Center work title block (exact vertical offset via space_before)
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

    # Right-side student & supervisor block (left_indent=8.2 cm fits nicely in 16.5 cm printable width)
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

    # Bottom city & year (cleanly positioned near bottom margin of Page 1 without spilling)
    city = str(meta.get("city") or "Москва")
    year = str(meta.get("year") or "2026")
    bottom_gap = 135.0 if not meta.get("reviewer_name") else 80.0
    add_p(f"{city} {year}", size_pt=14, align="center", bold=False, space_before_pt=bottom_gap, space_after_pt=0, line_spacing=1.0)

    return elems


def build_toc_elements(contract: dict[str, Any]) -> list[etree._Element]:
    font = contract["layout"]["font_name"]
    size = float(contract["layout"]["font_size_pt"])
    elems: list[etree._Element] = []

    # Non-Heading1 TOC title so it never appears inside itself
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

    # Word native TOC field
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
    """Style every footnote in word/footnotes.xml directly (TNR 10pt, 1.0 spacing, justified, 0 indent)."""
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
            # Alignment: justify
            for old in pPr.findall(f"{{{W_NS}}}jc"):
                pPr.remove(old)
            jc = etree.SubElement(pPr, f"{{{W_NS}}}jc")
            jc.set(f"{{{W_NS}}}val", "both")

            # Spacing: single 1.0, 0 before/after
            for old in pPr.findall(f"{{{W_NS}}}spacing"):
                pPr.remove(old)
            sp = etree.SubElement(pPr, f"{{{W_NS}}}spacing")
            sp.set(f"{{{W_NS}}}before", "0")
            sp.set(f"{{{W_NS}}}after", "0")
            sp.set(f"{{{W_NS}}}line", "240")
            sp.set(f"{{{W_NS}}}lineRule", "auto")

            # Indent: 0
            for old in pPr.findall(f"{{{W_NS}}}ind"):
                pPr.remove(old)
            ind = etree.SubElement(pPr, f"{{{W_NS}}}ind")
            ind.set(f"{{{W_NS}}}firstLine", "0")
            ind.set(f"{{{W_NS}}}left", "0")
            ind.set(f"{{{W_NS}}}right", "0")

            # Runs inside footnote
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
                # Ensure the leading footnoteRef mark inside footnotes.xml is superscript
                if r.find(f"{{{W_NS}}}footnoteRef") is not None:
                    if rPr.find(f"{{{W_NS}}}vertAlign") is None:
                        va = etree.SubElement(rPr, f"{{{W_NS}}}vertAlign")
                        va.set(f"{{{W_NS}}}val", "superscript")

    fn_part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return count


def format_tables_gost(doc: Document, contract: dict[str, Any]) -> None:
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

    for tbl in doc.tables:
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

        # Table label line ("Таблица 1") -> right-aligned, keep_with_next
        if re.fullmatch(r"Таблица\s+\d+(?:\.\d+)?", txt):
            pf.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            pf.first_line_indent = Cm(0)
            pf.space_before = Pt(6)
            pf.space_after = Pt(0)
            pf.line_spacing = 1.15
            pf.keep_with_next = True
            for r in p.runs:
                set_run_font(r, font, body_size, bold=False)
            # Also style immediately following table title line if present
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

        # Single-line GOST 7.32 table caption ("Таблица 1 — Название")
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

        # Already styled table title following "Таблица N"
        if idx > 0 and re.fullmatch(r"Таблица\s+\d+(?:\.\d+)?", paras[idx - 1].text.strip()):
            continue

        # Figure caption ("Рисунок 1 — Название")
        if re.match(r"^Рисунок\s+\d+(?:\.\d+)?\s*[—–-]", txt):
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.first_line_indent = Cm(0)
            pf.space_before = Pt(4)
            pf.space_after = Pt(8)
            pf.line_spacing = 1.15
            for r in p.runs:
                set_run_font(r, font, 12.0, bold=False)
            continue

        # Source note below table/figure
        if re.match(r"^Источник:", txt):
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf.first_line_indent = Cm(0)
            pf.space_before = Pt(2)
            pf.space_after = Pt(8)
            pf.line_spacing = 1.15
            for r in p.runs:
                set_run_font(r, font, 12.0, bold=False)
            continue

        # Image paragraph (contains drawing)
        if p._p.findall(f".//{{{W_NS}}}drawing"):
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.first_line_indent = Cm(0)
            pf.keep_with_next = True
            continue

        # Standard body paragraph
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.first_line_indent = Cm(indent_cm)
        pf.left_indent = Cm(0)
        pf.right_indent = Cm(0)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = line_sp
        pf.widow_control = True
        for r in p.runs:
            # Explicitly superscript every footnoteReference mark in document.xml
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

    # Ensure first-page header & footer are blank (no page number "1" on title page)
    if include_title:
        for p in sec.first_page_footer.paragraphs:
            p.text = ""
        for p in sec.first_page_header.paragraphs:
            p.text = ""

    # Clear default header/footer before adding page number
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
        description="Build GOST-compliant Word (.docx) with native footnotes, title page, and live TOC."
    )
    parser.add_argument("--drafts", default="drafts", help="Directory containing ordered Markdown draft files")
    parser.add_argument("--meta", default="meta.json", help="Path to title page meta.json")
    parser.add_argument("--contract", default="project_contract.json", help="Path to project_contract.json")
    parser.add_argument("--out", default="output/final.docx", help="Output .docx path")
    parser.add_argument("--no-word-toc", action="store_true", help="Skip live Microsoft Word COM TOC update")
    args = parser.parse_args()

    pandoc_exe = find_pandoc()
    if not pandoc_exe:
        print("ERROR: pandoc not found. Install via: python -m pip install pypandoc-binary", file=sys.stderr)
        return 2

    drafts_dir = Path(args.drafts)
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

    try:
        combined_md, warnings = normalize_markdown_for_pandoc(files)
    except ValueError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    out_path = Path(args.out).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

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
            "markdown+footnotes+pipe_tables+tex_math_dollars+raw_attribute",
            "-t",
            "docx",
            f"--reference-doc={ref_docx}",
            "-o",
            str(pandoc_docx),
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=90)
        if res.returncode != 0 or not pandoc_docx.exists():
            print(f"ERROR: pandoc failed (rc={res.returncode}): {res.stderr}", file=sys.stderr)
            return 1

        # Open Pandoc DOCX IN PLACE so native footnotes, images & relationships stay 100% intact
        doc = Document(str(pandoc_docx))
        # Ensure TOC 1/2/3 styles exist directly on final doc with 0 first-line indent (Pandoc omits unused styles)
        from docx.enum.style import WD_STYLE_TYPE

        font_n = contract["layout"]["font_name"]
        body_sz = float(contract["layout"]["font_size_pt"])
        for level, left_cm in ((1, 0.0), (2, 0.6), (3, 1.2)):
            sname = f"TOC {level}"
            st = doc.styles[sname] if sname in doc.styles else doc.styles.add_style(sname, WD_STYLE_TYPE.PARAGRAPH)
            configure_style(
                st,
                font_n,
                body_sz,
                line_spacing=1.15,
                first_line_cm=0.0,
                align=WD_ALIGN_PARAGRAPH.LEFT,
                bold=False,
                space_before_pt=2.0,
                space_after_pt=2.0,
            )
            st.paragraph_format.left_indent = Cm(left_cm)

        configure_page_setup_and_numbers(doc, contract)
        format_body_paragraphs(doc, contract)
        format_tables_gost(doc, contract)
        fn_count = style_footnotes_xml(
            doc,
            font_name=contract["layout"]["font_name"],
            fn_size_pt=float(contract["layout"]["footnote_font_size_pt"]),
        )

        # Prepend Title Page + TOC directly into doc.element.body
        prepend_elems: list[etree._Element] = []
        if contract["layout"].get("include_title_page", True) and meta:
            prepend_elems.extend(build_title_page_elements(meta, contract))
        if contract["layout"].get("include_toc", True):
            prepend_elems.extend(build_toc_elements(contract))
            enable_update_fields_on_open(doc)

        body = doc.element.body
        for idx, elem in enumerate(prepend_elems):
            body.insert(idx, elem)

        # If no title page and no TOC, remove unnecessary leading page_break_before on very first paragraph
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
    print(f"  Native Word footnotes styled: {fn_count}")
    print(f"  Tables formatted: {len(Document(str(out_path)).tables)}")
    print(f"  TOC update: {toc_updated_msg}")
    for w in warnings:
        print(f"  [WARN] {w}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
