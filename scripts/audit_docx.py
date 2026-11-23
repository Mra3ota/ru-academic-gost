#!/usr/bin/env python3
"""9-point GOST, hygiene, bibliography, PDF pagination, and PNG visual preview auditor."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

try:
    import pymupdf as fitz
except ImportError:
    import fitz  # type: ignore
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _common import (
    BANNED_UNICODE_RE,
    CYRILLIC_CHAR_RE,
    clean_hidden_unicode,
    ensure_utf8_io,
    export_docx_to_pdf,
    find_hidden_unicode,
    find_homoglyphs,
    load_contract,
    render_pdf_to_pngs,
    ru_sort_key,
    strip_bib_prefix,
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

PLACEHOLDER_PATTERNS: list[re.Pattern[str]] = [
    re.compile(r"_{3,}"),
    re.compile(r"\[\s*\.\.\.\s*\]"),
    re.compile(r"«\s*\.\.\.\s*»"),
    re.compile(r"\{\{[A-Za-z0-9_]+\}\}"),
    re.compile(r"@@GOST_EQ_NUM:[^@]+@@"),
    re.compile(r"\[@[A-Za-zА-Яа-яЁё0-9_\-]+"),
    re.compile(r"\b(?:TODO|TBD|FIXME|ХХ-000|XX-000|00\.00\.00)\b"),
]

SECTION_HEADING_RE = re.compile(
    r"^(?:ВВЕДЕНИЕ|ГЛАВА\s+\d+|ЗАКЛЮЧЕНИЕ|СПИСОК\s+(?:ИСПОЛЬЗОВАНН\w+\s+)?(?:ИСТОЧНИКОВ|ЛИТЕРАТУРЫ)|ПРИЛОЖЕНИЕ\s+[А-Я])\b",
    re.I,
)


def extract_footnotes_texts(doc: Document) -> list[str]:
    for part in doc.part.package.parts:
        if str(part.partname).endswith("footnotes.xml"):
            root = etree.fromstring(part.blob)
            out: list[str] = []
            for fn in root.findall(f"{{{W_NS}}}footnote"):
                if fn.get(f"{{{W_NS}}}type") in ("separator", "continuationSeparator"):
                    continue
                txt = "".join(t.text or "" for t in fn.findall(f".//{{{W_NS}}}t")).strip()
                out.append(txt)
            return out
    return []


def auto_fix_docx(docx_path: Path) -> int:
    """Safely strip hidden Unicode, replace U+2011 with '-', and clean spaces before footnoteReference."""
    doc = Document(str(docx_path))
    fixes = 0

    def _clean_paragraph_xml(p_el: etree._Element) -> None:
        nonlocal fixes
        runs = p_el.findall(f"{{{W_NS}}}r")
        for idx, r in enumerate(runs):
            for t in r.findall(f"{{{W_NS}}}t"):
                if t.text and BANNED_UNICODE_RE.search(t.text):
                    t.text = clean_hidden_unicode(t.text)
                    fixes += 1
            # Strip trailing space immediately before a footnoteReference run
            if r.find(f"{{{W_NS}}}footnoteReference") is not None and idx > 0:
                prev_ts = runs[idx - 1].findall(f"{{{W_NS}}}t")
                if prev_ts and prev_ts[-1].text and prev_ts[-1].text.endswith(" "):
                    prev_ts[-1].text = prev_ts[-1].text.rstrip(" ")
                    fixes += 1
            # Strip leading space immediately after a footnoteReference before punctuation
            if r.find(f"{{{W_NS}}}footnoteReference") is not None and idx + 1 < len(runs):
                next_ts = runs[idx + 1].findall(f"{{{W_NS}}}t")
                if next_ts and next_ts[0].text and re.match(r"^\s+[.,;:]", next_ts[0].text):
                    next_ts[0].text = next_ts[0].text.lstrip(" ")
                    fixes += 1

    for p in doc.paragraphs:
        _clean_paragraph_xml(p._p)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _clean_paragraph_xml(p._p)

    # Also clean footnotes.xml
    for part in doc.part.package.parts:
        if str(part.partname).endswith("footnotes.xml"):
            root = etree.fromstring(part.blob)
            changed = False
            for t in root.findall(f".//{{{W_NS}}}t"):
                if t.text and BANNED_UNICODE_RE.search(t.text):
                    t.text = clean_hidden_unicode(t.text)
                    fixes += 1
                    changed = True
            if changed:
                part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    if fixes > 0:
        doc.save(str(docx_path))
    return fixes


def check_footnote_spacing_xml(doc: Document) -> list[str]:
    issues: list[str] = []

    def _scan_p(p_el: etree._Element) -> None:
        runs = p_el.findall(f"{{{W_NS}}}r")
        for idx, r in enumerate(runs):
            if r.find(f"{{{W_NS}}}footnoteReference") is None:
                continue
            prev_txt = "".join(t.text or "" for t in runs[idx - 1].findall(f"{{{W_NS}}}t")) if idx > 0 else ""
            next_txt = "".join(t.text or "" for t in runs[idx + 1].findall(f"{{{W_NS}}}t")) if idx + 1 < len(runs) else ""
            if prev_txt.endswith(" "):
                issues.append(f"Пробел перед знаком сноски: '...{prev_txt[-20:]}[FN]'")
            if prev_txt.endswith((",", ";")):
                issues.append(f"Знак препинания стоит ДО сноски вместо ПОСЛЕ: '...{prev_txt[-20:]}[FN]'")
            elif prev_txt.endswith(".") and not re.search(
                r"(?:\b(?:руб|тыс|млн|млрд|ст|др|см|пр|гг|вв)|\s[гвс]|[а-яё]\.[а-яё])\.$", prev_txt, re.I
            ):
                issues.append(f"Точка стоит ДО сноски вместо ПОСЛЕ: '...{prev_txt[-20:]}[FN]'")
            if re.match(r"^\s+[.,;:]", next_txt):
                issues.append(f"Пробел между сноской и точкой/запятой: '[FN]{next_txt[:15]}...'")

    for p in doc.paragraphs:
        _scan_p(p._p)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _scan_p(p._p)
    return issues


def extract_bibliography_entries(doc: Document) -> list[str]:
    in_bib = False
    entries: list[str] = []
    for p in doc.paragraphs:
        txt = p.text.strip()
        if not txt:
            continue
        sname = p.style.name if p.style else ""
        # Skip TOC entries (TOC 1/2/3 / Оглавление) so live-updated TOC never triggers bibliography capture!
        if sname.lower().startswith(("toc", "оглавление")):
            continue
        if not in_bib:
            if sname == "Heading 1" and re.match(
                r"^список\s+(?:использованн\w+\s+)?(?:источников|литературы)\s*$", txt, re.I
            ):
                in_bib = True
            continue
        if re.match(r"^приложение\s+", txt, re.I) or sname == "Heading 1":
            break
        entries.append(txt)
    return entries


def check_bibliography(entries: list[str], contract: dict[str, Any], fn_texts: list[str], body_text: str) -> dict[str, Any]:
    bib_cfg = contract.get("bibliography", {})
    min_src = int(bib_cfg.get("min_sources", 10))
    max_src = int(bib_cfg.get("max_sources", 100))
    min_year = int(bib_cfg.get("min_year", 2020))
    max_classic = int(bib_cfg.get("max_classic_sources", 3))
    forbid_foreign = bool(bib_cfg.get("forbid_untranslated_foreign", True))
    npa_first = bool(bib_cfg.get("npa_first", True))

    issues: list[str] = []
    warnings: list[str] = []
    n = len(entries)
    if n < min_src or n > max_src:
        issues.append(f"Количество источников {n} вне диапазона [{min_src}..{max_src}]")

    # Alphabetical order check (НПА -> А-Я with Ё=Е -> A-Z)
    keys = [ru_sort_key(e, npa_first=npa_first) for e in entries]
    order_errors: list[str] = []
    for i in range(len(keys) - 1):
        if keys[i] > keys[i + 1]:
            order_errors.append(
                f"Нарушен алфавит между #{i + 1} («{strip_bib_prefix(entries[i])[:35]}») "
                f"и #{i + 2} («{strip_bib_prefix(entries[i + 1])[:35]}»)"
            )
    if order_errors:
        issues.extend(order_errors[:4])

    # Year & foreign language checks
    old_cnt = 0
    foreign_list: list[str] = []
    for e in entries:
        clean = strip_bib_prefix(e)
        cyr = len(CYRILLIC_CHAR_RE.findall(clean))
        if forbid_foreign and cyr < 12:
            foreign_list.append(clean[:45])
        years = [int(y) for y in re.findall(r"\b(19[5-9]\d|20[0-2]\d)\b", clean)]
        # Exclude access date year when publication year exists
        pub_part = re.split(r"дата\s+обращения", clean, flags=re.I)[0]
        pub_years = [int(y) for y in re.findall(r"\b(19[5-9]\d|20[0-2]\d)\b", pub_part)] or years
        if pub_years and max(pub_years) < min_year:
            old_cnt += 1

    if foreign_list:
        issues.append(f"Найдены непереведённые иностранные источники ({len(foreign_list)}): {foreign_list[:3]}")
    if old_cnt > max_classic:
        warnings.append(f"Источников старше {min_year} г.: {old_cnt} (квота классиков: {max_classic})")

    # Basic citation coverage check
    cite_style = contract.get("citations", {}).get("style", "footnote")
    if cite_style == "brackets" and n > 0:
        cited_nums = {int(m) for m in re.findall(r"\[(\d{1,3})(?:\s*[,;]|])", body_text)}
        uncited = sorted(set(range(1, n + 1)) - cited_nums)
        if uncited:
            warnings.append(f"Источники без внутритекстовых ссылок [N]: {uncited[:10]}")
    elif cite_style == "footnote" and n > 0 and fn_texts:
        fn_blob = " ".join(fn_texts).lower().replace("ё", "е")
        uncited_surnames: list[str] = []
        for idx, e in enumerate(entries, start=1):
            clean = strip_bib_prefix(e)
            toks = re.findall(r"[а-яё]{4,}", clean.lower())
            if toks and toks[0].replace("ё", "е") not in fn_blob:
                uncited_surnames.append(f"#{idx} ({toks[0]})")
        if len(uncited_surnames) > max(2, int(n * 0.2)):
            warnings.append(f"Возможно не процитированы в сносках: {', '.join(uncited_surnames[:6])}")

    return {
        "status": "PASS" if not issues else "FAIL",
        "count": n,
        "expected_range": [min_src, max_src],
        "issues": issues,
        "warnings": warnings,
    }


def analyze_rendered_pdf(pdf_path: Path, contract: dict[str, Any]) -> dict[str, Any]:
    doc = fitz.open(str(pdf_path))
    total_pages = len(doc)
    pages_cfg = contract.get("pages", {})
    min_p = int(pages_cfg.get("min", 1))
    max_p = int(pages_cfg.get("max", 999))
    min_lines_tail = int(contract.get("layout", {}).get("min_lines_last_section_page", 4))
    include_title = bool(contract.get("layout", {}).get("include_title_page", True))
    include_toc = bool(contract.get("layout", {}).get("include_toc", True))
    forbid_tam_zhe_top = bool(contract.get("citations", {}).get("forbid_tam_zhe_page_top", True))

    page_summaries: list[dict[str, Any]] = []
    tam_zhe_page_top_pages: list[int] = []
    orphan_headings: list[str] = []
    toc_placeholder_found = False

    for pno in range(total_pages):
        page = doc[pno]
        h = page.rect.height
        d = page.get_text("dict")
        body_lines: list[tuple[float, str, float]] = []  # (y0, text, max_font_size)
        footnote_lines: list[tuple[float, str]] = []
        first_nonempty_line = ""

        for blk in d.get("blocks", []):
            if blk.get("type") != 0:
                continue
            for ln in blk.get("lines", []):
                spans = ln.get("spans", [])
                txt = "".join(sp.get("text", "") for sp in spans).strip()
                if not txt:
                    continue
                y0, y1 = ln["bbox"][1], ln["bbox"][3]
                max_sz = max((float(sp.get("size", 12.0)) for sp in spans), default=12.0)

                # Ignore isolated page number in top/bottom 8% margin
                if (y0 < h * 0.08 or y1 > h * 0.92) and re.fullmatch(r"\d{1,3}", txt):
                    continue
                # Ignore footnote separator line
                if re.fullmatch(r"[_\-—\s]+", txt):
                    continue

                if "обновите поле оглавления" in txt.lower():
                    toc_placeholder_found = True

                # Footnote area: bottom half of page with small font <= 10.8pt
                if y0 > h * 0.50 and max_sz <= 10.8:
                    footnote_lines.append((y0, txt))
                else:
                    body_lines.append((y0, txt, max_sz))
                    if not first_nonempty_line:
                        first_nonempty_line = txt

        body_lines.sort(key=lambda x: x[0])
        footnote_lines.sort(key=lambda x: x[0])

        starts_new_section = bool(SECTION_HEADING_RE.match(first_nonempty_line))

        # Check page-top "Там же." in footnotes
        if forbid_tam_zhe_top and footnote_lines:
            # Strip leading footnote digit(s) like "12 Там же." or "12Там же."
            fn_first = re.sub(r"^\d+\s*", "", footnote_lines[0][1]).strip()
            if re.match(r"^там\s+же\b", fn_first, re.I):
                tam_zhe_page_top_pages.append(pno + 1)

        # Check orphan Heading 2 sitting alone at very bottom of page (>86% height)
        if body_lines:
            last_y0, last_txt, _ = body_lines[-1]
            if last_y0 > h * 0.85 and re.match(r"^\d+\.\d+\.?\s+[А-ЯЁ]", last_txt):
                orphan_headings.append(f"Стр. {pno + 1}: «{last_txt[:60]}»")

        page_summaries.append({
            "page": pno + 1,
            "body_lines": len(body_lines),
            "footnote_lines": len(footnote_lines),
            "starts_section": starts_new_section,
            "first_line": first_nonempty_line[:70],
        })

    doc.close()

    # Detect hanging section-end pages (< min_lines_tail body lines)
    skip_pages = set()
    if include_title:
        skip_pages.add(1)
    if include_toc:
        skip_pages.add(2 if include_title else 1)

    hanging_pages: list[str] = []
    for idx, ps in enumerate(page_summaries):
        pnum = ps["page"]
        if pnum in skip_pages:
            continue
        is_last_page = (idx == total_pages - 1)
        next_starts_section = (idx + 1 < total_pages and page_summaries[idx + 1]["starts_section"])
        if (is_last_page or next_starts_section) and ps["body_lines"] < min_lines_tail:
            hanging_pages.append(
                f"Стр. {pnum}: всего {ps['body_lines']} стр. основного текста (< {min_lines_tail})"
            )

    page_count_ok = (min_p <= total_pages <= max_p)
    return {
        "total_pages": total_pages,
        "expected_pages": [min_p, max_p],
        "page_count_status": "PASS" if page_count_ok else "FAIL",
        "hanging_pages": hanging_pages,
        "orphan_headings": orphan_headings,
        "tam_zhe_page_top_pages": tam_zhe_page_top_pages,
        "toc_placeholder_found": toc_placeholder_found,
        "page_summaries": page_summaries,
    }


def main() -> int:
    ensure_utf8_io()
    parser = argparse.ArgumentParser(
        description="Comprehensive 9-point GOST DOCX, hygiene, bibliography, PDF pagination & PNG preview auditor."
    )
    parser.add_argument("docx", help="Path to output/final.docx")
    parser.add_argument("--contract", default="project_contract.json", help="Path to project_contract.json")
    parser.add_argument("--render-dir", default="output/preview", help="Directory for PDF and PNG page previews")
    parser.add_argument("--fix", action="store_true", help="Auto-fix hidden Unicode and footnote spacing before auditing")
    parser.add_argument("--dpi", type=int, default=150, help="PNG preview DPI (default: 150)")
    parser.add_argument("--json", action="store_true", help="Output machine-readable JSON report")
    args = parser.parse_args()

    docx_path = Path(args.docx).resolve()
    if not docx_path.exists():
        print(f"ERROR: DOCX file not found: {docx_path}", file=sys.stderr)
        return 2

    contract = load_contract(args.contract)
    fixed_count = auto_fix_docx(docx_path) if args.fix else 0

    doc = Document(str(docx_path))
    fn_texts = extract_footnotes_texts(doc)

    # Collect all text across paragraphs, tables, and footnotes
    para_texts = [p.text for p in doc.paragraphs]
    table_texts = [
        p.text
        for tbl in doc.tables
        for row in tbl.rows
        for cell in row.cells
        for p in cell.paragraphs
    ]
    all_text = "\n".join(para_texts + table_texts + fn_texts)
    body_text = "\n".join(para_texts)

    # 1. Placeholders (allow legitimate signature/grade/date underlines on custom university title templates)
    sig_line_re = re.compile(r"(?:оценка|подпись|дата|нормоконтрол|допущен|зачтено|проверил|«_+»)", re.I)
    placeholders: list[str] = []
    for line in all_text.splitlines():
        for pat in PLACEHOLDER_PATTERNS:
            for m in pat.finditer(line):
                tok = m.group(0)
                if re.fullmatch(r"_{3,}", tok) and sig_line_re.search(line):
                    continue
                placeholders.append(tok)
    c_placeholders = "PASS" if not placeholders else "FAIL"

    # 2. Hidden Unicode & Homoglyphs
    hidden_uni = find_hidden_unicode(all_text)
    homoglyphs = find_homoglyphs(all_text)
    c_hygiene = "PASS" if (not hidden_uni and not homoglyphs) else "FAIL"

    # 3. Footnote & citation punctuation
    fn_spacing_issues = check_footnote_spacing_xml(doc)
    c_fn_spacing = "PASS" if not fn_spacing_issues else "FAIL"

    # 4. Structure, headings, margins, fonts
    struct_issues: list[str] = []
    h1_list = [p.text.strip() for p in doc.paragraphs if p.style and p.style.name == "Heading 1" and p.text.strip()]
    h2_list = [p.text.strip() for p in doc.paragraphs if p.style and p.style.name == "Heading 2" and p.text.strip()]
    if not h1_list:
        struct_issues.append("Не найдено ни одного заголовка стиля Heading 1")
    for h in h1_list + h2_list:
        if h.endswith("."):
            struct_issues.append(f"Точка в конце заголовка (нарушение ГОСТ 7.32): «{h[:55]}»")
    if contract.get("citations", {}).get("style", "footnote") == "footnote" and len(fn_texts) == 0:
        struct_issues.append("Режим 'footnote', но в документе 0 нативных сносок Word (word/footnotes.xml)")

    sec = doc.sections[0]
    exp_m = contract["layout"]["margins_mm"]
    for side, attr in (("left", "left_margin"), ("right", "right_margin"), ("top", "top_margin"), ("bottom", "bottom_margin")):
        actual_mm = round(getattr(sec, attr).mm, 1)
        if abs(actual_mm - float(exp_m[side])) > 1.5:
            struct_issues.append(f"Поле {side}: {actual_mm} мм != ожидаемых {exp_m[side]} мм")
    c_struct = "PASS" if not struct_issues else "FAIL"

    # 5. Bibliography
    bib_entries = extract_bibliography_entries(doc)
    bib_report = check_bibliography(bib_entries, contract, fn_texts, body_text)

    # 6-9. PDF conversion, pagination, tails, Там же page-top, and PNG previews
    render_dir = Path(args.render_dir).resolve()
    render_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = render_dir / (docx_path.stem + ".pdf")
    pdf_ok, pdf_engine = export_docx_to_pdf(docx_path, pdf_path, update_toc=False)

    pdf_report: dict[str, Any] = {"ok": pdf_ok, "engine": pdf_engine}
    png_files: list[str] = []
    png_engine = ""
    c_pages = "FAIL"
    c_layout = "FAIL"
    c_tam_zhe = "PASS"

    if pdf_ok and pdf_path.exists():
        pdf_metrics = analyze_rendered_pdf(pdf_path, contract)
        pdf_report.update(pdf_metrics)
        c_pages = pdf_metrics["page_count_status"]
        layout_fails = (
            bool(pdf_metrics["hanging_pages"])
            or bool(pdf_metrics["orphan_headings"])
            or bool(pdf_metrics["toc_placeholder_found"])
        )
        c_layout = "PASS" if not layout_fails else "FAIL"
        c_tam_zhe = "PASS" if not pdf_metrics["tam_zhe_page_top_pages"] else "FAIL"

        imgs, png_engine = render_pdf_to_pngs(pdf_path, render_dir, dpi=args.dpi)
        png_files = [str(p) for p in imgs]
    else:
        pdf_report["error"] = pdf_engine

    # Direct quote, equation, figure & data-table counts
    quotes_cnt = len(re.findall(r"«[^»]{25,500}»", body_text))
    eq_tables = 0
    data_tables = 0
    for tbl in doc.tables:
        tblPr = tbl._tbl.tblPr
        desc = tblPr.find(qn("w:tblDescription")) if tblPr is not None else None
        dval = desc.get(qn("w:val"), "") if desc is not None else ""
        if dval == "GOST_EQUATION_LAYOUT":
            eq_tables += 1
        elif dval != "CUSTOM_TITLE_TEMPLATE_TABLE":
            data_tables += 1
    figures_cnt = len(doc.element.body.findall(f".//{{{W_NS}}}drawing"))

    checks = {
        "1_placeholders": c_placeholders,
        "2_unicode_and_homoglyphs": c_hygiene,
        "3_footnote_punctuation": c_fn_spacing,
        "4_gost_structure_margins": c_struct,
        "5_bibliography": bib_report["status"],
        "6_page_count": c_pages,
        "7_no_tails_or_orphans": c_layout,
        "8_no_tam_zhe_page_top": c_tam_zhe,
        "9_png_previews_rendered": "PASS" if png_files else "FAIL",
    }
    overall = "PASS" if all(v == "PASS" for v in checks.values()) else "FAIL"

    report = {
        "overall": overall,
        "docx": str(docx_path),
        "auto_fixed_count": fixed_count,
        "stats": {
            "pages": pdf_report.get("total_pages"),
            "headings_h1": len(h1_list),
            "headings_h2": len(h2_list),
            "tables": data_tables,
            "equations": eq_tables,
            "figures": figures_cnt,
            "native_footnotes": len(fn_texts),
            "direct_quotes_ge_25c": quotes_cnt,
            "bibliography_entries": bib_report["count"],
        },
        "checks": checks,
        "details": {
            "placeholders": placeholders[:10],
            "hidden_unicode": hidden_uni,
            "homoglyphs": homoglyphs[:10],
            "footnote_spacing_issues": fn_spacing_issues[:10],
            "structure_issues": struct_issues,
            "bibliography": bib_report,
            "pdf": {k: v for k, v in pdf_report.items() if k != "page_summaries"},
            "png_engine": png_engine,
            "png_files": png_files,
        },
    }

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"=== GOST DOCX AUDIT: {docx_path.name} -> {overall} ===")
        if args.fix:
            print(f"  Auto-fixed items: {fixed_count}")
        st = report["stats"]
        print(
            f"  Статистика: стр.={st['pages']}, H1={st['headings_h1']}, H2={st['headings_h2']}, "
            f"таблиц={st['tables']}, рисунков={st['figures']}, формул={st['equations']}, "
            f"сносок={st['native_footnotes']}, цитат={st['direct_quotes_ge_25c']}, "
            f"источников={st['bibliography_entries']}"
        )
        print(f"  PDF-движок: {pdf_engine} | PNG-движок: {png_engine} ({len(png_files)} PNG)")
        print("  Проверки:")
        for name, val in checks.items():
            print(f"    [{val}] {name}")
        if placeholders:
            print(f"    -> Заглушки: {placeholders[:6]}")
        if hidden_uni:
            print(f"    -> Скрытый Unicode: {hidden_uni}")
        if homoglyphs:
            print(f"    -> Гомоглифы: {[h['token'] for h in homoglyphs[:6]]}")
        for iss in fn_spacing_issues[:5]:
            print(f"    -> {iss}")
        for iss in struct_issues:
            print(f"    -> {iss}")
        for iss in bib_report["issues"]:
            print(f"    -> Библиография: {iss}")
        for w in bib_report["warnings"]:
            print(f"    -> [WARN] Библиография: {w}")
        if pdf_ok:
            if c_pages == "FAIL":
                print(f"    -> Страницы: {pdf_report['total_pages']} вне [{pdf_report['expected_pages'][0]}..{pdf_report['expected_pages'][1]}]")
            for hp in pdf_report.get("hanging_pages", []):
                print(f"    -> Висячий хвост: {hp}")
            for oh in pdf_report.get("orphan_headings", []):
                print(f"    -> Висячий заголовок: {oh}")
            if pdf_report.get("toc_placeholder_found"):
                print("    -> Оглавление не обновлено (остался текст-заглушка)")
            if pdf_report.get("tam_zhe_page_top_pages"):
                print(f"    -> 'Там же.' стоит первой сноской на страницах: {pdf_report['tam_zhe_page_top_pages']}")
        else:
            print(f"    -> Ошибка PDF-экспорта: {pdf_engine}")
        if png_files:
            print(f"  PNG превью: {render_dir} ({Path(png_files[0]).name} .. {Path(png_files[-1]).name})")

    return 0 if overall == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
