#!/usr/bin/env python3
"""Shared utilities for the ru-academic-gost pipeline."""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any


def ensure_utf8_io() -> None:
    """Force UTF-8 stdout/stderr on Windows so Cyrillic & guillemets never crash."""
    for stream in (sys.stdout, sys.stderr):
        try:
            if hasattr(stream, "reconfigure"):
                stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


# Invisible / ChatGPT artifact characters banned in GOST academic DOCX
BANNED_UNICODE: dict[str, str] = {
    "\u200b": "ZERO WIDTH SPACE (U+200B)",
    "\u200c": "ZERO WIDTH NON-JOINER (U+200C)",
    "\u200d": "ZERO WIDTH JOINER (U+200D)",
    "\u2060": "WORD JOINER (U+2060)",
    "\ufeff": "BOM / ZWNBSP (U+FEFF)",
    "\u00ad": "SOFT HYPHEN (U+00AD)",
    "\u2011": "NON-BREAKING HYPHEN (U+2011)",
}

BANNED_UNICODE_RE = re.compile("[" + "".join(re.escape(k) for k in BANNED_UNICODE) + "]")

# Latin & Greek lookalike characters that can hide inside Cyrillic words
LOOKALIKE_LATIN_GREEK = set("ABCEHKMOPTXYaeopcyxIiJjKkMmNnOoPpTtXx\u0391\u0392\u0395\u0397\u039a\u039c\u039d\u039f\u03a1\u03a4\u03a7\u03b1\u03b5\u03bf\u03c1\u03c5\u03c7")
CYRILLIC_CHAR_RE = re.compile(r"[А-Яа-яЁё]")
LATIN_GREEK_CHAR_RE = re.compile(r"[A-Za-z\u0370-\u03ff]")

# Word tokens split on whitespace/punctuation (excluding hyphens so IT-рынок splits cleanly)
_TOKEN_RE = re.compile(r"[А-Яа-яЁёA-Za-z\u0370-\u03ff]+")


def find_homoglyphs(text: str) -> list[dict[str, Any]]:
    """Detect Latin/Greek lookalike letters embedded inside Cyrillic words.

    Tokens split on hyphens/punctuation, so legitimate compounds like
    'PR-кампания', 'IT-сфера', 'B2B-рынок', 'Wi-Fi' are NOT false positives.
    """
    findings: list[dict[str, Any]] = []
    for m in _TOKEN_RE.finditer(text):
        tok = m.group(0)
        has_cyr = bool(CYRILLIC_CHAR_RE.search(tok))
        has_other = bool(LATIN_GREEK_CHAR_RE.search(tok))
        if has_cyr and has_other:
            bad_chars = sorted({ch for ch in tok if LATIN_GREEK_CHAR_RE.match(ch)})
            findings.append({
                "token": tok,
                "bad_chars": bad_chars,
                "pos": m.start(),
            })
    return findings


def find_hidden_unicode(text: str) -> list[dict[str, Any]]:
    """Return list of banned invisible/artifact Unicode occurrences."""
    counts: dict[str, int] = {}
    for ch in BANNED_UNICODE_RE.findall(text):
        counts[ch] = counts.get(ch, 0) + 1
    return [
        {"char": f"U+{ord(ch):04X}", "name": BANNED_UNICODE[ch], "count": cnt}
        for ch, cnt in counts.items()
    ]


def clean_hidden_unicode(text: str) -> str:
    """Strip invisible Unicode characters and replace U+2011 with ASCII hyphen."""
    text = text.replace("\u2011", "-")
    for ch in ("\u200b", "\u200c", "\u200d", "\u2060", "\ufeff", "\u00ad"):
        text = text.replace(ch, "")
    return text


DEFAULT_CONTRACT: dict[str, Any] = {
    "work_type": "kursovaya",
    "autonomy_mode": "standard",
    "topic": "",
    "discipline": "",
    "pages": {"min": 20, "max": 25, "target": 23},
    "oral_speech": {"enabled": False, "duration_minutes": 5, "words_per_minute": 125},
    "citations": {
        "style": "footnote",
        "forbid_tam_zhe_page_top": True,
        "min_quotes_per_theory_page": 1,
        "target_quote_share_pct": [10, 20],
    },
    "bibliography": {
        "min_sources": 20,
        "max_sources": 25,
        "min_year": 2020,
        "max_classic_sources": 3,
        "forbid_untranslated_foreign": True,
        "npa_first": True,
        "sort_order": "alphabetical_ru",
    },
    "layout": {
        "font_name": "Times New Roman",
        "font_size_pt": 14,
        "line_spacing": 1.5,
        "first_line_indent_cm": 1.25,
        "margins_mm": {"left": 30, "right": 15, "top": 20, "bottom": 20},
        "footnote_font_size_pt": 10,
        "table_font_size_pt": 12,
        "include_title_page": True,
        "title_template_path": "",
        "include_toc": True,
        "page_number_position": "bottom_center",
        "min_lines_last_section_page": 4,
        "short_table_max_rows": 5,
    },
    "conclusion_style": "analytical_summary",
    "student_voice_sample": "",
    "special_instructions": [],
}


def deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    out = dict(base)
    for k, v in override.items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def load_contract(path: str | Path | None) -> dict[str, Any]:
    if not path:
        return json.loads(json.dumps(DEFAULT_CONTRACT))
    p = Path(path)
    if not p.exists():
        return json.loads(json.dumps(DEFAULT_CONTRACT))
    raw = json.loads(p.read_text(encoding="utf-8"))
    # Backwards-compat for flat right_margin_mm / min_pages / max_pages keys
    if "right_margin_mm" in raw:
        raw.setdefault("layout", {}).setdefault("margins_mm", {})["right"] = raw["right_margin_mm"]
    if "min_pages" in raw:
        raw.setdefault("pages", {})["min"] = raw["min_pages"]
    if "max_pages" in raw:
        raw.setdefault("pages", {})["max"] = raw["max_pages"]
    return deep_merge(DEFAULT_CONTRACT, raw)


# --- Russian GOST bibliography collation (НПА -> А–Я with Ё=Е -> A–Z) ---

_RU_ALPHABET = "абвгдежзийклмнопрстуфхцчшщъыьэюя"
_RU_RANK = {ch: i for i, ch in enumerate(_RU_ALPHABET)}

_NPA_PATTERNS: list[tuple[int, re.Pattern[str]]] = [
    (0, re.compile(r"^конституция\s+российской\s+федерации", re.I)),
    (1, re.compile(r"^(федеральный\s+конституционный\s+закон|гражданский\s+кодекс|налоговый\s+кодекс|трудовой\s+кодекс|уголовный\s+кодекс|бюджетный\s+кодекс|кодекс\s+российской)", re.I)),
    (2, re.compile(r"^(федеральный\s+закон|закон\s+рф|о\s+.*\s+федеральный\s+закон)", re.I)),
    (3, re.compile(r"^указ\s+президента", re.I)),
    (4, re.compile(r"^(постановление\s+правительства|распоряжение\s+правительства)", re.I)),
    (5, re.compile(r"^(приказ|письмо|положение|указание|гост\s|санпин)", re.I)),
]


def detect_npa_rank(entry: str) -> int | None:
    s = re.sub(r"^\s*\d+\.\s*", "", entry).lstrip("«\"' ").strip()
    for rank, pat in _NPA_PATTERNS:
        if pat.search(s):
            return rank
    return None


def strip_bib_prefix(entry: str) -> str:
    s = re.sub(r"^\s*\d+[\.\)]\s*", "", entry).strip()
    s = s.lstrip("«\"'„“ ")
    return s


def ru_sort_key(entry: str, npa_first: bool = True) -> tuple:
    """GOST R 7.0.100 collation key:
    Group 0: NPA (by legal hierarchy) if npa_first
    Group 1: Cyrillic A-Я (with Ё mapped to Е; single-author before co-authored)
    Group 2: Latin A-Z
    """
    clean = strip_bib_prefix(entry)
    if npa_first:
        npa_rank = detect_npa_rank(clean)
        if npa_rank is not None:
            return (0, npa_rank, _norm_cyr_tuple(clean))

    first_alpha = ""
    for ch in clean:
        if ch.isalpha():
            first_alpha = ch
            break
    is_cyr = bool(CYRILLIC_CHAR_RE.match(first_alpha))
    group = 1 if is_cyr else 2

    # Detect single-author vs co-authored ("Иванов И. И." vs "Иванов И. И., Петров П. П.")
    author_m = re.match(
        r"^([А-ЯЁ][а-яё\-]+(?:\s+[А-ЯЁ]\.\s*(?:[А-ЯЁ]\.)?)?)\s*(,)?",
        clean,
    )
    if is_cyr and author_m:
        primary_author = author_m.group(1)
        is_coauthored = 1 if author_m.group(2) == "," else 0
        rest = clean[author_m.end():].strip()
        return (group, _norm_cyr_tuple(primary_author), is_coauthored, _norm_cyr_tuple(rest))

    return (group, _norm_cyr_tuple(clean), 0, ())


def _norm_cyr_tuple(text: str) -> tuple[int, ...]:
    s = text.lower().replace("ё", "е")
    ranks: list[int] = []
    for ch in s:
        if ch in _RU_RANK:
            ranks.append(100 + _RU_RANK[ch])
        elif "a" <= ch <= "z":
            ranks.append(200 + ord(ch) - ord("a"))
        elif "0" <= ch <= "9":
            ranks.append(50 + ord(ch) - ord("0"))
        elif ch in (" ", "-", "."):
            ranks.append(10)
    return tuple(ranks)


# --- External tool discovery & PDF / PNG rendering ---


def find_pandoc() -> str | None:
    """Locate pandoc via PATH, standard paths, or pypandoc_binary."""
    found = shutil.which("pandoc")
    if found:
        return found
    for candidate in (
        r"C:\Program Files\Pandoc\pandoc.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Pandoc\pandoc.exe"),
    ):
        if os.path.isfile(candidate):
            return candidate
    try:
        import pypandoc  # type: ignore

        p = pypandoc.get_pandoc_path()
        for cand in (p, f"{p}.exe"):
            if cand and os.path.isfile(cand):
                return cand
    except Exception:
        pass
    return None


def find_soffice() -> str | None:
    """Locate LibreOffice soffice binary."""
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for candidate in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ):
        if os.path.isfile(candidate):
            return candidate
    return None


def word_com_available() -> bool:
    """Check whether Windows pywin32 + Microsoft Word COM is usable."""
    if sys.platform != "win32":
        return False
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError:
        return False
    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        _ = word.Version
        return True
    except Exception:
        return False
    finally:
        if word is not None:
            try:
                word.Quit(False)
            except Exception:
                pass
        pythoncom.CoUninitialize()


def update_word_toc_and_export_pdf(
    docx_path: str | Path,
    pdf_path: str | Path | None = None,
    update_toc: bool = True,
) -> tuple[bool, str]:
    """Use Microsoft Word COM (Windows) to update TOC/fields and optionally export PDF.

    Returns (success, engine_or_error_message).
    """
    if sys.platform != "win32":
        return False, "Word COM is only available on Windows"
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError:
        return False, "pywin32 not installed"

    docx_abs = str(Path(docx_path).resolve())
    pdf_abs = str(Path(pdf_path).resolve()) if pdf_path else None
    if pdf_abs:
        Path(pdf_abs).parent.mkdir(parents=True, exist_ok=True)

    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.Options.CheckSpellingAsYouType = False
        word.Options.CheckGrammarAsYouType = False

        doc = word.Documents.Open(
            docx_abs,
            ConfirmConversions=False,
            ReadOnly=not update_toc,
            AddToRecentFiles=False,
        )
        if update_toc:
            # Reset first-line indent on built-in TOC 1/2/3 styles (language-independent WdBuiltinStyle -20/-21/-22)
            # so TOC entries never inherit Normal's 1.25cm red line on Russian or English Word!
            for style_id, left_pt in ((-20, 0.0), (-21, 17.0), (-22, 34.0)):
                try:
                    st = doc.Styles(style_id)
                    st.ParagraphFormat.FirstLineIndent = 0.0
                    st.ParagraphFormat.LeftIndent = left_pt
                    st.ParagraphFormat.LineSpacingRule = 0  # single/clean
                    st.ParagraphFormat.SpaceBefore = 2.0
                    st.ParagraphFormat.SpaceAfter = 2.0
                    st.Font.Name = "Times New Roman"
                    st.Font.Size = 14.0
                except Exception:
                    pass
            try:
                for i in range(1, doc.TablesOfContents.Count + 1):
                    doc.TablesOfContents(i).Update()
            except Exception:
                pass
            try:
                doc.Fields.Update()
            except Exception:
                pass
            try:
                for sec in doc.Sections:
                    sec.Footers(1).Range.Fields.Update()
                    sec.Headers(1).Range.Fields.Update()
            except Exception:
                pass
            doc.Repaginate()
            doc.Save()

        if pdf_abs:
            # wdExportFormatPDF = 17
            doc.ExportAsFixedFormat(
                OutputFileName=pdf_abs,
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=0,
                CreateBookmarks=1,
            )
        return True, f"Microsoft Word {word.Version}"
    except Exception as exc:
        return False, f"Word COM error: {exc}"
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit(False)
            except Exception:
                pass
        pythoncom.CoUninitialize()


def export_docx_to_pdf(
    docx_path: str | Path,
    pdf_path: str | Path,
    update_toc: bool = False,
) -> tuple[bool, str]:
    """Convert DOCX to PDF using Microsoft Word COM first, falling back to LibreOffice."""
    docx_p = Path(docx_path).resolve()
    pdf_p = Path(pdf_path).resolve()
    pdf_p.parent.mkdir(parents=True, exist_ok=True)

    if sys.platform == "win32":
        ok, info = update_word_toc_and_export_pdf(docx_p, pdf_p, update_toc=update_toc)
        if ok and pdf_p.exists():
            return True, info

    soffice = find_soffice()
    if soffice:
        try:
            cmd = [
                soffice,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(pdf_p.parent),
                str(docx_p),
            ]
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
            expected = pdf_p.parent / (docx_p.stem + ".pdf")
            if res.returncode == 0 and expected.exists():
                if expected != pdf_p:
                    shutil.move(str(expected), str(pdf_p))
                return True, f"LibreOffice ({soffice})"
            return False, f"LibreOffice failed (rc={res.returncode}): {res.stderr.strip()}"
        except Exception as exc:
            return False, f"LibreOffice error: {exc}"

    return False, "Neither Microsoft Word COM nor LibreOffice (soffice) is available"


def render_pdf_to_pngs(
    pdf_path: str | Path,
    out_dir: str | Path,
    dpi: int = 150,
) -> tuple[list[Path], str]:
    """Render all PDF pages to PNGs via PyMuPDF (fitz), falling back to pdftoppm."""
    pdf_p = Path(pdf_path).resolve()
    out_p = Path(out_dir).resolve()
    out_p.mkdir(parents=True, exist_ok=True)

    # Clear stale page_*.png from previous audit runs
    for stale in out_p.glob("page_*.png"):
        try:
            stale.unlink()
        except Exception:
            pass

    try:
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz  # type: ignore

        doc = fitz.open(str(pdf_p))
        pages: list[Path] = []
        for idx, page in enumerate(doc, start=1):
            pix = page.get_pixmap(dpi=dpi)
            img_path = out_p / f"page_{idx:02d}.png"
            pix.save(str(img_path))
            pages.append(img_path)
        doc.close()
        return pages, "PyMuPDF (fitz)"
    except Exception as fitz_err:
        pdftoppm = shutil.which("pdftoppm")
        if pdftoppm:
            prefix = out_p / "page"
            res = subprocess.run(
                [pdftoppm, "-png", "-r", str(dpi), str(pdf_p), str(prefix)],
                capture_output=True,
                text=True,
                timeout=120,
            )
            if res.returncode == 0:
                imgs = sorted(out_p.glob("page-*.png"))
                renamed: list[Path] = []
                for idx, img in enumerate(imgs, start=1):
                    target = out_p / f"page_{idx:02d}.png"
                    shutil.move(str(img), str(target))
                    renamed.append(target)
                return renamed, "pdftoppm"
        return [], f"PNG render failed: {fitz_err}"
