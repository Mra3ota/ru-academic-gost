#!/usr/bin/env python3
"""Parse, deduplicate, index, and search downloaded PDF/TXT sources for ru-academic-gost."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path
from typing import Any

try:
    import pymupdf as fitz
except ImportError:
    import fitz  # type: ignore

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _common import (
    CYRILLIC_CHAR_RE,
    LATIN_GREEK_CHAR_RE,
    clean_hidden_unicode,
    ensure_utf8_io,
    load_contract,
    ru_sort_key,
)


def dehyphenate_ru(text: str) -> str:
    """Rejoin Russian/Latin words broken across lines with soft/ASCII hyphens."""
    text = text.replace("\u00ad", "")
    # Join word-\nword when next line starts with lowercase Cyrillic/Latin
    text = re.sub(r"([А-Яа-яЁёA-Za-z]{2,})-\s*\n\s*([а-яёa-z]{2,})", r"\1\2", text)
    # Collapse excessive blank lines
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def extract_page_reading_order(page: fitz.Page) -> tuple[str, int | None]:
    """Extract text in column-aware reading order and attempt to detect printed journal page number."""
    rect = page.rect
    blocks = page.get_text("blocks")  # (x0, y0, x1, y1, text, block_no, block_type)
    text_blocks = [b for b in blocks if b[6] == 0 and (b[4] or "").strip()]
    if not text_blocks:
        return "", None

    # Detect printed page number from top 9% or bottom 9% isolated numeric block
    printed_page: int | None = None
    for b in text_blocks:
        x0, y0, x1, y1, txt = b[0], b[1], b[2], b[3], b[4].strip()
        if y0 < rect.height * 0.10 or y1 > rect.height * 0.90:
            m = re.fullmatch(r"[-–—\s]*(\d{1,4})[-–—\s]*", txt)
            if m:
                val = int(m.group(1))
                if 1 <= val <= 2500:
                    printed_page = val
                    break

    # Detect two-column layout: significant text strictly in left half AND strictly in right half
    mid = rect.width * 0.50
    left_narrow = [b for b in text_blocks if b[2] <= mid + rect.width * 0.04 and (b[2] - b[0]) < rect.width * 0.55]
    right_narrow = [b for b in text_blocks if b[0] >= mid - rect.width * 0.04 and (b[2] - b[0]) < rect.width * 0.55]
    is_two_col = len(left_narrow) >= 3 and len(right_narrow) >= 3

    if is_two_col:
        header_blocks = [b for b in text_blocks if b[1] < rect.height * 0.12 and (b[2] - b[0]) > rect.width * 0.55]
        footer_blocks = [b for b in text_blocks if b[3] > rect.height * 0.90 and (b[2] - b[0]) > rect.width * 0.55]
        used = set(id(b) for b in header_blocks + footer_blocks)
        body = [b for b in text_blocks if id(b) not in used]
        left_col = sorted([b for b in body if (b[0] + b[2]) / 2.0 < mid], key=lambda b: (b[1], b[0]))
        right_col = sorted([b for b in body if (b[0] + b[2]) / 2.0 >= mid], key=lambda b: (b[1], b[0]))
        ordered = (
            sorted(header_blocks, key=lambda b: b[1])
            + left_col
            + right_col
            + sorted(footer_blocks, key=lambda b: b[1])
        )
    else:
        ordered = sorted(text_blocks, key=lambda b: (round(b[1] / 6.0), b[0]))

    raw = "\n".join(b[4].strip() for b in ordered)
    return dehyphenate_ru(clean_hidden_unicode(raw)), printed_page


def shingle_set(text: str, k: int = 5) -> set[str]:
    words = re.findall(r"[а-яёa-z0-9]{3,}", text.lower())
    if len(words) < k:
        return set(words)
    # Sample at most 1500 shingles for speed
    step = max(1, (len(words) - k) // 1500)
    return {" ".join(words[i : i + k]) for i in range(0, len(words) - k + 1, step)}


def jaccard(a: set[str], b: set[str]) -> float:
    if not a or not b:
        return 0.0
    inter = len(a & b)
    union = len(a | b)
    return inter / union if union else 0.0


def detect_language(text: str) -> tuple[str, float]:
    cyr = len(CYRILLIC_CHAR_RE.findall(text))
    lat = len(LATIN_GREEK_CHAR_RE.findall(text))
    total = cyr + lat
    if total == 0:
        return "unknown", 0.0
    ratio = cyr / total
    if ratio >= 0.65:
        return "ru", round(ratio, 3)
    if ratio <= 0.20:
        return "foreign", round(ratio, 3)
    return "mixed", round(ratio, 3)


def count_empirical_signals(text: str) -> int:
    pats = (
        r"\d+(?:[.,]\d+)?\s*%",
        r"\b\d+(?:[.,]\d+)?\s*(?:млн|млрд|тыс\.|руб\.|₽)\b",
        r"\b(?:n|N)\s*=\s*\d+",
        r"\b(?:таблица|рисунок|график|выборка|опрос|исследование|кейс)\b",
    )
    return sum(len(re.findall(p, text, re.I)) for p in pats)


def infer_year(meta: dict[str, Any], first_pages_text: str, filename: str) -> int | None:
    if meta.get("year"):
        try:
            y = int(str(meta["year"])[:4])
            if 1950 <= y <= 2030:
                return y
        except ValueError:
            pass
    # Prefer publication years near top of first page (skip bibliography/copyright ranges carefully)
    head = first_pages_text[:2500]
    years = [int(y) for y in re.findall(r"\b(19[6-9]\d|20[0-2]\d)\b", head)]
    if years:
        # Most recent year in header/metadata block (capped at current/next year)
        valid = [y for y in years if 1960 <= y <= 2027]
        if valid:
            return max(valid)
    fn_years = [int(y) for y in re.findall(r"\b(19[6-9]\d|20[0-2]\d)\b", filename)]
    return max(fn_years) if fn_years else None


def infer_title_and_author(meta: dict[str, Any], text: str, filename: str) -> tuple[str, list[str]]:
    title = (meta.get("title") or "").strip()
    authors = meta.get("authors") or []
    if isinstance(authors, str):
        authors = [a.strip() for a in authors.split(";") if a.strip()]

    lines = [ln.strip() for ln in text[:1800].splitlines() if len(ln.strip()) >= 4]
    if not title and lines:
        for ln in lines[:10]:
            if not re.match(r"^(УДК|ББК|DOI|ISSN|URL|TITLE|AUTHORS|SITE|PUBLISHED|YEAR|ACCESSED)", ln, re.I):
                if 15 <= len(ln) <= 220:
                    title = ln
                    break
    if not title:
        title = Path(filename).stem.replace("_", " ")

    if not authors:
        # Heuristic: "Иванов И. И." or "И. И. Иванов" in top lines
        found_auth: list[str] = []
        for ln in lines[:12]:
            for m in re.finditer(r"\b([А-ЯЁ][а-яё\-]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.|[А-ЯЁ]\.\s*[А-ЯЁ]\.\s*[А-ЯЁ][а-яё\-]+)\b", ln):
                cand = m.group(1).strip()
                if cand not in found_auth:
                    found_auth.append(cand)
            if len(found_auth) >= 3:
                break
        authors = found_auth
    return title, authors


def build_gost_draft(rec: dict[str, Any]) -> str:
    authors = rec.get("authors") or []
    auth_str = ", ".join(authors[:3])
    if len(authors) > 3:
        auth_str += " [и др.]"
    title = rec.get("title") or rec["filename"]
    journal = rec.get("journal") or rec.get("site_name") or ""
    year = rec.get("year") or "б.г."
    url = rec.get("url") or ""
    accessed = rec.get("accessed_date") or ""
    pages_range = ""
    if rec.get("firstpage") and rec.get("lastpage"):
        pages_range = f" — С. {rec['firstpage']}–{rec['lastpage']}."

    prefix = f"{auth_str} " if auth_str else ""
    if rec["source_kind"] == "web":
        site_part = f" // {journal} : [сайт]" if journal else ""
        url_part = f" — URL: {url}" if url else ""
        acc_part = f" (дата обращения: {accessed})." if accessed else "."
        return f"{prefix}{title} [Электронный ресурс]{site_part}. — {year}.{url_part}{acc_part}"
    else:
        j_part = f" // {journal}." if journal else "."
        vol_issue = ""
        if rec.get("volume"):
            vol_issue += f" — Т. {rec['volume']}."
        if rec.get("issue"):
            vol_issue += f" — № {rec['issue']}."
        return f"{prefix}{title}{j_part} — {year}.{vol_issue}{pages_range}"


def load_sidecar(path: Path) -> dict[str, Any]:
    for cand in (
        path.with_suffix(path.suffix + ".meta.json"),
        path.with_suffix(".meta.json"),
    ):
        if cand.exists():
            try:
                return json.loads(cand.read_text(encoding="utf-8"))
            except Exception:
                pass
    return {}


def parse_txt_file(path: Path, sidecar: dict[str, Any]) -> tuple[list[tuple[int, int | None, str]], dict[str, Any]]:
    raw = clean_hidden_unicode(path.read_text(encoding="utf-8", errors="replace"))
    meta = dict(sidecar)
    body = raw
    if "---" in raw[:1200]:
        head, rest = raw.split("---", 1)
        for ln in head.splitlines():
            if ":" in ln:
                k, v = ln.split(":", 1)
                kl = k.strip().lower()
                val = v.strip()
                if kl == "url" and val and not meta.get("url"):
                    meta["url"] = val
                elif kl == "title" and val and not meta.get("title"):
                    meta["title"] = val
                elif kl == "authors" and val and not meta.get("authors"):
                    meta["authors"] = [a.strip() for a in val.split(";") if a.strip()]
                elif kl in ("site", "journal") and val and not meta.get("journal"):
                    meta["journal"] = val
                elif kl == "year" and val and not meta.get("year"):
                    meta["year"] = val
                elif kl == "accessed" and val and not meta.get("accessed_date"):
                    meta["accessed_date"] = val
        body = rest.strip()
    return [(1, None, body)], meta


def parse_pdf_file(
    path: Path,
    sidecar: dict[str, Any],
    keywords: list[str],
    large_threshold: int = 80,
) -> tuple[list[tuple[int, int | None, str]], dict[str, Any], int]:
    meta = dict(sidecar)
    doc = fitz.open(str(path))
    total_pages = len(doc)

    # Pull embedded PDF metadata if missing
    pdf_info = doc.metadata or {}
    if not meta.get("title") and pdf_info.get("title") and len(pdf_info["title"].strip()) > 6:
        meta["title"] = pdf_info["title"].strip()
    if not meta.get("authors") and pdf_info.get("author"):
        meta["authors"] = [a.strip() for a in re.split(r"[;,]", pdf_info["author"]) if a.strip()]

    first_printed: int | None = None
    if meta.get("firstpage") and str(meta["firstpage"]).isdigit():
        first_printed = int(str(meta["firstpage"]))

    selected_indices: list[int]
    if total_pages <= large_threshold:
        selected_indices = list(range(total_pages))
    else:
        # Large book/monograph: first 8 pages (title + TOC + intro) + keyword-matching pages
        keep = set(range(min(8, total_pages)))
        kw_low = [k.lower() for k in keywords if k.strip()]
        scored: list[tuple[int, int]] = []
        for i in range(8, total_pages):
            txt = (doc[i].get_text("text") or "").lower()
            if "содержание" in txt or "оглавление" in txt:
                keep.add(i)
                continue
            if kw_low:
                hits = sum(txt.count(k) for k in kw_low)
                if hits > 0:
                    scored.append((hits, i))
        scored.sort(reverse=True)
        for _, idx in scored[:35]:
            keep.add(idx)
        selected_indices = sorted(keep)

    pages_out: list[tuple[int, int | None, str]] = []
    for idx in selected_indices:
        txt, detected_jpage = extract_page_reading_order(doc[idx])
        jpage = detected_jpage
        if jpage is None and first_printed is not None:
            jpage = first_printed + idx
        if txt:
            pages_out.append((idx + 1, jpage, txt))

    doc.close()
    return pages_out, meta, total_pages


def run_search(out_dir: Path, query: str, max_hits: int = 25) -> int:
    extracted_dir = out_dir / "extracted"
    if not extracted_dir.exists():
        print(f"ERROR: {extracted_dir} not found. Run parse_pdfs.py on sources/ first.", file=sys.stderr)
        return 2
    terms = [t.lower() for t in re.findall(r"[а-яёa-z0-9]{2,}", query)]
    if not terms:
        print("ERROR: empty search query.", file=sys.stderr)
        return 2

    hits: list[tuple[int, str, str, str]] = []
    for fpath in sorted(extracted_dir.glob("*.txt")):
        raw = fpath.read_text(encoding="utf-8", errors="replace")
        chunks = re.split(r"(=== ФАЙЛ:[^\n]+===)", raw)
        current_marker = fpath.name
        for chunk in chunks:
            if chunk.startswith("=== ФАЙЛ:"):
                current_marker = chunk.strip()
                continue
            for para in chunk.split("\n"):
                p = para.strip()
                if len(p) < 25:
                    continue
                low = p.lower()
                score = sum(1 for t in terms if t in low)
                if score > 0 and (len(terms) == 1 or score >= min(2, len(terms))):
                    hits.append((score, fpath.name, current_marker, p[:420]))

    hits.sort(key=lambda x: x[0], reverse=True)
    print(f"=== SEARCH: '{query}' ({len(hits)} matches, showing top {min(len(hits), max_hits)}) ===")
    for score, fname, marker, snippet in hits[:max_hits]:
        print(f"\n[{score}/{len(terms)}] {marker}")
        print(f"  {snippet}")
    return 0


def evaluate_exit_gate_1(records: list[dict[str, Any]], contract: dict[str, Any]) -> tuple[bool, list[str]]:
    bib = contract.get("bibliography", {})
    min_src = int(bib.get("min_sources", 10))
    min_year = int(bib.get("min_year", 2020))
    max_classic = int(bib.get("max_classic_sources", 3))
    forbid_foreign = bool(bib.get("forbid_untranslated_foreign", True))

    issues: list[str] = []
    valid = [r for r in records if not r.get("is_duplicate") and r.get("quality_score", 0) >= 35]
    if len(valid) < min_src:
        issues.append(f"Недостаточно валидных источников: {len(valid)} < {min_src} (нужно ещё {min_src - len(valid)}).")

    foreign = [r for r in valid if r.get("language") == "foreign"]
    if forbid_foreign and foreign:
        issues.append(
            f"Найдено {len(foreign)} непереведённых иностранных источников (запрещено контрактом): "
            + ", ".join(r["filename"] for r in foreign[:5])
        )

    old_sources = [r for r in valid if r.get("year") is not None and int(r["year"]) < min_year]
    if len(old_sources) > max_classic:
        issues.append(
            f"Слишком много источников старше {min_year} г.: {len(old_sources)} > квоты классиков {max_classic}."
        )

    if contract.get("work_type") in ("kursovaya", "vkr"):
        empirical = [r for r in valid if r.get("empirical_signals", 0) >= 5]
        if len(empirical) < 2:
            issues.append(
                f"Мало источников с твёрдой статистикой/цифрами для 2-й главы ({len(empirical)} < 2)."
            )

    return len(issues) == 0, issues


def main() -> int:
    ensure_utf8_io()
    parser = argparse.ArgumentParser(
        description="Index PDF/TXT sources into extracted files, catalog, and JSON index; or search them."
    )
    parser.add_argument("sources_dir", nargs="?", default="sources", help="Directory with downloaded PDF/TXT files")
    parser.add_argument("--out-dir", default="drafts", help="Output directory (default: drafts/)")
    parser.add_argument("--out", default="", help="Optional legacy path for combined sources_clean.txt")
    parser.add_argument("--index", default="", help="Optional explicit path for sources_index.json")
    parser.add_argument("--contract", default="", help="Path to project_contract.json for Exit Gate 1 check")
    parser.add_argument("--keywords", nargs="*", default=[], help="Keywords for slicing large (>80 pp.) PDFs")
    parser.add_argument("--combined", action="store_true", help="Also write combined drafts/sources_clean.txt")
    parser.add_argument("--search", default="", help="Search phrase across drafts/extracted/*.txt and exit")
    args = parser.parse_args()

    out_dir = Path(args.out_dir)
    if args.search:
        return run_search(out_dir, args.search)

    src_dir = Path(args.sources_dir)
    if not src_dir.exists():
        print(f"ERROR: sources directory not found: {src_dir}", file=sys.stderr)
        return 2

    extracted_dir = out_dir / "extracted"
    extracted_dir.mkdir(parents=True, exist_ok=True)
    # Clean previous extracted files to avoid stale sources
    for stale in extracted_dir.glob("*.txt"):
        stale.unlink()

    files = sorted(
        [p for p in src_dir.iterdir() if p.is_file() and p.suffix.lower() in (".pdf", ".txt")],
        key=lambda p: p.name.lower(),
    )
    if not files:
        print(f"ERROR: no .pdf or .txt files found in {src_dir}", file=sys.stderr)
        return 1

    seen_md5: dict[str, str] = {}
    seen_shingles: list[tuple[str, set[str]]] = []
    records: list[dict[str, Any]] = []
    combined_chunks: list[str] = []

    for idx, fpath in enumerate(files, start=1):
        md5 = hashlib.md5(fpath.read_bytes()).hexdigest()
        sidecar = load_sidecar(fpath)
        if md5 in seen_md5:
            print(f"[DUP-MD5] {fpath.name} == {seen_md5[md5]} (skipped)")
            continue
        seen_md5[md5] = fpath.name

        try:
            if fpath.suffix.lower() == ".pdf":
                pages, meta, total_pages = parse_pdf_file(fpath, sidecar, args.keywords)
                source_kind = "pdf"
            else:
                pages, meta = parse_txt_file(fpath, sidecar)
                total_pages = 1
                source_kind = "web"
        except Exception as exc:
            print(f"[ERROR] {fpath.name}: {exc}")
            continue

        full_text = "\n\n".join(txt for _, _, txt in pages)
        char_count = len(full_text)
        if char_count < 300:
            print(f"[SKIP-EMPTY] {fpath.name}: only {char_count} chars (likely scan)")
            continue

        sh = shingle_set(full_text)
        dup_of = None
        for prev_id, prev_sh in seen_shingles:
            if jaccard(sh, prev_sh) >= 0.78:
                dup_of = prev_id
                break
        if dup_of:
            print(f"[DUP-TEXT] {fpath.name} ~= {dup_of} (skipped)")
            continue

        src_id = f"src_{len(records) + 1:02d}"
        seen_shingles.append((src_id, sh))

        lang, cyr_ratio = detect_language(full_text)
        year = infer_year(meta, full_text, fpath.name)
        title, authors = infer_title_and_author(meta, full_text, fpath.name)
        emp_signals = count_empirical_signals(full_text)

        # Quality score 0..100
        q = 40
        if char_count >= 5000:
            q += 25
        elif char_count >= 2000:
            q += 15
        if authors:
            q += 10
        if year and year >= 2020:
            q += 15
        elif year:
            q += 5
        if emp_signals >= 5:
            q += 10
        quality_score = min(100, q)

        rec: dict[str, Any] = {
            "id": src_id,
            "filename": fpath.name,
            "extracted_file": f"extracted/{src_id}.txt",
            "source_kind": source_kind,
            "title": title,
            "authors": authors,
            "year": year,
            "journal": meta.get("journal", ""),
            "volume": meta.get("volume", ""),
            "issue": meta.get("issue", ""),
            "firstpage": meta.get("firstpage", ""),
            "lastpage": meta.get("lastpage", ""),
            "url": meta.get("url", ""),
            "site_name": meta.get("site_name", ""),
            "accessed_date": meta.get("accessed_date", ""),
            "pages_total": total_pages,
            "pages_extracted": len(pages),
            "char_count": char_count,
            "language": lang,
            "cyrillic_ratio": cyr_ratio,
            "empirical_signals": emp_signals,
            "quality_score": quality_score,
            "is_duplicate": False,
        }
        rec["gost_draft"] = build_gost_draft(rec)
        records.append(rec)

        # Write per-source extracted text with explicit PDF + Journal page markers
        per_blocks: list[str] = []
        for pdf_p, j_p, txt in pages:
            j_part = f" | СТР. ЖУРНАЛА: {j_p}" if j_p is not None else ""
            marker = f"=== ФАЙЛ: {fpath.name} ({src_id}) | СТР. PDF: {pdf_p}{j_part} ==="
            per_blocks.append(f"{marker}\n{txt}")
        ext_content = "\n\n".join(per_blocks) + "\n"
        (extracted_dir / f"{src_id}.txt").write_text(ext_content, encoding="utf-8")
        if args.combined or args.out:
            combined_chunks.append(ext_content)

    # Sort records alphabetically by GOST draft for easy Phase 2 planning
    records_sorted = sorted(records, key=lambda r: ru_sort_key(r["gost_draft"]))

    index_path = Path(args.index) if args.index else (out_dir / "sources_index.json")
    index_path.parent.mkdir(parents=True, exist_ok=True)
    index_path.write_text(json.dumps(records_sorted, ensure_ascii=False, indent=2), encoding="utf-8")

    # Compact catalog for agent planning (saves tens of thousands of tokens!)
    catalog_lines = [
        "# Каталог проиндексированных источников (`sources_catalog.md`)",
        "",
        f"Всего валидных уникальных источников: **{len(records_sorted)}**",
        "",
        "| ID | Год | Язык | Стр. | Симв. | Цифры | Балл | Черновик ГОСТ / Файл |",
        "|---|:---:|:---:|:---:|---:|:---:|:---:|---|",
    ]
    for r in records_sorted:
        gost_esc = r["gost_draft"].replace("|", "/")
        catalog_lines.append(
            f"| `{r['id']}` | {r['year'] or '—'} | {r['language']} | {r['pages_total']} | "
            f"{r['char_count']} | {r['empirical_signals']} | {r['quality_score']} | {gost_esc} (`{r['extracted_file']}`) |"
        )
    catalog_path = out_dir / "sources_catalog.md"
    catalog_path.write_text("\n".join(catalog_lines) + "\n", encoding="utf-8")

    if args.combined or args.out:
        comb_path = Path(args.out) if args.out else (out_dir / "sources_clean.txt")
        comb_path.write_text("\n\n".join(combined_chunks), encoding="utf-8")

    print(f"Indexed {len(records_sorted)} unique sources -> {index_path}, {catalog_path}, {extracted_dir}/")

    if args.contract:
        contract = load_contract(args.contract)
        passed, issues = evaluate_exit_gate_1(records_sorted, contract)
        print(f"EXIT GATE 1: {'PASS' if passed else 'FAIL'}")
        for iss in issues:
            print(f"  - {iss}")
        return 0 if passed else 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
