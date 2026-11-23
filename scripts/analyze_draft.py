#!/usr/bin/env python3
"""Phase 3-4 Quality Gate: budget, Russian anti-slop metrics, quote grounding, citations, and humanizer-ru."""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _common import (
    ensure_utf8_io,
    find_hidden_unicode,
    find_homoglyphs,
    load_contract,
    strip_bib_prefix,
)

# Cliché ban-list (checked outside «...» direct quotes)
CLICHE_PATTERNS: list[tuple[str, re.Pattern[str]]] = [
    ("В современном мире", re.compile(r"\bв\s+современном\s+мире\b", re.I)),
    ("В условиях современной реальности", re.compile(r"\bв\s+условиях\s+современн\w+\s+реальност\w+\b", re.I)),
    ("В эпоху стремительной цифровизации", re.compile(r"\bв\s+эпоху\s+\w*\s*цифровизац\w+\b", re.I)),
    ("На сегодняшний день", re.compile(r"\bна\s+сегодняшний\s+день\b", re.I)),
    ("Ни для кого не секрет", re.compile(r"\bни\s+для\s+кого\s+не\s+секрет\b", re.I)),
    ("Стоит отметить, что", re.compile(r"\bстоит\s+отметить\s*,\s*что\b", re.I)),
    ("Важно понимать/отметить, что", re.compile(r"\bважно\s+(?:понимать|отметить|подчеркнуть)\s*,\s*что\b", re.I)),
    ("Следует подчеркнуть, что", re.compile(r"\bследует\s+подчеркнуть\s*,\s*что\b", re.I)),
    ("Неудивительно, что", re.compile(r"\bнеудивительно\s*,\s*что\b", re.I)),
    ("Нельзя не упомянуть", re.compile(r"\bнельзя\s+не\s+упомянуть\b", re.I)),
    ("Таким образом, можно с уверенностью", re.compile(r"\bтаким\s+образом\s*,\s*можно\s+с\s+уверенностью\b", re.I)),
    ("Подводя итог вышесказанному", re.compile(r"\bподводя\s+итог\s+вышесказанн\w+\b", re.I)),
    ("краеугольный камень", re.compile(r"\bкраеугольн\w+\s+камн\w*\b", re.I)),
    ("ключевой драйвер", re.compile(r"\bключев\w+\s+драйвер\w*\b", re.I)),
    ("широкий спектр", re.compile(r"\bширок\w+\s+спектр\w*\b", re.I)),
    ("комплексный подход", re.compile(r"\bкомплексн\w+\s+подход\w*\b", re.I)),
    ("играет решающую/ключевую роль", re.compile(r"\bигра\w+\s+(?:решающ\w+|ключев\w+)\s+рол\w+\b", re.I)),
    ("открывает новые горизонты", re.compile(r"\bоткрыва\w+\s+нов\w+\s+горизонт\w*\b", re.I)),
    ("тонкая грань", re.compile(r"\bтонк\w+\s+гран\w+\b", re.I)),
    ("яркий пример того, как", re.compile(r"\bярк\w+\s+пример\w*\s+того\s*,\s*как\b", re.I)),
    ("многогранный феномен", re.compile(r"\bмногогранн\w+\s+феномен\w*\b", re.I)),
    ("не просто X, а Y", re.compile(r"\bне\s+просто\s+[^.,;!?]{2,50},\s*а\s+", re.I)),
    ("не столько X, сколько Y", re.compile(r"\bне\s+столько\s+[^.,;!?]{2,50},\s*сколько\s+", re.I)),
]

CONCLUSION_SELF_REVIEW_PATTERNS: list[re.Pattern[str]] = [
    re.compile(r"\bв\s+данной\s+(?:работе|курсовой|вкр|статье)\s+(?:успешно|подробно|всесторонне)\b", re.I),
    re.compile(r"\bбыл(?:а|о)?\s+(?:успешно\s+)?(?:рассмотрена\s+актуальная\s+тема|провед[её]н\s+всесторонн\w+\s+анализ)\b", re.I),
    re.compile(r"\bработа\s+имеет\s+большую\s+(?:теоретическую|практическую)\s+значимость\b", re.I),
]

VERBAL_NOUN_RE = re.compile(
    r"\b[А-Яа-яЁё]{4,}(?:ени[еяюй]|ани[еяюй]|ениям|аниям|ениями|аниями|ениях|аниях|"
    r"ост[иь]|остью|остей|остям|остями|остях|ти[еяюй]|ство|ства|ству|ством|стве)\b"
)

LADDER_OPENERS_RE = re.compile(
    r"^\s*(?:во-первых|во-вторых|в-третьих|в-четв[её]ртых|прежде\s+всего|кроме\s+того|более\s+того|наконец)\b",
    re.I,
)


def strip_footnotes_and_markdown(md: str) -> tuple[str, dict[str, str]]:
    """Separate Markdown body prose from footnote definitions [^id]: ..."""
    fn_defs: dict[str, str] = {}
    body_lines: list[str] = []
    cur_id: str | None = None
    for ln in md.splitlines():
        m = re.match(r"^\[\^([^\]]+)\]:\s*(.*)$", ln)
        if m:
            cur_id = m.group(1)
            fn_defs[cur_id] = m.group(2).strip()
            continue
        if cur_id is not None and (ln.startswith("    ") or ln.startswith("\t")):
            fn_defs[cur_id] += " " + ln.strip()
            continue
        cur_id = None
        body_lines.append(ln)
    return "\n".join(body_lines), fn_defs


def normalize_for_match(text: str) -> str:
    s = text.lower().replace("ё", "е")
    s = re.sub(r"[^а-яa-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def compact_for_match(text: str) -> str:
    """Alphanumeric-only normalization immune to PDF line-wrap dehyphenation (e.g. бренд-метрик vs брендметрик)."""
    return re.sub(r"[^а-яa-z0-9]+", "", text.lower().replace("ё", "е"))


def load_extracted_corpus(extracted_dir: Path | None) -> list[tuple[str, str, str]]:
    if not extracted_dir or not extracted_dir.exists():
        return []
    corpus: list[tuple[str, str, str]] = []
    for f in sorted(extracted_dir.glob("*.txt")):
        raw = f.read_text(encoding="utf-8", errors="replace")
        corpus.append((f.name, normalize_for_match(raw), compact_for_match(raw)))
    return corpus


def verify_quote_in_corpus(quote: str, corpus: list[tuple[str, str, str]]) -> str | None:
    nq = normalize_for_match(quote)
    cq = compact_for_match(quote)
    if len(nq) < 20:
        return "short_skip"
    # Direct normalized or compact match (immune to PDF line-break hyphenation)
    for fname, ctext, ccomp in corpus:
        if nq in ctext or (len(cq) >= 18 and cq in ccomp):
            return fname
    # Allow ellipsis [...] splits inside scholarly quotes
    parts = [compact_for_match(p) for p in re.split(r"\[\.{3}\]|\.{3}|…", quote) if len(compact_for_match(p)) >= 14]
    if parts:
        for fname, _, ccomp in corpus:
            if all(p in ccomp for p in parts):
                return fname
    return None


def run_cli_json(cmd: list[str]) -> dict[str, Any] | None:
    if not shutil.which(cmd[0]):
        return None
    env = dict(os.environ)
    env["PYTHONUTF8"] = "1"
    try:
        res = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=45, env=env)
        out = (res.stdout or "").strip()
        if out.startswith("{"):
            data = json.loads(out)
            data["_returncode"] = res.returncode
            return data
        return {"_returncode": res.returncode, "_raw": out[:600], "_stderr": (res.stderr or "")[:300]}
    except Exception as exc:
        return {"_error": str(exc)}


def analyze_file(
    fpath: Path,
    contract: dict[str, Any],
    corpus: list[tuple[str, str, str]],
    before_path: Path | None,
    biblio_count: int | None,
    biblio_keys: set[str],
    target_chars: int | None,
) -> dict[str, Any]:
    raw = fpath.read_text(encoding="utf-8")
    body_md, fn_defs = strip_footnotes_and_markdown(raw)

    # Strip display math blocks $$...$$ before prose style analysis so LaTeX never skews metrics
    body_no_math = re.sub(r"\$\$[\s\S]*?\$\$", " ", body_md)

    # Plain prose excluding headings, images & markdown table borders for style metrics
    prose_lines = [
        ln for ln in body_no_math.splitlines()
        if ln.strip() and not ln.lstrip().startswith(("#", "|", "```", ">", "!["))
    ]
    prose = "\n".join(prose_lines)
    prose_no_fn = re.sub(r"\[\^[^\]]+\]", "", prose)
    prose_no_fn = re.sub(r"\[@[A-Za-zА-Яа-яЁё0-9_\-][^\]]*\]", "", prose_no_fn)
    prose_no_quotes = re.sub(r"«[^»]{2,400}»", " ", prose_no_fn)

    chars_with_spaces = len(re.sub(r"\n{2,}", "\n", prose_no_fn).strip())
    words = re.findall(r"[А-Яа-яЁёA-Za-z0-9\-]+", prose_no_fn)
    word_count = len(words)
    citation_style = contract.get("citations", {}).get("style", "footnote")
    chars_per_page = 1650 if citation_style == "footnote" else 1800
    est_pages = round(chars_with_spaces / chars_per_page, 2)

    # Budget check
    budget_info: dict[str, Any] = {
        "chars_with_spaces": chars_with_spaces,
        "words": word_count,
        "estimated_pages": est_pages,
        "target_chars": target_chars,
        "deviation_pct": None,
        "status": "INFO",
    }
    if target_chars and target_chars > 0:
        dev = round((chars_with_spaces - target_chars) / target_chars * 100.0, 1)
        budget_info["deviation_pct"] = dev
        budget_info["status"] = "PASS" if abs(dev) <= 15.0 else "FAIL"

    # Verbal nouns & cascades
    vn_matches = list(VERBAL_NOUN_RE.finditer(prose_no_quotes))
    vn_ratio_pct = round(len(vn_matches) / max(1, word_count) * 100.0, 1)
    cascades = 0
    for sent in re.split(r"[.!?]+", prose_no_quotes):
        if len(VERBAL_NOUN_RE.findall(sent)) >= 3:
            cascades += 1

    # Clichés (outside direct quotes)
    cliches_found: list[str] = []
    for label, pat in CLICHE_PATTERNS:
        hits = pat.findall(prose_no_quotes)
        if hits:
            cliches_found.append(f"{label} (x{len(hits)})")

    # Conclusion self-review check
    is_conclusion = "conclusion" in fpath.name.lower() or bool(re.search(r"^#\s*заключение\b", body_md, re.I | re.M))
    self_review_hits: list[str] = []
    if is_conclusion:
        for pat in CONCLUSION_SELF_REVIEW_PATTERNS:
            for m in pat.finditer(prose_no_quotes):
                self_review_hits.append(m.group(0))

    # Paragraph opener ladders (3+ consecutive paragraphs starting with ladder markers)
    paras = [p.strip() for p in re.split(r"\n\s*\n", prose_no_fn) if p.strip()]
    max_ladder_run = 0
    cur_run = 0
    for p in paras:
        if LADDER_OPENERS_RE.match(p):
            cur_run += 1
            max_ladder_run = max(max_ladder_run, cur_run)
        else:
            cur_run = 0

    # Bold-lead bullet lists inside prose (`- **Заголовок:** ...`)
    ai_bullet_lists = len(re.findall(r"^\s*[-*]\s+\*\*[^*]{2,60}:\*\*", body_md, re.M))

    # Punctuation & rhythm
    em_dashes_per_100 = round(prose_no_quotes.count("—") / max(1, word_count) * 100.0, 2)
    colons_per_100 = round(prose_no_quotes.count(":") / max(1, word_count) * 100.0, 2)
    sent_lens = [len(re.findall(r"[А-Яа-яЁёA-Za-z0-9]+", s)) for s in re.split(r"[.!?]+", prose_no_fn) if len(s.strip()) > 10]
    if len(sent_lens) >= 3:
        mean_s = sum(sent_lens) / len(sent_lens)
        std_s = round(math.sqrt(sum((x - mean_s) ** 2 for x in sent_lens) / len(sent_lens)), 1)
    else:
        mean_s, std_s = 0.0, 0.0

    # Direct quotes «...» audit + corpus grounding
    quotes = [m.group(1).strip() for m in re.finditer(r"«([^»]{25,500})»", prose_no_fn)]
    quote_chars = sum(len(q) for q in quotes)
    quote_share_pct = round(quote_chars / max(1, chars_with_spaces) * 100.0, 1)
    unverified_quotes: list[str] = []
    verified_count = 0
    if corpus:
        for q in quotes:
            hit = verify_quote_in_corpus(q, corpus)
            if hit:
                verified_count += 1
            else:
                unverified_quotes.append(q[:110] + ("..." if len(q) > 110 else ""))

    # Footnotes / bracket citations check
    fn_refs = re.findall(r"\[\^([^\]]+)\]", body_md)
    missing_fn_defs = sorted(set(fn_refs) - set(fn_defs.keys()))
    unused_fn_defs = sorted(set(fn_defs.keys()) - set(fn_refs))
    bad_fn_spacing = len(
        re.findall(
            r"\s+\[\^[^\]]+\]|\[\^[^\]]+\]\s+[.,;:]|[,;:]\[\^[^\]]+\]|"
            r"(?<!\bруб)(?<!\bтыс)(?<!\bмлн)(?<!\bмлрд)(?<!\bст)(?<!\bдр)(?<!\bсм)(?<!\bпр)(?<!\bгг)(?<!\bвв)(?<!\s[гвс])(?<![а-яё]\.[а-яё])\.\[\^[^\]]+\](?!\.)",
            body_md,
        )
    )
    tam_zhe_count = sum(1 for v in fn_defs.values() if re.match(r"^\s*там\s+же\b", v, re.I))

    bracket_refs = [int(n) for n in re.findall(r"\[(\d{1,3})(?:\s*,\s*с\.\s*[\d–\-]+)?\]", body_md)]
    out_of_range_brackets = (
        sorted({n for n in bracket_refs if n < 1 or n > biblio_count})
        if (biblio_count and bracket_refs)
        else []
    )

    sym_refs = re.findall(r"@([A-Za-zА-Яа-яЁё0-9_\-]+)", " ".join(re.findall(r"\[(@[^\]]+)\]", body_md)))
    missing_sym_keys = sorted(set(sym_refs) - biblio_keys) if (biblio_keys and sym_refs) else []

    # Local Unicode & homoglyph check
    hidden_uni = find_hidden_unicode(raw)
    homoglyphs = find_homoglyphs(raw)

    # External humanizer-ru checks
    hum_markers = run_cli_json(["humanizer-markers", "--json", str(fpath)])
    hum_scan = run_cli_json(["humanizer-scan", "--genre", "academic", "--json", str(fpath)])
    hum_facts = None
    hum_report = None
    if before_path and before_path.exists():
        hum_facts = run_cli_json(["humanizer-facts", "diff", "--json", str(before_path), str(fpath)])
        hum_report = run_cli_json(["humanizer-report", "--json", str(before_path), str(fpath)])

    # Extract concise summary from humanizer-ru outputs
    markers_count = 0
    if hum_markers and isinstance(hum_markers.get("files"), list):
        for fobj in hum_markers["files"]:
            markers_count += len(fobj.get("findings") or fobj.get("markers") or [])

    scan_signals = 0
    scan_cats = 0
    if hum_scan and isinstance(hum_scan.get("files"), list) and hum_scan["files"]:
        f0 = hum_scan["files"][0]
        scan_signals = int(f0.get("total", f0.get("signals", len(f0.get("findings") or []))))
        cats = f0.get("categories") or []
        scan_cats = len(cats) if isinstance(cats, (list, dict)) else int(cats or 0)

    facts_ok = True
    facts_lost = 0
    facts_changed = 0
    if hum_facts and isinstance(hum_facts.get("counts"), dict):
        facts_lost = int(hum_facts["counts"].get("lost", 0))
        facts_changed = int(hum_facts["counts"].get("changed", 0))
        facts_ok = (facts_lost == 0 and facts_changed == 0)

    # Hard fail criteria
    failures: list[str] = []
    if budget_info["status"] == "FAIL":
        failures.append(f"Отклонение объёма {budget_info['deviation_pct']}% (> ±15% от цели {target_chars} симв.)")
    if hidden_uni:
        failures.append(f"Найдены скрытые символы Unicode: {hidden_uni}")
    if homoglyphs:
        failures.append(f"Найдены гомоглифы в кириллических словах: {[h['token'] for h in homoglyphs[:5]]}")
    if markers_count > 0:
        failures.append(f"humanizer-markers нашёл {markers_count} артефактов копипасты/вставки")
    if cliches_found:
        failures.append(f"Штампы из бан-листа: {', '.join(cliches_found[:6])}")
    if self_review_hits:
        failures.append(f"Тон рецензии в Заключении: {self_review_hits}")
    if max_ladder_run >= 3:
        failures.append(f"Лесенка зачинов абзацев ({max_ladder_run} подряд)")
    if ai_bullet_lists > 0:
        failures.append(f"ИИ-списки вида '- **Заголовок:** ...' в прозе: {ai_bullet_lists}")
    if missing_fn_defs:
        failures.append(f"Ссылки на неопределённые сноски: {missing_fn_defs}")
    if unused_fn_defs:
        failures.append(f"Неиспользуемые определения сносок: {unused_fn_defs}")
    if out_of_range_brackets:
        failures.append(f"Ссылки [N] вне диапазона библиографии 1..{biblio_count}: {out_of_range_brackets}")
    if missing_sym_keys:
        failures.append(f"Неопределённые символьные ключи ссылок [@key]: {missing_sym_keys}")
    if unverified_quotes:
        failures.append(f"Неверифицированные прямые цитаты ({len(unverified_quotes)}): {unverified_quotes[:2]}")
    if not facts_ok:
        failures.append(f"Нарушен Факт-замок (humanizer-facts): lost={facts_lost}, changed={facts_changed}")

    warnings: list[str] = []
    if vn_ratio_pct > 15.0:
        warnings.append(f"Высокая доля отглагольных существительных: {vn_ratio_pct}% (цель ≤ 14%)")
    if cascades >= 4:
        warnings.append(f"Каскады из ≥3 отглагольных существительных в предложении: {cascades}")
    if bad_fn_spacing > 0:
        warnings.append(f"Пробелы/порядок знаков у сносок ({bad_fn_spacing} мест — будет автоисправлено в build_gost.py)")
    if em_dashes_per_100 > 2.5:
        warnings.append(f"Много длинных тире: {em_dashes_per_100} на 100 слов")
    if scan_signals >= 5 and scan_cats >= 2:
        warnings.append(f"humanizer-scan отметил {scan_signals} мягких признаков в {scan_cats} категориях")
    if std_s and std_s < 5.0 and len(sent_lens) >= 8:
        warnings.append(f"Монотонный ритм длины предложений (std={std_s} слов; чередуйте короткие и длинные)")

    return {
        "file": str(fpath),
        "status": "PASS" if not failures else "FAIL",
        "budget": budget_info,
        "style": {
            "verbal_nouns_pct": vn_ratio_pct,
            "verbal_noun_cascades": cascades,
            "cliches": cliches_found,
            "max_ladder_run": max_ladder_run,
            "ai_bullet_lists": ai_bullet_lists,
            "em_dashes_per_100w": em_dashes_per_100,
            "colons_per_100w": colons_per_100,
            "sentence_len_mean": round(mean_s, 1),
            "sentence_len_std": std_s,
        },
        "quotes": {
            "count_ge_25c": len(quotes),
            "quote_share_pct": quote_share_pct,
            "verified_in_corpus": verified_count if corpus else None,
            "unverified": unverified_quotes,
        },
        "citations": {
            "footnote_refs": len(fn_refs),
            "footnote_defs": len(fn_defs),
            "missing_defs": missing_fn_defs,
            "unused_defs": unused_fn_defs,
            "bad_fn_spacing": bad_fn_spacing,
            "tam_zhe_count": tam_zhe_count,
            "bracket_refs": len(bracket_refs),
            "out_of_range_brackets": out_of_range_brackets,
        },
        "hygiene": {
            "hidden_unicode": hidden_uni,
            "homoglyphs": homoglyphs,
            "humanizer_markers_count": markers_count,
            "humanizer_scan_signals": scan_signals,
            "humanizer_scan_categories": scan_cats,
            "humanizer_facts_ok": facts_ok if before_path else None,
            "humanizer_facts_lost": facts_lost if before_path else None,
            "humanizer_facts_changed": facts_changed if before_path else None,
        },
        "failures": failures,
        "warnings": warnings,
    }


def parse_biblio_info(biblio_path: Path | None) -> tuple[int | None, set[str]]:
    if not biblio_path or not biblio_path.exists():
        return None, set()
    lines = biblio_path.read_text(encoding="utf-8").splitlines()
    sym_keys = {
        m.group(1)
        for ln in lines
        for m in [re.match(r"^\s*\[@([A-Za-zА-Яа-яЁё0-9_\-]+)\]\s+\S", ln)]
        if m
    }
    if sym_keys:
        return len(sym_keys), sym_keys
    entries = [ln for ln in lines if re.match(r"^\s*\d+\.\s+\S", ln)]
    return (len(entries) if entries else None), set()


def main() -> int:
    ensure_utf8_io()
    parser = argparse.ArgumentParser(
        description="Analyze Markdown drafts for volume, style, quote grounding, citations, and humanizer-ru checks."
    )
    parser.add_argument("files", nargs="+", help="Draft Markdown file(s) to analyze")
    parser.add_argument("--before", default="", help="Snapshot .before.md file for humanizer-facts diff (single file mode)")
    parser.add_argument("--contract", default="", help="Path to project_contract.json")
    parser.add_argument("--extracted", default="", help="Directory drafts/extracted/ for quote grounding")
    parser.add_argument("--biblio", default="", help="Path to bibliography .md file for [N] / [@key] validation")
    parser.add_argument("--target-chars", type=int, default=0, help="Target characters with spaces for this file")
    parser.add_argument("--json", action="store_true", help="Output JSON report")
    args = parser.parse_args()

    contract = load_contract(args.contract)
    corpus = load_extracted_corpus(Path(args.extracted) if args.extracted else None)
    bib_cnt, bib_keys = parse_biblio_info(Path(args.biblio) if args.biblio else None)
    before_p = Path(args.before) if args.before else None

    results = []
    for f in args.files:
        fp = Path(f)
        if not fp.exists():
            print(f"ERROR: file not found: {fp}", file=sys.stderr)
            return 2
        res = analyze_file(
            fp,
            contract,
            corpus,
            before_p if len(args.files) == 1 else None,
            bib_cnt,
            bib_keys,
            args.target_chars if args.target_chars > 0 else None,
        )
        results.append(res)

    if args.json:
        print(json.dumps(results, ensure_ascii=False, indent=2))
    else:
        for r in results:
            print(f"\n=== ANALYZE: {r['file']} -> {r['status']} ===")
            b = r["budget"]
            dev_s = f" ({b['deviation_pct']:+.1f}% vs {b['target_chars']})" if b["target_chars"] else ""
            print(f"  Объём: {b['chars_with_spaces']} симв. с пробелами (~{b['estimated_pages']} стр., {b['words']} слов){dev_s}")
            st = r["style"]
            print(
                f"  Стиль: отглагольные={st['verbal_nouns_pct']}% (каскады={st['verbal_noun_cascades']}), "
                f"тире/100сл={st['em_dashes_per_100w']}, ритм={st['sentence_len_mean']}±{st['sentence_len_std']} слов"
            )
            q = r["quotes"]
            v_s = f", проверено в базе: {q['verified_in_corpus']}/{q['count_ge_25c']}" if q["verified_in_corpus"] is not None else ""
            print(f"  Цитаты «...»: {q['count_ge_25c']} шт. ({q['quote_share_pct']}% текста){v_s}")
            h = r["hygiene"]
            facts_s = f", facts_lock={'PASS' if h['humanizer_facts_ok'] else 'FAIL'}" if h["humanizer_facts_ok"] is not None else ""
            print(
                f"  humanizer-ru: markers={h['humanizer_markers_count']}, "
                f"scan={h['humanizer_scan_signals']} ({h['humanizer_scan_categories']} кат.){facts_s}"
            )
            for fail in r["failures"]:
                print(f"  [FAIL] {fail}")
            for warn in r["warnings"]:
                print(f"  [WARN] {warn}")

    return 0 if all(r["status"] == "PASS" for r in results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
