#!/usr/bin/env python3
"""Download and validate Russian academic PDFs and clean web articles into sources/."""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import re
import sys
import time
from pathlib import Path
from typing import Any
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _common import CYRILLIC_CHAR_RE, ensure_utf8_io

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/131.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,application/pdf,*/*;q=0.8",
    "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.6,en;q=0.4",
}

CAPTCHA_HINTS = (
    "captcha",
    "smartcaptcha",
    "cf-browser-verification",
    "checking your browser",
    "just a moment",
    "verify you are human",
    "ddos-guard",
    "подтвердите, что вы не робот",
)


def slugify(text: str, max_len: int = 60) -> str:
    s = re.sub(r"[^A-Za-zА-Яа-яЁё0-9]+", "_", text).strip("_")
    if not s:
        s = "source"
    return s[:max_len]


def short_hash(url: str) -> str:
    return hashlib.sha1(url.encode("utf-8")).hexdigest()[:8]


def looks_like_captcha_or_block(html: str, status_code: int) -> bool:
    if status_code in (401, 403, 429, 503):
        return True
    low = html[:8000].lower()
    return any(h in low for h in CAPTCHA_HINTS)


def extract_html_metadata(soup: BeautifulSoup, url: str) -> dict[str, Any]:
    meta: dict[str, Any] = {
        "url": url,
        "accessed_date": dt.date.today().strftime("%d.%m.%Y"),
        "title": "",
        "authors": [],
        "year": None,
        "journal": "",
        "volume": "",
        "issue": "",
        "firstpage": "",
        "lastpage": "",
        "doi": "",
        "citation_pdf_url": "",
        "published_date": "",
        "site_name": urlparse(url).netloc.replace("www.", ""),
    }

    def get_meta_all(names: tuple[str, ...]) -> list[str]:
        vals: list[str] = []
        for tag in soup.find_all("meta"):
            key = (tag.get("name") or tag.get("property") or "").strip().lower()
            if key in names:
                content = (tag.get("content") or "").strip()
                if content:
                    vals.append(content)
        return vals

    titles = get_meta_all(("citation_title", "dc.title", "og:title"))
    if titles:
        meta["title"] = titles[0]
    elif soup.title and soup.title.string:
        meta["title"] = soup.title.string.strip()

    meta["authors"] = get_meta_all(("citation_author", "dc.creator", "author"))
    journals = get_meta_all(("citation_journal_title", "dc.source", "og:site_name"))
    if journals:
        meta["journal"] = journals[0]
    for field, keys in (
        ("volume", ("citation_volume",)),
        ("issue", ("citation_issue",)),
        ("firstpage", ("citation_firstpage",)),
        ("lastpage", ("citation_lastpage",)),
        ("doi", ("citation_doi", "dc.identifier")),
        ("citation_pdf_url", ("citation_pdf_url",)),
        ("published_date", ("citation_publication_date", "citation_date", "article:published_time", "dc.date")),
    ):
        vals = get_meta_all(keys)
        if vals:
            meta[field] = vals[0]

    # Parse JSON-LD for structured ScholarlyArticle / NewsArticle metadata
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            raw = json.loads(script.string or "")
            items = raw if isinstance(raw, list) else [raw]
            for obj in items:
                if not isinstance(obj, dict):
                    continue
                if not meta["title"] and obj.get("headline"):
                    meta["title"] = str(obj["headline"]).strip()
                if not meta["published_date"] and obj.get("datePublished"):
                    meta["published_date"] = str(obj["datePublished"]).strip()
                if not meta["authors"] and obj.get("author"):
                    auth = obj["author"]
                    if isinstance(auth, dict) and auth.get("name"):
                        meta["authors"] = [str(auth["name"]).strip()]
                    elif isinstance(auth, list):
                        meta["authors"] = [
                            str(a["name"]).strip()
                            for a in auth
                            if isinstance(a, dict) and a.get("name")
                        ]
        except Exception:
            pass

    # Extract 4-digit year
    date_blob = " ".join([meta["published_date"], meta["title"], url])
    ym = re.search(r"\b(19[5-9]\d|20[0-3]\d)\b", date_blob)
    if ym:
        meta["year"] = int(ym.group(1))

    if meta["citation_pdf_url"]:
        meta["citation_pdf_url"] = urljoin(url, meta["citation_pdf_url"])
    return meta


def validate_pdf_bytes(data: bytes, allow_foreign: bool = False) -> tuple[bool, str, dict[str, Any]]:
    if not data.startswith(b"%PDF-"):
        return False, "not a PDF header (%PDF- missing)", {}
    if len(data) < 30_000:
        return False, f"PDF too small ({len(data)} bytes < 30 KB)", {}

    try:
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz  # type: ignore

        doc = fitz.open(stream=data, filetype="pdf")
        num_pages = len(doc)
        sample_pages = min(num_pages, 12)
        extracted = []
        for i in range(sample_pages):
            extracted.append(doc[i].get_text("text") or "")
        doc.close()
        text = "\n".join(extracted).strip()
    except Exception as exc:
        return False, f"PyMuPDF failed to open PDF: {exc}", {}

    total_chars = len(text)
    cyr_chars = len(CYRILLIC_CHAR_RE.findall(text))
    if total_chars < 1000:
        return False, f"likely scanned image PDF (only {total_chars} chars extracted)", {}
    if not allow_foreign and cyr_chars < 400:
        return False, f"insufficient Cyrillic text layer ({cyr_chars} Cyrillic chars)", {}

    return True, "OK", {"pages": num_pages, "sample_chars": total_chars, "sample_cyr": cyr_chars}


def extract_clean_article_text(soup: BeautifulSoup) -> str:
    # Work on a copy of DOM nodes we remove
    for tag in soup.find_all([
        "script", "style", "noscript", "svg", "iframe", "form",
        "nav", "header", "footer", "aside", "button",
    ]):
        tag.decompose()

    noise_pat = re.compile(
        r"(comment|cookie|banner|subscribe|share|social|sidebar|related|recommend|popup|modal|advert|promo)",
        re.I,
    )
    for tag in list(soup.find_all(True)):
        if getattr(tag, "attrs", None) is None:
            continue
        cls = " ".join(tag.get("class") or [])
        tid = str(tag.get("id") or "")
        if noise_pat.search(cls) or noise_pat.search(tid):
            # Do not accidentally remove main article container
            if tag.name not in ("article", "main", "body"):
                tag.decompose()

    candidates = []
    for selector in (
        "article",
        "[itemprop='articleBody']",
        "main",
        ".article-body",
        ".article__body",
        ".post-content",
        ".content",
    ):
        for el in soup.select(selector):
            txt = el.get_text("\n", strip=True)
            if len(txt) > 400:
                candidates.append((len(txt), txt))

    if candidates:
        candidates.sort(key=lambda x: x[0], reverse=True)
        raw = candidates[0][1]
    else:
        paragraphs = [p.get_text(" ", strip=True) for p in soup.find_all(["h1", "h2", "h3", "p", "li", "td"])]
        raw = "\n".join(p for p in paragraphs if len(p) > 25)

    lines = [re.sub(r"\s+", " ", ln).strip() for ln in raw.splitlines()]
    lines = [ln for ln in lines if len(ln) >= 2]
    return "\n".join(lines)


def request_with_retry(
    session: requests.Session,
    url: str,
    timeout: int = 25,
    referer: str | None = None,
    retries: int = 2,
) -> requests.Response:
    hdrs = dict(HEADERS)
    if referer:
        hdrs["Referer"] = referer
    last_exc: Exception | None = None
    for attempt in range(retries + 1):
        try:
            resp = session.get(url, headers=hdrs, timeout=timeout, allow_redirects=True)
            if resp.status_code in (429, 502, 503, 504) and attempt < retries:
                time.sleep(1.5 * (attempt + 1))
                continue
            return resp
        except Exception as exc:
            last_exc = exc
            if attempt < retries:
                time.sleep(1.2 * (attempt + 1))
    raise RuntimeError(str(last_exc))


def save_sidecar(file_path: Path, meta: dict[str, Any]) -> None:
    sidecar = file_path.with_suffix(file_path.suffix + ".meta.json")
    sidecar.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")


def fetch_one(
    session: requests.Session,
    url: str,
    out_dir: Path,
    timeout: int,
    allow_foreign: bool,
) -> dict[str, Any]:
    parsed = urlparse(url)
    host = parsed.netloc.lower()
    h = short_hash(url)
    is_cyberleninka = "cyberleninka.ru" in host and "/article/n/" in parsed.path
    direct_pdf_hint = parsed.path.lower().endswith(".pdf") or parsed.path.lower().endswith("/pdf")

    meta: dict[str, Any] = {
        "url": url,
        "accessed_date": dt.date.today().strftime("%d.%m.%Y"),
        "site_name": host.replace("www.", ""),
    }

    # Case 1: CyberLeninka article page -> fetch HTML for metadata + PDF URL, then download PDF
    if is_cyberleninka and not parsed.path.rstrip("/").endswith("/pdf"):
        try:
            resp_html = request_with_retry(session, url, timeout=timeout)
        except Exception as exc:
            print(f"[NEED_BROWSER] {url}  # network error: {exc}")
            return {"url": url, "status": "NEED_BROWSER", "reason": str(exc)}

        if looks_like_captcha_or_block(resp_html.text, resp_html.status_code):
            print(f"[NEED_BROWSER] {url}  # HTTP {resp_html.status_code} / bot challenge")
            return {"url": url, "status": "NEED_BROWSER", "reason": f"HTTP {resp_html.status_code}"}

        soup = BeautifulSoup(resp_html.text, "lxml")
        meta = extract_html_metadata(soup, url)
        pdf_url = meta.get("citation_pdf_url") or (url.rstrip("/") + "/pdf")

        try:
            resp_pdf = request_with_retry(session, pdf_url, timeout=timeout, referer=url)
            ok, reason, stats = validate_pdf_bytes(resp_pdf.content, allow_foreign=allow_foreign)
            if ok:
                stem = f"{slugify(meta.get('title') or parsed.path.split('/')[-1])}_{h}"
                pdf_path = out_dir / f"{stem}.pdf"
                pdf_path.write_bytes(resp_pdf.content)
                meta.update({"source_type": "pdf_article", "pdf_url": pdf_url, **stats})
                save_sidecar(pdf_path, meta)
                print(f"[OK PDF] {pdf_path.name} ({stats['pages']} стр., {len(resp_pdf.content)//1024} КБ) <- {url}")
                return {"url": url, "status": "OK", "file": str(pdf_path), "type": "pdf"}
            else:
                # Fallback: check if CyberLeninka HTML itself has full OCR/text body
                body_txt = extract_clean_article_text(soup)
                if len(body_txt) >= 3500 and len(CYRILLIC_CHAR_RE.findall(body_txt)) >= 1500:
                    stem = f"{slugify(meta.get('title') or parsed.path.split('/')[-1])}_{h}"
                    txt_path = out_dir / f"{stem}.txt"
                    header = (
                        f"URL: {url}\n"
                        f"TITLE: {meta.get('title', '')}\n"
                        f"AUTHORS: {'; '.join(meta.get('authors') or [])}\n"
                        f"JOURNAL: {meta.get('journal', '')}\n"
                        f"YEAR: {meta.get('year') or ''}\n"
                        f"ACCESSED: {meta['accessed_date']}\n"
                        f"---\n\n"
                    )
                    txt_path.write_text(header + body_txt, encoding="utf-8")
                    meta["source_type"] = "web_article_fulltext"
                    save_sidecar(txt_path, meta)
                    print(f"[OK TXT-FALLBACK] {txt_path.name} (PDF failed: {reason}, saved full HTML text)")
                    return {"url": url, "status": "OK", "file": str(txt_path), "type": "txt"}
                print(f"[NEED_BROWSER] {url}  # CyberLeninka PDF check failed: {reason}")
                return {"url": url, "status": "NEED_BROWSER", "reason": reason}
        except Exception as exc:
            print(f"[NEED_BROWSER] {url}  # PDF fetch error: {exc}")
            return {"url": url, "status": "NEED_BROWSER", "reason": str(exc)}

    # Case 2: Generic URL (could be direct PDF or HTML article/report)
    try:
        resp = request_with_retry(session, url, timeout=timeout)
    except Exception as exc:
        print(f"[NEED_BROWSER] {url}  # network error: {exc}")
        return {"url": url, "status": "NEED_BROWSER", "reason": str(exc)}

    ctype = (resp.headers.get("Content-Type") or "").lower()
    if resp.status_code >= 400:
        print(f"[NEED_BROWSER] {url}  # HTTP {resp.status_code}")
        return {"url": url, "status": "NEED_BROWSER", "reason": f"HTTP {resp.status_code}"}

    # Direct PDF response
    if "application/pdf" in ctype or resp.content.startswith(b"%PDF-") or direct_pdf_hint:
        ok, reason, stats = validate_pdf_bytes(resp.content, allow_foreign=allow_foreign)
        if not ok:
            print(f"[NEED_BROWSER] {url}  # invalid PDF: {reason}")
            return {"url": url, "status": "NEED_BROWSER", "reason": reason}
        base_name = Path(parsed.path).stem or "document"
        pdf_path = out_dir / f"{slugify(base_name)}_{h}.pdf"
        pdf_path.write_bytes(resp.content)
        meta.update({"source_type": "pdf_direct", **stats})
        save_sidecar(pdf_path, meta)
        print(f"[OK PDF] {pdf_path.name} ({stats['pages']} стр., {len(resp.content)//1024} КБ) <- {url}")
        return {"url": url, "status": "OK", "file": str(pdf_path), "type": "pdf"}

    # HTML page
    resp.encoding = resp.apparent_encoding or "utf-8"
    html = resp.text
    if looks_like_captcha_or_block(html, resp.status_code):
        print(f"[NEED_BROWSER] {url}  # JS/bot challenge detected")
        return {"url": url, "status": "NEED_BROWSER", "reason": "bot challenge"}

    soup = BeautifulSoup(html, "lxml")
    meta = extract_html_metadata(soup, url)

    # If HTML advertises citation_pdf_url (e.g. OJS/VAK university journals), prefer PDF!
    if meta.get("citation_pdf_url"):
        try:
            resp_pdf = request_with_retry(session, meta["citation_pdf_url"], timeout=timeout, referer=url)
            ok, _, stats = validate_pdf_bytes(resp_pdf.content, allow_foreign=allow_foreign)
            if ok:
                stem = f"{slugify(meta.get('title') or 'article')}_{h}"
                pdf_path = out_dir / f"{stem}.pdf"
                pdf_path.write_bytes(resp_pdf.content)
                meta.update({"source_type": "pdf_article", **stats})
                save_sidecar(pdf_path, meta)
                print(f"[OK PDF] {pdf_path.name} ({stats['pages']} стр., via citation_pdf_url) <- {url}")
                return {"url": url, "status": "OK", "file": str(pdf_path), "type": "pdf"}
        except Exception:
            pass

    clean_text = extract_clean_article_text(soup)
    if len(clean_text) < 800:
        print(f"[NEED_BROWSER] {url}  # too little text extracted ({len(clean_text)} chars; likely JS-rendered)")
        return {"url": url, "status": "NEED_BROWSER", "reason": "JS-rendered or empty"}

    stem = f"{slugify(meta.get('title') or host)}_{h}"
    txt_path = out_dir / f"{stem}.txt"
    header = (
        f"URL: {url}\n"
        f"TITLE: {meta.get('title', '')}\n"
        f"AUTHORS: {'; '.join(meta.get('authors') or [])}\n"
        f"SITE: {meta.get('site_name', '')}\n"
        f"PUBLISHED: {meta.get('published_date', '')}\n"
        f"YEAR: {meta.get('year') or ''}\n"
        f"ACCESSED: {meta['accessed_date']}\n"
        f"---\n\n"
    )
    txt_path.write_text(header + clean_text, encoding="utf-8")
    meta["source_type"] = "web_article"
    meta["chars"] = len(clean_text)
    save_sidecar(txt_path, meta)
    print(f"[OK WEB] {txt_path.name} ({len(clean_text)} симв.) <- {url}")
    return {"url": url, "status": "OK", "file": str(txt_path), "type": "txt"}


def collect_urls(raw_args: list[str]) -> list[str]:
    urls: list[str] = []
    for item in raw_args:
        p = Path(item)
        if p.is_file():
            for line in p.read_text(encoding="utf-8").splitlines():
                ln = line.strip()
                if ln and not ln.startswith("#"):
                    urls.append(ln)
        elif item.startswith(("http://", "https://")):
            urls.append(item.strip())
    # Deduplicate while preserving order
    seen: set[str] = set()
    uniq: list[str] = []
    for u in urls:
        if u not in seen:
            seen.add(u)
            uniq.append(u)
    return uniq


def main() -> int:
    ensure_utf8_io()
    parser = argparse.ArgumentParser(
        description="Fetch and validate Russian academic PDFs and web case studies."
    )
    parser.add_argument(
        "--urls",
        nargs="+",
        required=True,
        help="List of URLs and/or paths to text files containing one URL per line.",
    )
    parser.add_argument("--out", default="sources", help="Output directory (default: sources/)")
    parser.add_argument("--delay", type=float, default=0.8, help="Delay between requests in seconds")
    parser.add_argument("--timeout", type=int, default=25, help="HTTP timeout in seconds")
    parser.add_argument(
        "--allow-foreign",
        action="store_true",
        help="Allow non-Cyrillic PDFs (when foreign sources are permitted by contract)",
    )
    args = parser.parse_args()

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    urls = collect_urls(args.urls)
    if not urls:
        print("ERROR: no valid URLs supplied.", file=sys.stderr)
        return 2

    session = requests.Session()
    results = []
    for idx, u in enumerate(urls):
        if idx > 0 and args.delay > 0:
            time.sleep(args.delay)
        results.append(fetch_one(session, u, out_dir, args.timeout, args.allow_foreign))

    ok_cnt = sum(1 for r in results if r["status"] == "OK")
    need_cnt = sum(1 for r in results if r["status"] == "NEED_BROWSER")
    print(f"\nSUMMARY: {ok_cnt} downloaded & validated, {need_cnt} require browser/replacement (total {len(results)}).")
    return 0 if ok_cnt > 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
