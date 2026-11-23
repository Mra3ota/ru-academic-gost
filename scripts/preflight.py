#!/usr/bin/env python3
"""Environment diagnostic for ru-academic-gost skill."""

from __future__ import annotations

import argparse
import importlib
import json
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _common import ensure_utf8_io, find_pandoc, find_soffice, word_com_available


def check_module(name: str) -> tuple[bool, str]:
    try:
        mod = importlib.import_module(name)
        ver = getattr(mod, "__version__", getattr(mod, "VersionBind", "installed"))
        return True, str(ver)
    except Exception as exc:
        return False, str(exc)


def main() -> int:
    ensure_utf8_io()
    parser = argparse.ArgumentParser(
        description="Check all Python packages and system tools required by ru-academic-gost."
    )
    parser.add_argument("--json", action="store_true", help="Output JSON report")
    args = parser.parse_args()

    py_modules = {
        "requests": "requests",
        "beautifulsoup4": "bs4",
        "pymupdf": "pymupdf",
        "python-docx": "docx",
        "lxml": "lxml",
        "pillow": "PIL",
    }
    py_status: dict[str, dict[str, object]] = {}
    missing_pip: list[str] = []
    for pkg, mod in py_modules.items():
        ok, info = check_module(mod)
        py_status[pkg] = {"ok": ok, "info": info}
        if not ok:
            missing_pip.append(pkg)

    humanizer_cmds = [
        "humanizer-scan",
        "humanizer-markers",
        "humanizer-clean",
        "humanizer-facts",
        "humanizer-report",
    ]
    hum_status = {cmd: shutil.which(cmd) for cmd in humanizer_cmds}
    hum_ok = all(hum_status.values())
    if not hum_ok:
        missing_pip.append("humanizer-ru")

    pandoc_path = find_pandoc()
    if not pandoc_path:
        missing_pip.append("pypandoc-binary")

    word_ok = word_com_available()
    soffice_path = find_soffice()
    pdf_engine = (
        "Microsoft Word COM"
        if word_ok
        else (f"LibreOffice ({soffice_path})" if soffice_path else "MISSING")
    )
    png_engine = (
        "PyMuPDF (fitz)"
        if py_status["pymupdf"]["ok"]
        else (shutil.which("pdftoppm") or "MISSING")
    )

    ready = (
        all(v["ok"] for v in py_status.values())
        and hum_ok
        and bool(pandoc_path)
        and (word_ok or bool(soffice_path))
        and png_engine != "MISSING"
    )

    report = {
        "ready": ready,
        "python": sys.version.split()[0],
        "python_packages": py_status,
        "humanizer_ru": {"ok": hum_ok, "commands": hum_status},
        "pandoc": {"ok": bool(pandoc_path), "path": pandoc_path},
        "pdf_renderer": {"ok": word_ok or bool(soffice_path), "engine": pdf_engine},
        "png_renderer": {"ok": png_engine != "MISSING", "engine": png_engine},
        "install_hint": (
            f"python -m pip install --upgrade {' '.join(missing_pip)}"
            if missing_pip
            else None
        ),
    }

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"=== ru-academic-gost preflight (Python {report['python']}) ===")
        for pkg, st in py_status.items():
            mark = "PASS" if st["ok"] else "FAIL"
            print(f"  [{mark}] {pkg}: {st['info']}")
        print(f"  [{'PASS' if hum_ok else 'FAIL'}] humanizer-ru CLI: {sum(bool(v) for v in hum_status.values())}/{len(hum_status)} commands")
        print(f"  [{'PASS' if pandoc_path else 'FAIL'}] pandoc: {pandoc_path or 'MISSING'}")
        print(f"  [{'PASS' if (word_ok or soffice_path) else 'FAIL'}] DOCX->PDF renderer: {pdf_engine}")
        print(f"  [{'PASS' if png_engine != 'MISSING' else 'FAIL'}] PDF->PNG renderer: {png_engine}")
        if missing_pip:
            print(f"\nInstall missing Python packages:\n  python -m pip install --upgrade {' '.join(missing_pip)}")
        if not (word_ok or soffice_path):
            print("\nInstall LibreOffice for DOCX->PDF rendering:\n  winget install TheDocumentFoundation.LibreOffice")
        print(f"\nOVERALL: {'PASS (ready)' if ready else 'FAIL (action needed)'}")

    return 0 if ready else 1


if __name__ == "__main__":
    raise SystemExit(main())
