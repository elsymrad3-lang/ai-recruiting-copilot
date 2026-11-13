#!/usr/bin/env python3
"""Send every CV in cvs/ to the n8n CV webhook, one file at a time.

Reads N8N_CV_WEBHOOK from .env. Each request is JSON {"file_name": ..., "text": ...}.
No file is ever skipped: if text extraction fails or finds nothing, the text is sent empty.
"""
import os
import sys
from pathlib import Path

import requests
from docx import Document
from docx.oxml.ns import qn
from dotenv import load_dotenv
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parent
CV_DIR = ROOT / "cvs"
TIMEOUT = 120


def extract_text(path):
    """Return the file's text, or "" if it can't be read or has no text layer."""
    try:
        if path.suffix.lower() == ".pdf":
            return "\n".join(page.extract_text() or "" for page in PdfReader(path).pages).strip()
        if path.suffix.lower() == ".docx":
            # walk every paragraph in the body, including those inside tables
            body = Document(path).element.body
            return "\n".join("".join(t.text or "" for t in p.iter(qn("w:t"))) for p in body.iter(qn("w:p"))).strip()
    except Exception as e:  # a broken file still gets sent, with empty text
        print(f"  ! could not read {path.name}: {e}")
    return ""


def main():
    load_dotenv(ROOT / ".env")
    url = os.getenv("N8N_CV_WEBHOOK", "").strip()
    if not url:
        sys.exit("N8N_CV_WEBHOOK is not set in .env")

    files = sorted((p for p in CV_DIR.iterdir() if p.is_file()), key=lambda p: p.name)
    sent = failed = 0
    for i, path in enumerate(files, 1):
        text = extract_text(path)
        try:
            resp = requests.post(url, json={"file_name": path.name, "text": text}, timeout=TIMEOUT)
            status = str(resp.status_code)
            ok = resp.ok
        except requests.RequestException as e:
            status, ok = f"ERROR ({type(e).__name__})", False
        sent += ok
        failed += not ok
        note = "" if text else "  [empty text]"
        print(f"[{i:>2}/{len(files)}] {status:<6} {path.name}{note}")

    print(f"\nDone: {sent} sent, {failed} failed, {len(files)} total")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
