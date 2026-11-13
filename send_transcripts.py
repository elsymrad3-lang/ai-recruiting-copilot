#!/usr/bin/env python3
"""Send each client call transcript in transcripts/ to the n8n JD webhook, one at a time.

Reads N8N_JD_WEBHOOK from .env. Each request is JSON {"client": ..., "transcript": ...}.
The client name comes from the file name, "<Client> - <Role>.txt" -> "<Client>".
"""
import os
import sys
from pathlib import Path

import requests
from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent
TRANSCRIPT_DIR = ROOT / "transcripts"
TIMEOUT = 120


def client_name(path):
    return path.stem.split(" - ")[0].strip()


def main():
    load_dotenv(ROOT / ".env")
    url = os.getenv("N8N_JD_WEBHOOK", "").strip()
    if not url:
        sys.exit("N8N_JD_WEBHOOK is not set in .env")

    files = sorted(TRANSCRIPT_DIR.glob("*.txt"), key=lambda p: p.name)
    failed = 0
    for path in files:
        payload = {"client": client_name(path), "transcript": path.read_text(encoding="utf-8")}
        try:
            resp = requests.post(url, json=payload, timeout=TIMEOUT)
            status, ok = str(resp.status_code), resp.ok
        except requests.RequestException as e:
            status, ok = f"ERROR ({type(e).__name__})", False
        failed += not ok
        print(f"{status:<6} {payload['client']}  ({path.name})")

    print(f"\nDone: {len(files) - failed} sent, {failed} failed")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
