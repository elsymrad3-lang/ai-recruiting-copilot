#!/usr/bin/env python3
"""Verify the recruiting demo data: counts, CV contents, messy cases and test questions.

Reads the real files (PDF text via pypdf, Word text via python-docx) and checks them
against data_key.json, so a broken generator or a hand-edited file shows up here.
Exit code 0 = all checks passed.
"""
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parent
sys.stdout.reconfigure(encoding="utf-8")

ROLES = ["GHL Automation Specialist", "Executive Virtual Assistant", "QuickBooks Bookkeeper",
         "Social Media Manager", "Customer Support Agent"]
EMAIL_RE = re.compile(r"^elsymrad3\+cand\d{2}[a-z]?@gmail\.com$")
PAY_RE = re.compile(r"\$\s?\d|\b\d+\s?(usd|dollars|k a year|per hour|an hour|/hr)\b", re.I)

results = []


def check(section, ok, label, detail=""):
    results.append((section, bool(ok), label, detail))


def norm(s):
    return re.sub(r"\s+", " ", s).strip()


def pdf_text(path):
    reader = PdfReader(path)
    images = sum(len(p.images) for p in reader.pages)
    return norm(" ".join(p.extract_text() or "" for p in reader.pages)), images


def docx_text(path):
    body = Document(path).element.body
    return norm(" ".join("".join(t.text or "" for t in p.iter(qn("w:t"))) for p in body.iter(qn("w:p")))), 0


def matches(f, c):
    return (c.get("role") in (None, f["role"])
            and set(c.get("skills_all", [])) <= set(f["skills"])
            and set(c.get("languages_all", [])) <= set(f["languages"])
            and set(c.get("certs_all", [])) <= set(f["certifications"])
            and set(c.get("flags_all", [])) <= set(f["flags"])
            and f["years_experience"] >= c.get("min_years", 0)
            and f["ghl_years"] >= c.get("min_ghl_years", 0))


def latest_readable(files, include_scanned=False):
    latest = {}
    for f in files:
        if f["messy_case"] == "scanned_no_text_layer" and not include_scanned:
            continue
        if f["person_id"] not in latest or f["last_updated"] > latest[f["person_id"]]["last_updated"]:
            latest[f["person_id"]] = f
    return latest


def main():
    key = json.loads((ROOT / "data_key.json").read_text(encoding="utf-8"))
    questions = json.loads((ROOT / "test_questions.json").read_text(encoding="utf-8"))
    files = key["files"]
    by_file = {f["file"]: f for f in files}

    # ---- transcripts
    S = "Transcripts"
    tdir = ROOT / "transcripts"
    txts = sorted(tdir.glob("*.txt"))
    check(S, len(txts) == 5, "exactly 5 .txt transcripts", f"found {len(txts)}")
    all_t = {p.name: p.read_text(encoding="utf-8") for p in txts}
    for role in ROLES:
        hits = [n for n, t in all_t.items() if role.lower() in t.lower()[:400]]  # role named in the header
        check(S, len(hits) == 1, f"one transcript for '{role}'", ", ".join(hits) or "none")
    for name, t in all_t.items():
        fillers = len(re.findall(r"\b(um+|uh+|like|you know|sorry)\b", t, re.I))
        check(S, fillers >= 5, f"{name}: has filler talk", f"{fillers} filler words")
        check(S, not PAY_RE.search(t), f"{name}: pay never stated (missing detail)")
    for t in key["transcripts"]:
        check(S, t["changed_mind"] and t["missing"], f"{Path(t['file']).name}: key lists changed mind + missing details",
              f"changed: {t['changed_mind']} | missing: {', '.join(t['missing'])}")

    # ---- folder and key counts
    S = "CV counts"
    on_disk = sorted(p.name for p in (ROOT / "cvs").iterdir())
    exts = Counter(Path(n).suffix.lower() for n in on_disk)
    check(S, len(on_disk) == 60, "exactly 60 files in cvs/", f"found {len(on_disk)}")
    check(S, set(exts) == {".pdf", ".docx"}, "only PDF and Word files, both present",
          f"{exts['.pdf']} PDF, {exts['.docx']} DOCX")
    check(S, sorted(Path(f["file"]).name for f in files) == on_disk, "data_key.json lists every file exactly once")
    people = {f["person_id"] for f in files}
    names = {f["full_name"] for f in files}
    check(S, len(people) == 55 and len(names) == 55, "exactly 55 distinct people", f"{len(people)} ids, {len(names)} names")
    layouts = Counter(f["layout"] for f in files)
    check(S, len(layouts) >= 6, "varied layouts", ", ".join(f"{k}:{v}" for k, v in sorted(layouts.items())))

    # ---- per-file contents
    S = "CV contents"
    texts, scanned = {}, []
    problems = defaultdict(list)
    for f in files:
        path = ROOT / f["file"]
        text, images = pdf_text(path) if path.suffix == ".pdf" else docx_text(path)
        texts[f["file"]] = text
        if not text:
            scanned.append((f["file"], images))
            continue
        m = re.search(r"Last updated:\s*(\d{4}-\d{2}-\d{2})", text)
        if not m or m.group(1) != f["last_updated"]:
            problems["Last updated date (YYYY-MM-DD) matches key"].append(f["file"])
        if f["full_name"] not in text and f["full_name"].upper() not in text:
            problems["full name present"].append(f["file"])
        if f["email"] not in text or not EMAIL_RE.match(f["email"]):
            problems["email present and elsymrad3+candNN@gmail.com style"].append(f["file"])
        if re.sub(r"\D", "", f["phone"]) not in re.sub(r"\D", "", text):
            problems["phone present"].append(f["file"])
        missing_sk = [s for s in f["skills_as_written"] if norm(s) not in text]
        if missing_sk or len(f["skills"]) < 5:
            problems["skills present"].append(f"{f['file']} {missing_sk}")
        if not all(lang in text for lang in f["languages"]):
            problems["languages present"].append(f["file"])
        if len(f["jobs"]) < 2 or not all(j["start"][:4] in text for j in f["jobs"]):
            problems["job history with dates"].append(f["file"])
    readable = len(files) - len(scanned)
    for label in ["Last updated date (YYYY-MM-DD) matches key", "full name present",
                  "email present and elsymrad3+candNN@gmail.com style", "phone present", "skills present",
                  "languages present", "job history with dates"]:
        bad = problems[label]
        check(S, not bad, f"{label}: {readable - len(bad)}/{readable} readable CVs", "; ".join(bad[:5]))

    # ---- messy cases
    S = "Messy cases"
    scan_keyed = [f["file"] for f in files if f["messy_case"] == "scanned_no_text_layer"]
    check(S, len(scanned) == 1 and [s[0] for s in scanned] == scan_keyed and scanned[0][1] >= 1
          and scan_keyed[0].endswith(".pdf"),
          "1 scanned CV: PDF with a page image and no text layer", f"{scanned}")
    groups = defaultdict(list)
    for f in files:
        groups[f["person_id"]].append(f)
    multi = {pid: fs for pid, fs in groups.items() if len(fs) > 1}
    dups = {pid: fs for pid, fs in multi.items() if len({f["last_updated"] for f in fs}) == 1}
    upds = {pid: fs for pid, fs in multi.items() if len({f["last_updated"] for f in fs}) > 1}
    check(S, all(len(fs) == 2 for fs in multi.values()) and len(multi) == 5,
          "5 people have 2 files (3 duplicates + 2 updated) -> 55 + 5 = 60")
    check(S, len(dups) == 3, "3 duplicate pairs with the same Last updated date",
          "; ".join(f"{fs[0]['full_name']} ({fs[0]['last_updated']})" for fs in dups.values()))

    def dup_text_facts(fs):
        a, b = (texts[f["file"]] for f in fs)
        ea, eb = (re.search(r"elsymrad3\+\S+@gmail\.com", t).group(0) for t in (a, b))
        same_phone = re.sub(r"\D", "", fs[0]["phone"]) in re.sub(r"\D", "", a) and \
            re.sub(r"\D", "", fs[1]["phone"]) in re.sub(r"\D", "", b) and fs[0]["phone"] == fs[1]["phone"]
        return ea != eb and same_phone

    diff_email = [fs for fs in dups.values() if dup_text_facts(fs)]
    check(S, len(diff_email) == 1, "one duplicate pair: different emails, same phone (read from the files)",
          ", ".join(f"{f['file']} <{f['email']}>" for fs in diff_email for f in fs))
    mixed = [fs for fs in dups.values() if {f["format"] for f in fs} == {"pdf", "docx"}]
    check(S, len(mixed) == 1, "one duplicate pair: one PDF + one Word file",
          ", ".join(f["file"] for fs in mixed for f in fs))
    check(S, all(f["messy_case"].startswith("duplicate") for fs in dups.values() for f in fs),
          "duplicate files are labelled as duplicates in data_key.json")
    check(S, len(upds) == 2, "2 updated-CV pairs (older + newer date)",
          "; ".join(f"{fs[0]['full_name']}" for fs in upds.values()))
    for fs in upds.values():
        old, new = sorted(fs, key=lambda f: f["last_updated"])
        added = sorted(set(new["skills"]) - set(old["skills"]))
        added_in_text = [s for s, w in zip(new["skills"], new["skills_as_written"])
                         if s in added and norm(w) in texts[new["file"]] and norm(w) not in texts[old["file"]]]
        more_years = new["years_experience"] > old["years_experience"]
        check(S, old["messy_case"] == "updated_older" and new["messy_case"] == "updated_newer"
              and (added_in_text or more_years),
              f"{new['full_name']}: newer CV adds skills or years",
              f"{old['last_updated']} -> {new['last_updated']}; years {old['years_experience']} -> "
              f"{new['years_experience']}; added {added_in_text}")

    latest = latest_readable(files, include_scanned=True)
    both = [f["full_name"] for f in latest.values() if f["ghl_years"] >= 3 and "Arabic" in f["languages"]]
    ghl_only = [f["full_name"] for f in latest.values() if f["ghl_years"] >= 3 and "Arabic" not in f["languages"]]
    ar_only = [f["full_name"] for f in latest.values() if f["ghl_years"] < 3 and "Arabic" in f["languages"]]
    check(S, both and ghl_only and ar_only, "some have 3+ yrs GHL AND Arabic, some only one of the two",
          f"both: {len(both)}, GHL 3+ only: {len(ghl_only)}, Arabic only: {len(ar_only)}")
    all_text = " ".join(texts.values())
    spell = {s: len(re.findall(rf"\b{s}\b", all_text)) for s in ["GHL", "GoHighLevel", "Go High Level", "HighLevel"]}
    check(S, spell["GHL"] and spell["GoHighLevel"] and spell["Go High Level"],
          "GHL written differently on purpose (GHL / GoHighLevel / Go High Level)",
          ", ".join(f"'{k}' x{v}" for k, v in spell.items()))

    # ---- questions
    S = "Test questions"
    qtypes = key["test_question_types"]
    check(S, len(questions) == 10 and len(qtypes) == 10, "exactly 10 questions")
    check(S, all(set(q) == {"question", "expected", "why"} for q in questions), "format {question, expected, why}")
    tc = Counter(t["type"] for t in qtypes)
    want = {"normal": 4, "impossible": 2, "needs_newer_cv": 1, "duplicate_person": 1, "scanned_only": 1,
            "skill_plus_language": 1}
    check(S, tc == Counter(want), "mix: 4 normal, 2 impossible, 1 newer-CV, 1 duplicate, 1 scanned, 1 skill+language",
          ", ".join(f"{k}:{v}" for k, v in tc.items()))
    readable_latest = latest_readable(files)
    for q, t in zip(questions, qtypes):
        exp, crit, i = q["expected"], t["criteria"], t["index"]
        got = sorted(f["full_name"] for f in readable_latest.values() if matches(f, crit))
        label = f"Q{i} [{t['type']}]"
        if exp == "NO_MATCH":
            check(S, not got, f"{label}: NO_MATCH confirmed against the key", q["question"])
        else:
            ok = isinstance(exp, list) and 1 <= len(exp) <= 5 and len(set(exp)) == len(exp) \
                and set(exp) <= names and sorted(exp) == got
            check(S, ok, f"{label}: {len(exp)} expected (<=5), matches key", ", ".join(exp))
            # the facts must also be visible in the CV text itself
            missing = []
            for name in exp:
                f = next(v for v in readable_latest.values() if v["full_name"] == name)
                written = dict(zip(f["skills"], f["skills_as_written"]))
                for s in crit.get("skills_all", []):
                    if norm(written[s]) not in texts[f["file"]]:
                        missing.append(f"{name}: {s}")
                for lang in crit.get("languages_all", []):
                    if lang not in texts[f["file"]]:
                        missing.append(f"{name}: {lang}")
            check(S, not missing, f"{label}: required skills/languages visible in each expected CV", "; ".join(missing))
        if t["type"] == "needs_newer_cv":
            only_new = [n for n in (exp if isinstance(exp, list) else [])
                        if n in {fs[0]["full_name"] for fs in upds.values()}
                        and not any(matches(f, crit) for f in groups[next(p for p, fs in upds.items()
                                                                          if fs[0]["full_name"] == n)]
                                    if f["messy_case"] == "updated_older")]
            check(S, only_new, f"{label}: an expected person qualifies only via the newer CV", ", ".join(only_new))
        if t["type"] == "duplicate_person":
            dup_names = {fs[0]["full_name"] for fs in dups.values()}
            hit = [n for n in exp if n in dup_names] if isinstance(exp, list) else []
            check(S, hit and exp.count(hit[0]) == 1, f"{label}: duplicated person included exactly once", ", ".join(hit))
        if t["type"] == "scanned_only":
            scan_name = by_file[scan_keyed[0]]["full_name"]
            with_scan = [f["full_name"] for f in latest_readable(files, include_scanned=True).values() if matches(f, crit)]
            check(S, with_scan == [scan_name] and (exp == "NO_MATCH" or scan_name not in exp),
                  f"{label}: only the scanned CV answers it, and expected excludes that person",
                  f"scanned person: {scan_name}; expected: {exp}")
        if t["type"] == "skill_plus_language":
            check(S, crit.get("skills_all") and crit.get("languages_all"), f"{label}: combines a skill and a language")

    # ---- report
    width = max(len(r[2]) for r in results) + 2
    current = None
    for section, ok, label, detail in results:
        if section != current:
            print(f"\n== {section} ==")
            current = section
        print(f"  [{'PASS' if ok else 'FAIL'}] {label.ljust(width)}{('  ' + detail) if detail else ''}")
    failed = [r for r in results if not r[1]]
    print(f"\n{len(results) - len(failed)}/{len(results)} checks passed"
          f" | {len(files)} files, {len(people)} people, {exts['.pdf']} PDF / {exts['.docx']} DOCX")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
