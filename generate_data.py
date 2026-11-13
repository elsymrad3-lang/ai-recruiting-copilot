#!/usr/bin/env python3
"""Generate the fictional recruiting demo data set.

Writes cvs/ (60 files for 55 people), data_key.json and test_questions.json.
The call transcripts in transcripts/ are hand-written and not touched here.
Seeds are fixed, so re-running produces the same data.
All people, companies and schools are fictional.
"""
import html
import json
import random
import shutil
import textwrap
import unicodedata
from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
from PIL import Image, ImageDraw, ImageFilter, ImageFont
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (BaseDocTemplate, Frame, FrameBreak, HRFlowable, PageTemplate,
                                Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle)

ROOT = Path(__file__).resolve().parent
CV_DIR = ROOT / "cvs"
FONT_DIR = Path(r"C:\Windows\Fonts")
EMAIL = "elsymrad3+cand{:02d}{}@gmail.com"

for fam, (reg, bold) in {"Arial": ("arial.ttf", "arialbd.ttf"),
                         "Georgia": ("georgia.ttf", "georgiab.ttf"),
                         "Calibri": ("calibri.ttf", "calibrib.ttf")}.items():
    pdfmetrics.registerFont(TTFont(fam, str(FONT_DIR / reg)))
    pdfmetrics.registerFont(TTFont(fam + "-Bold", str(FONT_DIR / bold)))

# ---------------------------------------------------------------- reference data

COUNTRIES = {
    "Egypt": ([("Arabic", "Native"), ("English", "Fluent")],
              lambda r: f"+20 1{r.choice('0125')}{r.randint(0, 9)} {r.randint(100, 999)} {r.randint(1000, 9999)}"),
    "Jordan": ([("Arabic", "Native"), ("English", "Fluent")],
               lambda r: f"+962 7{r.choice('789')} {r.randint(100, 999)} {r.randint(1000, 9999)}"),
    "Lebanon": ([("Arabic", "Native"), ("English", "Fluent"), ("French", "Intermediate")],
                lambda r: f"+961 {r.choice(['3', '70', '71', '76'])} {r.randint(100, 999)} {r.randint(100, 999)}"),
    "Morocco": ([("Arabic", "Native"), ("French", "Fluent"), ("English", "Professional")],
                lambda r: f"+212 6{r.randint(10, 99)}-{r.randint(100000, 999999)}"),
    "Philippines": ([("Filipino", "Native"), ("English", "Fluent")],
                    lambda r: f"+63 9{r.randint(10, 99)} {r.randint(100, 999)} {r.randint(1000, 9999)}"),
    "Colombia": ([("Spanish", "Native"), ("English", "C1")],
                 lambda r: f"+57 3{r.randint(0, 2)}{r.randint(0, 9)} {r.randint(100, 999)} {r.randint(1000, 9999)}"),
    "Mexico": ([("Spanish", "Native"), ("English", "C1")],
               lambda r: f"+52 {r.randint(33, 99)} {r.randint(1000, 9999)} {r.randint(1000, 9999)}"),
    "Brazil": ([("Portuguese", "Native"), ("English", "Fluent"), ("Spanish", "Basic")],
               lambda r: f"+55 {r.randint(11, 99)} 9{r.randint(1000, 9999)}-{r.randint(1000, 9999)}"),
    "Kenya": ([("Swahili", "Native"), ("English", "Fluent")],
              lambda r: f"+254 7{r.randint(10, 99)} {r.randint(100, 999)} {r.randint(100, 999)}"),
    "South Africa": ([("English", "Native"), ("Zulu", "Native"), ("Afrikaans", "Conversational")],
                     lambda r: f"+27 {r.choice(['72', '73', '82', '83'])} {r.randint(100, 999)} {r.randint(1000, 9999)}"),
    "Pakistan": ([("Urdu", "Native"), ("English", "Fluent"), ("Punjabi", "Conversational")],
                 lambda r: f"+92 3{r.randint(0, 4)}{r.randint(0, 9)} {r.randint(1000000, 9999999)}"),
    "India": ([("Hindi", "Fluent"), ("English", "Fluent")],
              lambda r: f"+91 9{r.randint(1000, 9999)} {r.randint(10000, 99999)}"),
    "Serbia": ([("Serbian", "Native"), ("English", "C1")],
               lambda r: f"+381 6{r.randint(0, 6)} {r.randint(100, 999)} {r.randint(1000, 9999)}"),
}

ROLE_TITLE = {"ghl": "GHL Automation Specialist", "eva": "Executive Virtual Assistant",
              "bookkeeper": "QuickBooks Bookkeeper", "social": "Social Media Manager",
              "support": "Customer Support Agent"}

# Skills that questions hinge on are never in these pools; they are only assigned on purpose.
POOLS = {
    "ghl": ["Zapier", "Funnel building", "Workflow automation", "Webhooks & API integrations",
            "Twilio / LC Phone", "Mailgun email setup", "CRM pipeline design", "Stripe", "Calendly",
            "Google Sheets", "ClickFunnels", "WordPress"],
    "eva": ["Google Workspace", "Microsoft 365", "Calendar management", "Inbox management",
            "Travel booking", "Notion", "Asana", "Slack", "Zoom", "Expensify", "Meeting minutes",
            "Trello", "ClickUp", "Canva", "Data entry"],
    "bookkeeper": ["Xero", "Bank reconciliations", "Accounts payable", "Accounts receivable",
                   "Month-end close", "Excel", "Sales tax filing", "1099 preparation", "Dext",
                   "Hubdoc", "Bill.com", "Gusto payroll", "Shopify payouts reconciliation"],
    "social": ["Canva", "Meta Business Suite", "CapCut", "Later", "Buffer", "Hootsuite",
               "Adobe Premiere Pro", "Instagram Reels", "TikTok", "LinkedIn", "Content calendars",
               "Copywriting", "Community management", "Google Analytics", "Meta Ads"],
    "support": ["Zendesk", "Freshdesk", "Intercom", "Help Scout", "Shopify", "Live chat",
                "Email support", "Phone support", "Refunds & returns", "Macros", "Slack", "Jira",
                "CSAT tracking"],
}
CORE = {"ghl": ["GoHighLevel"], "bookkeeper": ["QuickBooks Online"]}
RESERVED = {"Make.com", "Keap", "Gorgias", "Kustomer", "TikTok Ads", "NetSuite"}

SPELLINGS = {
    "Make.com": ["Make.com", "Make (Integromat)", "Integromat / Make"],
    "QuickBooks Online": ["QuickBooks Online", "QBO", "Quickbooks Online (QBO)"],
    "Google Workspace": ["Google Workspace", "G Suite", "GSuite"],
    "Microsoft 365": ["Microsoft 365", "MS Office 365"],
    "Meta Business Suite": ["Meta Business Suite", "Facebook Business Manager"],
    "Zendesk": ["Zendesk", "ZenDesk"],
    "Keap": ["Keap (Infusionsoft)"],
    "Canva": ["Canva", "Canva Pro"],
    "TikTok Ads": ["TikTok Ads", "TikTok Ads Manager"],
    "Excel": ["Excel (pivot tables, VLOOKUP)", "MS Excel"],
}
GHL_SPELLINGS = ["GHL", "GoHighLevel", "Go High Level", "HighLevel"]

PREFIXES = ["Corvane", "Tallowbrook", "Bluepeak", "Marrowgate", "Quillon", "Sablewood", "Fennix",
            "Orchardine", "Halvard", "Juniperra", "Kestrelton", "Lumenwick", "Nettlefield", "Ostravia",
            "Pinecrest Hollow", "Redfernly", "Silverbirch", "Tamsinvale", "Umberleigh", "Vellumark",
            "Wrenfield", "Yarrowby", "Zephyrine", "Callowmere", "Driftmoor", "Emberlyn"]
SUFFIXES = {
    "ghl": ["Digital", "Growth Agency", "Med Spa", "Roofing Co.", "Realty Group", "Coaching"],
    "eva": ["Capital Partners", "Ventures", "Advisory", "Legal Group", "Holdings", "Consulting"],
    "bookkeeper": ["Bookkeeping Services", "Accounting Outsourcing", "Ledger Co.", "Tax & Books",
                   "Finance Hub", "Outsourcing Inc."],
    "social": ["Skincare", "Fitness Studio", "Apparel", "Creative Agency", "Coffee Co.", "Pet Supply"],
    "support": ["Home Goods", "Outfitters", "Subscription Box", "Electronics", "Beauty", "Support Center"],
    "pre": ["Retail", "Trading", "Call Center Services", "Logistics", "Foods", "Printing"],
}
CUR_TITLES = {
    "ghl": ["{ghl} Automation Specialist", "Marketing Automation Specialist", "{ghl} Workflow Builder"],
    "eva": ["Executive Virtual Assistant", "Remote Executive Assistant", "Virtual Executive Assistant"],
    "bookkeeper": ["Bookkeeper", "Remote Bookkeeper", "Senior Bookkeeper"],
    "social": ["Social Media Manager", "Content & Social Media Manager"],
    "support": ["Customer Support Agent", "Customer Experience Associate", "Support Specialist"],
}
PREV_TITLES = {
    "ghl": ["CRM Automation Assistant", "Junior {ghl} Specialist"],
    "eva": ["Virtual Assistant", "Administrative Assistant"],
    "bookkeeper": ["Junior Bookkeeper", "Bookkeeping Assistant"],
    "social": ["Social Media Coordinator", "Content Assistant"],
    "support": ["Customer Support Representative", "Live Chat Agent"],
}
PRE_TITLES = ["Sales Associate", "Data Entry Clerk", "Receptionist", "Retail Cashier", "Office Assistant"]
HEADLINES = {
    "ghl": ["{ghl} Automation Specialist", "CRM & Automation Specialist", "Marketing Automation Specialist"],
    "eva": ["Executive Virtual Assistant", "Executive Assistant", "Remote Executive Assistant"],
    "bookkeeper": ["Bookkeeper", "QuickBooks Bookkeeper", "Remote Bookkeeper"],
    "social": ["Social Media Manager", "Social Media & Content Manager"],
    "support": ["Customer Support Agent", "Customer Success Representative", "E-commerce Support Specialist"],
}
BULLETS = {
    "ghl": ["Built and maintained {n} {ghl} workflows for appointment reminders and lead nurture",
            "Set up sub-accounts, pipelines and calendars for {n2} agency clients",
            "Designed funnels and landing pages that lifted booked calls by {p}%",
            "Connected forms, SMS and email sequences using webhooks",
            "Cleaned up duplicate contacts and rebuilt the tagging structure"],
    "eva": ["Managed inbox and calendar for {exec}",
            "Prepared meeting agendas, minutes and follow-up tasks",
            "Coordinated vendor calls and kept shared drives organised",
            "Tracked expenses and prepared monthly reports",
            "Wrote team SOPs in a shared knowledge base"],
    "bookkeeper": ["Reconciled {n2} bank and credit card accounts every month",
                   "Handled payables and receivables for {n2} small-business clients",
                   "Prepared month-end close packages and P&L reports",
                   "Cleaned up 18 months of backlog bookkeeping for a retail client",
                   "Coded receipts and categorised transactions"],
    "social": ["Planned and scheduled {n} posts per month across Instagram, Facebook and LinkedIn",
               "Created Reels and short videos for product launches",
               "Grew a brand account from {a}k to {b}k followers",
               "Answered comments and DMs within two hours",
               "Wrote captions and monthly content calendars"],
    "support": ["Handled {n}+ tickets per day via email and live chat",
                "Kept CSAT above {p2}%",
                "Processed refunds, returns and order changes",
                "Wrote and updated macros and help-centre articles",
                "Escalated bugs to the product team with clear reproduction steps"],
}
SPECIAL_BULLETS = {
    "Make.com": "Built multi-step scenarios in {make} to sync {ghl} with Google Sheets and Slack",
    "Keap": "Migrated 12 clients from Keap (Infusionsoft) to {ghl}, including tags and campaigns",
    "Gorgias": "Worked in Gorgias for a Shopify skincare store, using macros and order actions",
    "Kustomer": "Supported a subscription brand on Kustomer, covering chat and email queues",
    "TikTok Ads": "Launched and optimised TikTok Ads campaigns for product drops",
}
SCHOOLS = ["Northgate Institute of Business", "Lakeshore Polytechnic", "Al-Noor College of Commerce",
           "Riverside State College", "Eastvale University", "Cedar Ridge Technical Institute",
           "Harborview College", "St. Aldric University"]
DEGREES = {"ghl": "B.Sc. Information Technology", "eva": "B.A. Business Administration",
           "bookkeeper": "B.S. Accountancy", "social": "B.A. Mass Communication",
           "support": "B.A. English"}

HEADINGS = {"summary": ["Summary", "Profile", "About Me", "Professional Summary"],
            "skills": ["Skills", "Core Skills", "Tools & Skills", "Technical Skills"],
            "experience": ["Experience", "Work Experience", "Employment History", "Professional Experience"],
            "education": ["Education", "Academic Background"],
            "languages": ["Languages", "Language Skills"],
            "certifications": ["Certifications", "Licenses & Certifications"]}
MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTH = ["January", "February", "March", "April", "May", "June", "July", "August", "September",
         "October", "November", "December"]

# ---------------------------------------------------------------- the 55 people
# (num, full name, role, country, city, years in role, extras)
# extras: add/remove skills, certs, flags, ghl_years (for non-GHL roles), start (fixed role start)
PEOPLE = [
    (1, "Karim Abdelrahman", "ghl", "Egypt", "Cairo", 4, {}),
    (2, "Lina Haddad", "ghl", "Jordan", "Amman", 5, {"add": ["Make.com"]}),
    (3, "Youssef Benali", "ghl", "Morocco", "Casablanca", 6, {}),
    (4, "Joanna Dimaculangan", "ghl", "Philippines", "Cebu City", 4, {"add": ["Make.com"]}),
    (5, "Mark Anthony Villanueva", "ghl", "Philippines", "Davao City", 5, {}),
    (6, "Santiago Restrepo", "ghl", "Colombia", "Medellín", 4, {}),
    (7, "Nourhan Sabry", "ghl", "Egypt", "Alexandria", 1, {}),
    (8, "Tariq Al-Masri", "ghl", "Jordan", "Irbid", 2, {}),
    (9, "Bilal Qureshi", "ghl", "Pakistan", "Lahore", None, {"start": (2023, 1), "nosplit": True}),
    (10, "Rhea Macaraeg", "ghl", "Philippines", "Baguio City", 2, {"add": ["Keap"]}),
    (11, "Rami Khoury", "ghl", "Lebanon", "Beirut", 7, {}),
    (12, "Aditi Raghavan", "ghl", "India", "Pune", 2, {}),
    (13, "Camila Ortega", "eva", "Mexico", "Guadalajara", 5, {"flags": ["supported_ceo"], "add": ["Travel booking"]}),
    (14, "Mariam Ezzat", "eva", "Egypt", "Giza", 3, {"ghl_years": 1}),
    (15, "Patricia Lumbao", "eva", "Philippines", "Quezon City", 4, {"ghl_years": 1}),
    (16, "Hannah Mwangi", "eva", "Kenya", "Nairobi", 6, {"flags": ["supported_ceo"], "add": ["Travel booking"]}),
    (17, "Ana Beatriz Rocha", "eva", "Brazil", "Curitiba", 3, {}),
    (18, "Salma Tawfik", "eva", "Egypt", "Cairo", 4, {}),
    (19, "Danica Petrović", "eva", "Serbia", "Novi Sad", 5, {"flags": ["supported_ceo"], "remove": ["Travel booking"]}),
    (20, "Dana Al-Zoubi", "eva", "Jordan", "Amman", 2, {}),
    (21, "Kristine Joy Bautista", "eva", "Philippines", "Iloilo City", 7, {"flags": ["supported_ceo"], "add": ["Travel booking"]}),
    (22, "Valeria Montoya", "eva", "Colombia", "Bogotá", 3, {}),
    (23, "Zainab Farooq", "eva", "Pakistan", "Karachi", 2, {}),
    (24, "Ma. Teresa Gonzaga", "bookkeeper", "Philippines", "Makati City", 7, {"certs": ["QuickBooks ProAdvisor"]}),
    (25, "Hesham Ragab", "bookkeeper", "Egypt", "Mansoura", 4, {}),
    (26, "Thandiwe Nkosi", "bookkeeper", "South Africa", "Durban", 5, {}),
    (27, "Rowena Castillo", "bookkeeper", "Philippines", "Pasig City", 6, {"certs": ["QuickBooks ProAdvisor"]}),
    (28, "Nadia Bouzid", "bookkeeper", "Morocco", "Rabat", 3, {}),
    (29, "Arjun Mehta", "bookkeeper", "India", "Ahmedabad", 8, {}),
    (30, "Grace Achieng", "bookkeeper", "Kenya", "Kisumu", 3, {"certs": ["QuickBooks ProAdvisor"]}),
    (31, "Jerome Santos", "bookkeeper", "Philippines", "Santa Rosa, Laguna", None,
     {"start": (2021, 3), "nosplit": True, "remove": ["Xero", "Gusto payroll"]}),
    (32, "Layal Nassar", "bookkeeper", "Lebanon", "Jounieh", 4, {}),
    (33, "Imran Siddiqui", "bookkeeper", "Pakistan", "Islamabad", 9, {"certs": ["QuickBooks ProAdvisor"]}),
    (34, "Lucía Fernández", "bookkeeper", "Mexico", "Monterrey", 2, {}),
    (35, "Bea Soriano", "social", "Philippines", "Manila", 3, {}),
    (36, "Daniela Quintero", "social", "Colombia", "Cali", 4, {"add": ["CapCut", "TikTok Ads"]}),
    (37, "Farah Mansour", "social", "Jordan", "Zarqa", 2, {"ghl_years": 1}),
    (38, "Kevin Otieno", "social", "Kenya", "Mombasa", 3, {}),
    (39, "Isabela Duarte", "social", "Brazil", "Recife", 5, {}),
    (40, "Nicole Dela Cruz", "social", "Philippines", "Taguig City", 4, {"add": ["CapCut", "TikTok Ads"], "ghl_years": 2}),
    (41, "Yasmine El Idrissi", "social", "Morocco", "Marrakesh", 3, {}),
    (42, "Priya Nair", "social", "India", "Kochi", 4, {"langs": [("Malayalam", "Native")]}),
    (43, "Mateo Salazar", "social", "Mexico", "Puebla", 2, {}),
    (44, "Hiba Saad", "social", "Lebanon", "Tripoli", 3, {"add": ["CapCut", "TikTok Ads"]}),
    (45, "Sipho Dlamini", "social", "South Africa", "Johannesburg", 6, {}),
    (46, "Jessa Mae Tolentino", "support", "Philippines", "Imus, Cavite", 3, {}),
    (47, "Ahmed Shalaby", "support", "Egypt", "Cairo", 4, {"add": ["Gorgias", "Shopify"]}),
    (48, "Carlos Mejía", "support", "Colombia", "Barranquilla", 2, {}),
    (49, "Aisha Rahman", "support", "Pakistan", "Faisalabad", 3, {"add": ["Kustomer"]}),
    (50, "Paolo Navarro", "support", "Philippines", "Bacolod City", 5, {"add": ["Gorgias", "Shopify"]}),
    (51, "Wanjiru Kamau", "support", "Kenya", "Nakuru", 2, {}),
    (52, "Omar Haddadin", "support", "Jordan", "Amman", 4, {"add": ["Kustomer"]}),
    (53, "Fernanda Lima", "support", "Brazil", "Belo Horizonte", 3, {"add": ["Gorgias", "Shopify"]}),
    (54, "Rania Aoun", "support", "Lebanon", "Zahlé", 2, {}),
    (55, "Diego Herrera", "support", "Mexico", "Tijuana", 6, {}),
]

# Files per person. Default is one file; these people get special treatment.
# Each version: (messy_case, last_updated, format, layout, overrides)
SPECIAL_FILES = {
    9: [("updated_older", date(2025, 2, 10), "pdf", "classic", {}),
        ("updated_newer", date(2026, 8, 15), "docx", "table", {"add": ["Make.com", "Twilio / LC Phone"]})],
    31: [("updated_older", date(2024, 6, 20), "pdf", "compact", {}),
         ("updated_newer", date(2026, 5, 12), "pdf", "sidebar", {"add": ["Xero", "Gusto payroll"]})],
    18: [("duplicate_diff_email_same_phone", date(2026, 3, 3), "pdf", "banner", {}),
         ("duplicate_diff_email_same_phone", date(2026, 3, 3), "pdf", "banner", {"email_suffix": "b"})],
    52: [("duplicate_pdf_and_docx", date(2026, 7, 21), "pdf", "classic", {}),
         ("duplicate_pdf_and_docx", date(2026, 7, 21), "docx", "plain", {})],
    42: [("duplicate_identical_copy", date(2025, 11, 30), "docx", "modern", {}),
         ("duplicate_identical_copy", date(2025, 11, 30), "docx", "modern", {"copy_of_previous": True})],
    10: [("scanned_no_text_layer", date(2026, 1, 19), "pdf", "scanned", {})],
}
SPECIAL_NAMES = {
    9: ["Bilal_Qureshi_CV_2025.pdf", "Bilal Qureshi CV - updated.docx"],
    31: ["Jerome Santos Resume.pdf", "Jerome_Santos_Resume_2026.pdf"],
    18: ["Resume - Salma Tawfik.pdf", "SalmaTawfik_resume.pdf"],
    52: ["Omar Haddadin CV.pdf", "Omar Haddadin CV.docx"],
    42: ["Priya_Nair_CV.docx", "Priya_Nair_CV (1).docx"],
    10: ["scan_0047.pdf"],
}

# ---------------------------------------------------------------- helpers


def add_months(ym, n):
    t = ym[0] * 12 + ym[1] - 1 + n
    return (t // 12, t % 12 + 1)


def months_between(ym, d):
    return d.year * 12 + d.month - 1 - (ym[0] * 12 + ym[1] - 1)


def ascii_slug(name):
    s = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode()
    return s.replace(".", "").replace("-", " ").split()


def esc(s):
    return html.escape(s, quote=False)


def default_filename(num, name, fmt, lu):
    parts = ascii_slug(name)
    first, last = parts[0], parts[-1]
    pats = [f"{first}_{last}_CV", f"Resume - {first} {last}", f"{first.lower()}.{last.lower()}.resume",
            f"CV {last} {first} {lu.year}", f"{last}_{first[0]}_cv_final", f"{first}{last}-Resume"]
    return f"{pats[num % len(pats)]}.{fmt}"


class Person:
    """Stable per-person data: everything that is the same in all of a person's files."""

    def __init__(self, spec, ghl_spelling):
        self.num, self.name, self.role, self.country, self.city, self.years, self.extra = spec
        self.id = f"cand{self.num:02d}"
        r = random.Random(self.num * 7919)
        self.phone = COUNTRIES[self.country][1](r)
        self.languages = COUNTRIES[self.country][0] + self.extra.get("langs", [])
        self.ghl_spelling = ghl_spelling
        rs = random.Random(self.num * 104729)
        self.spell = {k: rs.choice(v) for k, v in SPELLINGS.items()}
        self.date_style = rs.randrange(4)
        self.lang_style = rs.randrange(3)
        self.headings = {k: rs.choice(v) for k, v in HEADINGS.items()}
        self.lu_default = date.fromordinal(
            date(2024, 4, 1).toordinal() + rs.randrange((date(2026, 10, 25) - date(2024, 4, 1)).days))
        self.fmt_default = "pdf" if rs.random() < 0.5 else "docx"
        self.layout_default = rs.choice(["classic", "sidebar", "banner", "compact"] if self.fmt_default == "pdf"
                                        else ["plain", "table", "modern"])
        self.school = rs.choice(SCHOOLS)

    def version(self, lu, overrides):
        """Facts for one file of this person, as of its Last updated date."""
        r = random.Random(self.num * 7919 + 1)
        role, ex = self.role, self.extra
        skills = CORE.get(role, []) + r.sample(POOLS[role], r.randint(5, 8))
        for s in ex.get("add", []) + overrides.get("add", []):
            if s not in skills:
                skills.append(s)
        skills = [s for s in skills if s not in ex.get("remove", []) or s in overrides.get("add", [])]
        if ex.get("ghl_years") and "GoHighLevel" not in skills:
            skills.append("GoHighLevel")
        lu_ym = (lu.year, lu.month)
        start = ex.get("start") or add_months(lu_ym, -(self.years * 12 + r.randint(1, 8)))
        years = months_between(start, lu) // 12
        ghl_years = years if role == "ghl" else ex.get("ghl_years", 0)

        # jobs, newest first
        comps = [f"{p} {s}" for p, s in zip(r.sample(PREFIXES, 3), r.sample(SUFFIXES[role], 3))]
        pre_comp = f"{r.choice(PREFIXES)} {r.choice(SUFFIXES['pre'])}"
        flags = ex.get("flags", [])
        if "supported_ceo" in flags:
            cur_title = r.choice(["Executive Assistant to the CEO", "Executive Assistant to the Founder"])
            exec_who = "the CEO" if "CEO" in cur_title else "the founder"
        else:
            cur_title = r.choice(CUR_TITLES[role])
            exec_who = r.choice(["a regional sales director", "a real estate team lead", "two law-firm partners"])
        prev_title = r.choice(PREV_TITLES[role])
        total = months_between(start, lu)
        roll, offset = r.random(), r.random()
        fill = dict(ghl=self.ghl_spelling, make=self.spell["Make.com"], exec=exec_who,
                    n=r.randint(25, 90), n2=r.randint(6, 20), p=r.randint(15, 45), p2=r.randint(91, 98),
                    a=r.randint(2, 9), b=r.randint(12, 60))
        cur_bullets = [b.format(**fill) for b in r.sample(BULLETS[role], 3)]
        prev_bullets = [b.format(**fill) for b in r.sample(BULLETS[role], 2)]
        for s in skills:
            if s in SPECIAL_BULLETS:
                cur_bullets.append(SPECIAL_BULLETS[s].format(**fill))
        if "Travel booking" in skills and "supported_ceo" in flags:
            cur_bullets.append(f"Booked international travel and built itineraries for {exec_who}")
        jobs = []
        if total >= 36 and roll < 0.6 and not ex.get("nosplit"):
            split = add_months(start, 12 + int(offset * (total - 24)))
            jobs.append(dict(title=cur_title.format(ghl=self.ghl_spelling), company=comps[0],
                             start=split, end=None, bullets=cur_bullets))
            jobs.append(dict(title=prev_title.format(ghl=self.ghl_spelling), company=comps[1],
                             start=start, end=add_months(split, -1), bullets=prev_bullets))
        else:
            jobs.append(dict(title=cur_title.format(ghl=self.ghl_spelling), company=comps[0],
                             start=start, end=None, bullets=cur_bullets))
        pre_end = add_months(start, -1 - r.randint(0, 3))
        jobs.append(dict(title=r.choice(PRE_TITLES), company=pre_comp,
                         start=add_months(pre_end, -r.randint(10, 30)), end=pre_end,
                         bullets=["Handled customer questions and daily reporting"]))
        grad = jobs[-1]["start"][0] - r.randint(0, 1)

        certs = list(ex.get("certs", []))
        email = EMAIL.format(self.num, overrides.get("email_suffix", ""))
        headline = r.choice(HEADLINES[role]).format(ghl=self.ghl_spelling)
        return dict(person=self, lu=lu, skills=skills, years=years, ghl_years=ghl_years, jobs=jobs,
                    certs=certs, flags=flags, email=email, headline=headline, grad=grad)


# ---------------------------------------------------------------- text rendering


def fmt_ym(ym, style):
    y, m = ym
    return [f"{MON[m - 1]} {y}", f"{m:02d}/{y}", f"{y}-{m:02d}", f"{MONTH[m - 1]} {y}"][style]


def job_dates(job, style):
    end = fmt_ym(job["end"], style) if job["end"] else ["Present", "Present", "Current", "now"][style]
    return fmt_ym(job["start"], style) + [" – ", " - ", " to ", " — "][style] + end


def skill_text(p, v, s):
    if s == "GoHighLevel":
        y = v["ghl_years"]
        yrs = "1 yr" if y == 1 else f"{y} yrs"
        return f"{p.ghl_spelling} ({yrs})" if p.role == "ghl" else f"{p.ghl_spelling} (basic, {yrs})"
    return p.spell.get(s, s)


def cert_text(c):
    return {"QuickBooks ProAdvisor": "Certified QuickBooks ProAdvisor (QuickBooks Online)"}.get(c, c)


def lang_text(p, name, level):
    return [f"{name} ({level})", f"{name} – {level}", f"{name}: {level}"][p.lang_style]


def cv_content(v):
    p = v["person"]
    role_word = {"ghl": "building CRM automations", "eva": "supporting busy executives",
                 "bookkeeper": "keeping small-business books clean", "social": "growing brand accounts",
                 "support": "helping customers of online stores"}[p.role]
    yrs = "1 year" if v["years"] == 1 else f"{v['years']} years"
    return dict(
        name=p.name, headline=v["headline"], location=f"{p.city}, {p.country}", phone=p.phone,
        email=v["email"], last_updated=f"Last updated: {v['lu'].isoformat()}",
        summary=f"{v['headline']} with {yrs} of experience {role_word}. Remote-ready, reliable and detail-oriented.",
        skills=[skill_text(p, v, s) for s in v["skills"]],
        jobs=[dict(title=j["title"], company=j["company"], dates=job_dates(j, p.date_style), bullets=j["bullets"])
              for j in v["jobs"]],
        education=[f"{DEGREES[p.role]}, {p.school} ({v['grad']})"],
        languages=[lang_text(p, n, l) for n, l in p.languages],
        certs=[cert_text(c) for c in v["certs"]],
        h=p.headings,
    )


# ---------------------------------------------------------------- PDF layouts

def styles(font, accent):
    return dict(
        name=ParagraphStyle("n", fontName=font + "-Bold", fontSize=20, leading=24, textColor=accent),
        sub=ParagraphStyle("s", fontName=font, fontSize=11, leading=14, textColor=colors.HexColor("#333333")),
        head=ParagraphStyle("h", fontName=font + "-Bold", fontSize=11.5, leading=14, spaceBefore=9, spaceAfter=3,
                            textColor=accent),
        body=ParagraphStyle("b", fontName=font, fontSize=9.5, leading=13),
        small=ParagraphStyle("sm", fontName=font, fontSize=8.8, leading=12, textColor=colors.HexColor("#444444")),
        bold=ParagraphStyle("bd", fontName=font + "-Bold", fontSize=9.8, leading=13),
        bullet=ParagraphStyle("bl", fontName=font, fontSize=9.3, leading=12.5, leftIndent=11, bulletIndent=2),
    )


def pdf_jobs(c, st):
    out = []
    for j in c["jobs"]:
        out.append(Paragraph(f"<b>{esc(j['title'])}</b> · {esc(j['company'])}", st["body"]))
        out.append(Paragraph(esc(j["dates"]), st["small"]))
        out += [Paragraph(esc(b), st["bullet"], bulletText="•") for b in j["bullets"]]
        out.append(Spacer(1, 4))
    return out


def pdf_classic(c, path):
    st = styles("Arial", colors.HexColor("#1f3b5c"))
    st["name"].alignment = st["sub"].alignment = st["small"].alignment = 1
    story = [Paragraph(esc(c["name"]), st["name"]), Paragraph(esc(c["headline"]), st["sub"]),
             Paragraph(esc(f"{c['location']}  |  {c['phone']}  |  {c['email']}"), st["small"]),
             Paragraph(esc(c["last_updated"]), st["small"]), Spacer(1, 4),
             HRFlowable(width="100%", color=colors.HexColor("#1f3b5c"))]
    st["small"] = ParagraphStyle("sm2", parent=st["small"], alignment=0)
    story += [Paragraph(c["h"]["summary"], st["head"]), Paragraph(esc(c["summary"]), st["body"]),
              Paragraph(esc(c["h"]["skills"]), st["head"]), Paragraph(esc(", ".join(c["skills"])), st["body"]),
              Paragraph(esc(c["h"]["experience"]), st["head"])] + pdf_jobs(c, st)
    story += [Paragraph(esc(c["h"]["education"]), st["head"])] + [Paragraph(esc(e), st["body"]) for e in c["education"]]
    story += [Paragraph(esc(c["h"]["languages"]), st["head"]), Paragraph(esc(" · ".join(c["languages"])), st["body"])]
    if c["certs"]:
        story += [Paragraph(esc(c["h"]["certifications"]), st["head"])] + [Paragraph(esc(x), st["body"]) for x in c["certs"]]
    SimpleDocTemplate(str(path), pagesize=A4, title=f"{c['name']} CV", author=c["name"],
                      leftMargin=18 * mm, rightMargin=18 * mm, topMargin=16 * mm, bottomMargin=16 * mm).build(story)


def pdf_sidebar(c, path):
    st = styles("Calibri", colors.HexColor("#2d6a4f"))
    W, H = A4
    left_w, m = 62 * mm, 12 * mm

    def bg(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(colors.HexColor("#e9f2ec"))
        canvas.rect(0, 0, left_w + m, H, stroke=0, fill=1)
        canvas.restoreState()

    frames = [Frame(m, m, left_w - 4 * mm, H - 2 * m, id="left"),
              Frame(left_w + m + 6 * mm, m, W - left_w - 2 * m - 6 * mm, H - 2 * m, id="right")]
    doc = BaseDocTemplate(str(path), pagesize=A4, title=f"{c['name']} - Resume", author=c["name"],
                          pageTemplates=[PageTemplate(frames=frames, onPage=bg)])
    left = [Paragraph("Contact", st["head"]), Paragraph(esc(c["phone"]), st["small"]),
            Paragraph(esc(c["email"]), st["small"]), Paragraph(esc(c["location"]), st["small"]),
            Spacer(1, 4), Paragraph(esc(c["last_updated"]), st["small"]),
            Paragraph(esc(c["h"]["skills"]), st["head"])]
    left += [Paragraph(esc(s), st["bullet"], bulletText="•") for s in c["skills"]]
    left += [Paragraph(esc(c["h"]["languages"]), st["head"])] + [Paragraph(esc(l), st["small"]) for l in c["languages"]]
    if c["certs"]:
        left += [Paragraph(esc(c["h"]["certifications"]), st["head"])] + [Paragraph(esc(x), st["small"]) for x in c["certs"]]
    right = [Paragraph(esc(c["name"]), st["name"]), Paragraph(esc(c["headline"]), st["sub"]),
             Paragraph(esc(c["h"]["summary"]), st["head"]), Paragraph(esc(c["summary"]), st["body"]),
             Paragraph(esc(c["h"]["experience"]), st["head"])] + pdf_jobs(c, st)
    right += [Paragraph(esc(c["h"]["education"]), st["head"])] + [Paragraph(esc(e), st["body"]) for e in c["education"]]
    doc.build(left + [FrameBreak()] + right)


def pdf_banner(c, path):
    accent = colors.HexColor("#5a2a82")
    st = styles("Calibri", accent)
    W, H = A4

    def band(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(accent)
        canvas.rect(0, H - 36 * mm, W, 36 * mm, stroke=0, fill=1)
        canvas.setFillColor(colors.white)
        canvas.setFont("Calibri-Bold", 24)
        canvas.drawString(18 * mm, H - 19 * mm, c["name"])
        canvas.setFont("Calibri", 12)
        canvas.drawString(18 * mm, H - 27 * mm, c["headline"])
        canvas.restoreState()

    story = [Paragraph(esc(f"Tel: {c['phone']}    Email: {c['email']}    {c['location']}"), st["small"]),
             Paragraph(esc(c["h"]["experience"]).upper(), st["head"])] + pdf_jobs(c, st)
    story += [Paragraph(esc(c["h"]["skills"]).upper(), st["head"])]
    story += [Paragraph(esc(s), st["bullet"], bulletText="▪") for s in c["skills"]]
    story += [Paragraph(esc(c["h"]["languages"]).upper(), st["head"]),
              Paragraph(esc(", ".join(c["languages"])), st["body"]),
              Paragraph(esc(c["h"]["education"]).upper(), st["head"])] + [Paragraph(esc(e), st["body"]) for e in c["education"]]
    if c["certs"]:
        story += [Paragraph(esc(c["h"]["certifications"]).upper(), st["head"])] + [Paragraph(esc(x), st["body"]) for x in c["certs"]]
    story += [Spacer(1, 10), Paragraph(esc(c["last_updated"]), st["small"])]
    SimpleDocTemplate(str(path), pagesize=A4, title="Curriculum Vitae", author=c["name"], topMargin=42 * mm,
                      leftMargin=18 * mm, rightMargin=18 * mm).build(story, onFirstPage=band)


def pdf_compact(c, path):
    st = styles("Georgia", colors.HexColor("#7a1f1f"))
    head = Table([[[Paragraph(esc(c["name"]), st["name"]), Paragraph(esc(c["headline"]), st["sub"])],
                   [Paragraph(esc(c["phone"]), st["small"]), Paragraph(esc(c["email"]), st["small"]),
                    Paragraph(esc(c["location"]), st["small"]), Paragraph(esc(c["last_updated"]), st["small"])]]],
                 colWidths=[105 * mm, 70 * mm])
    head.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                              ("LINEBELOW", (0, 0), (-1, 0), 1, colors.HexColor("#7a1f1f"))]))
    rows = []
    for j in c["jobs"]:
        cell = [Paragraph(f"<b>{esc(j['title'])}</b>", st["body"]), Paragraph(esc(j["company"]), st["small"])]
        cell += [Paragraph(esc(b), st["bullet"], bulletText="–") for b in j["bullets"]]
        rows.append([Paragraph(esc(j["dates"]), st["small"]), cell])
    jobs = Table(rows, colWidths=[38 * mm, 137 * mm])
    jobs.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
    story = [head, Paragraph(esc(c["h"]["summary"]), st["head"]), Paragraph(esc(c["summary"]), st["body"]),
             Paragraph(esc(c["h"]["experience"]), st["head"]), jobs,
             Paragraph(esc(c["h"]["skills"]), st["head"]), Paragraph(esc("  ·  ".join(c["skills"])), st["body"]),
             Paragraph(esc(c["h"]["languages"]), st["head"]), Paragraph(esc("; ".join(c["languages"])), st["body"]),
             Paragraph(esc(c["h"]["education"]), st["head"])] + [Paragraph(esc(e), st["body"]) for e in c["education"]]
    if c["certs"]:
        story += [Paragraph(esc(c["h"]["certifications"]), st["head"])] + [Paragraph(esc(x), st["body"]) for x in c["certs"]]
    SimpleDocTemplate(str(path), pagesize=A4, title=c["name"], author=c["name"],
                      leftMargin=17 * mm, rightMargin=17 * mm, topMargin=15 * mm).build(story)


# ---------------------------------------------------------------- Word layouts

def docx_base(font):
    d = Document()
    d.styles["Normal"].font.name = font
    d.styles["Normal"].font.size = Pt(10.5)
    d.core_properties.author = ""
    return d


def docx_jobs_list(d, c):
    for j in c["jobs"]:
        para = d.add_paragraph()
        para.add_run(j["title"]).bold = True
        para.add_run(f", {j['company']}")
        d.add_paragraph(j["dates"]).runs[0].italic = True
        for b in j["bullets"]:
            d.add_paragraph(b, style="List Bullet")


def docx_plain(c, path):
    d = docx_base("Calibri")
    d.add_heading(c["name"], 0)
    d.add_paragraph(c["headline"])
    d.add_paragraph(f"{c['location']} • {c['phone']} • {c['email']}")
    d.add_paragraph(c["last_updated"])
    d.add_heading(c["h"]["summary"], 1)
    d.add_paragraph(c["summary"])
    d.add_heading(c["h"]["skills"], 1)
    for s in c["skills"]:
        d.add_paragraph(s, style="List Bullet")
    d.add_heading(c["h"]["experience"], 1)
    docx_jobs_list(d, c)
    d.add_heading(c["h"]["education"], 1)
    for e in c["education"]:
        d.add_paragraph(e)
    if c["certs"]:
        d.add_heading(c["h"]["certifications"], 1)
        for x in c["certs"]:
            d.add_paragraph(x)
    d.add_heading(c["h"]["languages"], 1)
    for l in c["languages"]:
        d.add_paragraph(l, style="List Bullet")
    d.save(path)


def docx_table(c, path):
    d = docx_base("Arial")
    top = d.add_table(rows=1, cols=2)
    left, right = top.rows[0].cells
    r = left.paragraphs[0].add_run(c["name"])
    r.bold, r.font.size = True, Pt(18)
    left.add_paragraph(c["headline"])
    right.paragraphs[0].add_run(f"Phone: {c['phone']}")
    right.add_paragraph(f"Email: {c['email']}")
    right.add_paragraph(c["location"])
    right.add_paragraph(c["last_updated"])
    d.add_heading(c["h"]["summary"], 2)
    d.add_paragraph(c["summary"])
    d.add_heading(c["h"]["experience"], 2)
    t = d.add_table(rows=1, cols=3)
    t.style = "Light Grid Accent 1"
    for cell, txt in zip(t.rows[0].cells, ["Dates", "Role", "Highlights"]):
        cell.text = txt
    for j in c["jobs"]:
        row = t.add_row().cells
        row[0].text = j["dates"]
        row[1].text = f"{j['title']}\n{j['company']}"
        row[2].text = "\n".join(f"• {b}" for b in j["bullets"])
    d.add_heading(c["h"]["skills"], 2)
    st = d.add_table(rows=0, cols=2)
    for i in range(0, len(c["skills"]), 2):
        row = st.add_row().cells
        for k, s in enumerate(c["skills"][i:i + 2]):
            row[k].text = s
    d.add_heading(c["h"]["languages"], 2)
    d.add_paragraph(", ".join(c["languages"]))
    d.add_heading(c["h"]["education"], 2)
    for e in c["education"]:
        d.add_paragraph(e)
    if c["certs"]:
        d.add_heading(c["h"]["certifications"], 2)
        for x in c["certs"]:
            d.add_paragraph(x)
    d.save(path)


def docx_modern(c, path):
    d = docx_base("Georgia")
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(c["last_updated"])
    run.italic, run.font.size, run.font.color.rgb = True, Pt(9), RGBColor(0x77, 0x77, 0x77)
    p = d.add_paragraph()
    run = p.add_run(c["name"].upper())
    run.bold, run.font.size, run.font.color.rgb = True, Pt(22), RGBColor(0x0B, 0x6E, 0x79)
    d.add_paragraph(c["headline"]).runs[0].italic = True
    d.add_paragraph(f"{c['email']}  /  {c['phone']}  /  {c['location']}")
    d.add_heading(c["h"]["languages"], 2)
    d.add_paragraph("  |  ".join(c["languages"]))
    d.add_heading(c["h"]["skills"], 2)
    d.add_paragraph(" · ".join(c["skills"]))
    d.add_heading(c["h"]["experience"], 2)
    for j in c["jobs"]:
        para = d.add_paragraph()
        para.add_run(f"{j['dates']}   ").bold = True
        para.add_run(f"{j['title']} @ {j['company']}")
        for b in j["bullets"]:
            d.add_paragraph(b, style="List Bullet")
    d.add_heading(c["h"]["education"], 2)
    for e in c["education"]:
        d.add_paragraph(e)
    if c["certs"]:
        d.add_heading(c["h"]["certifications"], 2)
        for x in c["certs"]:
            d.add_paragraph(x)
    d.add_heading(c["h"]["summary"], 2)
    d.add_paragraph(c["summary"])
    d.save(path)


# ---------------------------------------------------------------- scanned (image-only) PDF

def scanned_pdf(c, path):
    W, H = 1654, 2339  # A4 at 200 dpi
    img = Image.new("L", (W, H), 246)
    draw = ImageDraw.Draw(img)
    f_big = ImageFont.truetype(str(FONT_DIR / "arialbd.ttf"), 54)
    f_head = ImageFont.truetype(str(FONT_DIR / "arialbd.ttf"), 34)
    f_body = ImageFont.truetype(str(FONT_DIR / "arial.ttf"), 27)
    y = 130

    def line(text, font, indent=0, gap=10):
        nonlocal y
        draw.text((140 + indent, y), text, fill=35, font=font)
        y += font.size + gap

    def wrapped(text, indent=0):
        for i, part in enumerate(textwrap.wrap(text, 92)):
            line(part, f_body, indent + (24 if i else 0), 6)

    line(c["name"], f_big, gap=16)
    line(c["headline"], f_body)
    line(f"{c['location']}  |  {c['phone']}  |  {c['email']}", f_body)
    line(c["last_updated"], f_body, gap=28)
    line(c["h"]["summary"].upper(), f_head)
    wrapped(c["summary"])
    y += 18
    line(c["h"]["skills"].upper(), f_head)
    wrapped(", ".join(c["skills"]))
    y += 18
    line(c["h"]["experience"].upper(), f_head)
    for j in c["jobs"]:
        line(f"{j['title']} - {j['company']}   ({j['dates']})", f_body, gap=6)
        for b in j["bullets"]:
            wrapped("• " + b, 20)
        y += 10
    line(c["h"]["education"].upper(), f_head)
    for e in c["education"]:
        wrapped(e)
    y += 18
    line(c["h"]["languages"].upper(), f_head)
    wrapped(", ".join(c["languages"]))

    # make it look photocopied: speckles, slight tilt, blur
    rng = random.Random(4711)
    for _ in range(9000):
        draw.point((rng.randrange(W), rng.randrange(H)), fill=rng.randint(120, 200))
    img = img.rotate(0.8, resample=Image.BICUBIC, fillcolor=246).filter(ImageFilter.GaussianBlur(0.7))
    img.save(path, "PDF", resolution=200.0)


RENDER = {"classic": pdf_classic, "sidebar": pdf_sidebar, "banner": pdf_banner, "compact": pdf_compact,
          "plain": docx_plain, "table": docx_table, "modern": docx_modern, "scanned": scanned_pdf}

# ---------------------------------------------------------------- questions


def matches(f, c):
    return (c.get("role") in (None, f["role"])
            and set(c.get("skills_all", [])) <= set(f["skills"])
            and set(c.get("languages_all", [])) <= set(f["languages"])
            and set(c.get("certs_all", [])) <= set(f["certifications"])
            and set(c.get("flags_all", [])) <= set(f["flags"])
            and f["years_experience"] >= c.get("min_years", 0)
            and f["ghl_years"] >= c.get("min_ghl_years", 0))


def answer(files, criteria, include_scanned=False):
    """Names of people whose newest readable CV matches. Duplicates collapse to one person."""
    latest = {}
    for f in files:
        if f["messy_case"] == "scanned_no_text_layer" and not include_scanned:
            continue
        if f["person_id"] not in latest or f["last_updated"] > latest[f["person_id"]]["last_updated"]:
            latest[f["person_id"]] = f
    return sorted(f["full_name"] for f in latest.values() if matches(f, criteria))


QUESTIONS = [
    ("normal", "Which QuickBooks bookkeepers are Certified QuickBooks ProAdvisors with at least 5 years of bookkeeping experience?",
     {"role": "bookkeeper", "skills_all": ["QuickBooks Online"], "certs_all": ["QuickBooks ProAdvisor"], "min_years": 5},
     "Four bookkeepers hold the ProAdvisor certification, but Grace Achieng has only 3 years, so she is excluded."),
    ("skill_plus_language", "Find GoHighLevel specialists with 3+ years of GHL experience who speak Arabic.",
     {"skills_all": ["GoHighLevel"], "min_ghl_years": 3, "languages_all": ["Arabic"]},
     "GHL is written as GHL / GoHighLevel / Go High Level / HighLevel across CVs. Several people have 3+ years "
     "of GHL but no Arabic, and several Arabic speakers have under 3 years of GHL; only these have both."),
    ("normal", "Which customer support agents have worked with Gorgias on a Shopify store?",
     {"role": "support", "skills_all": ["Gorgias", "Shopify"]},
     "Only three support agents list Gorgias; all three also list Shopify. Zendesk-only agents do not count."),
    ("impossible", "Find a bookkeeper who is a licensed CPA with 8+ years of NetSuite experience.",
     {"role": "bookkeeper", "certs_all": ["CPA"], "skills_all": ["NetSuite"], "min_years": 8},
     "No CV mentions a CPA licence or NetSuite. The right answer is to say nobody fits."),
    ("needs_newer_cv", "Who has 3+ years of GoHighLevel and also builds automations in Make.com (Integromat)?",
     {"skills_all": ["GoHighLevel", "Make.com"], "min_ghl_years": 3},
     "Bilal Qureshi only qualifies through his newer CV (2026-08-15: 3 years GHL, adds Make). His older CV "
     "(2025-02-10) shows 2 years and no Make. Lina Haddad and Joanna Dimaculangan qualify from single CVs."),
    ("normal", "Which social media managers edit video in CapCut and have run TikTok Ads campaigns?",
     {"role": "social", "skills_all": ["CapCut", "TikTok Ads"]},
     "Many social media managers list CapCut or organic TikTok, but only three have run TikTok Ads."),
    ("duplicate_person", "Which customer support agents have used Kustomer?",
     {"skills_all": ["Kustomer"]},
     "Omar Haddadin has two copies of the same CV (one PDF, one Word, same Last updated date). He must be "
     "listed once, not twice."),
    ("scanned_only", "Who has migrated clients from Keap (Infusionsoft) to GoHighLevel?",
     {"skills_all": ["Keap"]},
     "The only person with this experience is Rhea Macaraeg, and it appears only in her scanned CV "
     "(cvs/scan_0047.pdf), which has no text layer. A text-based search cannot see it, so the expected answer "
     "excludes her and nobody else matches."),
    ("normal", "Which executive virtual assistants have worked directly for a CEO or founder and handle travel booking?",
     {"role": "eva", "flags_all": ["supported_ceo"], "skills_all": ["Travel booking"]},
     "Four EAs worked for a CEO or founder (job title says so), but Danica Petrović does not list travel booking."),
    ("impossible", "Find a GoHighLevel specialist who is fluent in German.",
     {"skills_all": ["GoHighLevel"], "languages_all": ["German"]},
     "Nobody in the pool lists German."),
]

# ---------------------------------------------------------------- main


def main():
    CV_DIR.mkdir(exist_ok=True)
    for old in CV_DIR.iterdir():
        old.unlink()

    ghl_people = [s[0] for s in PEOPLE if s[2] == "ghl" or s[6].get("ghl_years")]
    people = [Person(s, GHL_SPELLINGS[ghl_people.index(s[0]) % 4] if s[0] in ghl_people else "GHL") for s in PEOPLE]
    assert len({p.phone for p in people}) == 55, "phone collision"

    files, prev_path = [], None
    for p in people:
        versions = SPECIAL_FILES.get(p.num, [("none", p.lu_default, p.fmt_default, p.layout_default, {})])
        names = SPECIAL_NAMES.get(p.num) or [default_filename(p.num, p.name, p.fmt_default, p.lu_default)]
        for (case, lu, fmt, layout, ov), fname in zip(versions, names):
            v = p.version(lu, ov)
            path = CV_DIR / fname
            assert not path.exists(), fname
            if ov.get("copy_of_previous"):
                shutil.copyfile(prev_path, path)
            else:
                RENDER[layout](cv_content(v), path)
            prev_path = path
            files.append(dict(
                file=f"cvs/{fname}", person_id=p.id, full_name=p.name, role=p.role,
                target_role=ROLE_TITLE[p.role], format=fmt, layout=layout, last_updated=lu.isoformat(),
                email=v["email"], phone=p.phone, location=f"{p.city}, {p.country}",
                languages=[n for n, _ in p.languages], language_levels=dict(p.languages),
                skills=v["skills"], skills_as_written=[skill_text(p, v, s) for s in v["skills"]],
                ghl_years=v["ghl_years"], years_experience=v["years"], certifications=v["certs"],
                flags=v["flags"],
                jobs=[dict(title=j["title"], company=j["company"], start="%04d-%02d" % j["start"],
                           end="%04d-%02d" % j["end"] if j["end"] else "present") for j in v["jobs"]],
                messy_case=case, related_files=[]))
    # link files of the same person
    for f in files:
        f["related_files"] = [g["file"] for g in files if g["person_id"] == f["person_id"] and g is not f]
    notes = {
        "scanned_no_text_layer": "PDF is a single page image (photocopy look). No extractable text.",
        "duplicate_diff_email_same_phone": "Same person, same Last updated date, same phone; the two files use different emails.",
        "duplicate_pdf_and_docx": "Same person and content saved once as PDF and once as Word.",
        "duplicate_identical_copy": "Byte-identical copy of the same Word file (\" (1)\" in the name).",
        "updated_older": "Older CV of a person who also has a newer one. Use the newer file.",
        "updated_newer": "Newer CV of the same person: more years and extra skills compared to the older file.",
        "none": "Regular CV.",
    }
    for f in files:
        f["messy_note"] = notes[f["messy_case"]]

    for f in files:
        assert not (RESERVED - {"Make.com", "Keap", "Gorgias", "Kustomer", "TikTok Ads"}) & set(f["skills"])

    questions, qtypes = [], []
    for i, (qtype, text, crit, why) in enumerate(QUESTIONS, 1):
        names = answer(files, crit)
        assert len(names) <= 5, (text, names)
        if qtype in ("impossible", "scanned_only"):
            assert not names, (text, names)
        else:
            assert names, text
        if qtype == "scanned_only":
            assert answer(files, crit, include_scanned=True) == ["Rhea Macaraeg"]
        questions.append({"question": text, "expected": names or "NO_MATCH", "why": why})
        qtypes.append({"index": i, "type": qtype, "criteria": crit})

    key = {
        "description": "Answer key for the fictional recruiting demo. All people, companies and schools are fictional.",
        "reference_date": "2026-11-04",
        "file_count": len(files),
        "person_count": len({f["person_id"] for f in files}),
        "transcripts": [
            {"file": "transcripts/Brightwell Dental Studios - GHL Automation Specialist.txt", "role": "GHL Automation Specialist",
             "client": "Brightwell Dental Studios", "changed_mind": "part-time 20h -> full-time; Arabic nice-to-have -> required",
             "missing": ["pay rate"], "final_needs": "3+ yrs GoHighLevel, Arabic required, full-time, ASAP"},
            {"file": "transcripts/Hollis Crane Advisory - Executive Virtual Assistant.txt", "role": "Executive Virtual Assistant",
             "client": "Hollis Crane Advisory", "changed_mind": "US Eastern hours -> US Pacific hours",
             "missing": ["pay rate", "exact start date"],
             "final_needs": "has supported a CEO/founder, travel booking, Google Workspace, Notion, Pacific hours, full-time"},
            {"file": "transcripts/Fernhill Coffee Roasters - QuickBooks Bookkeeper.txt", "role": "QuickBooks Bookkeeper",
             "client": "Fernhill Coffee Roasters", "changed_mind": "QuickBooks Desktop -> QuickBooks Online",
             "missing": ["pay/budget", "weekly hours"],
             "final_needs": "QuickBooks Online, 5+ years, catch-up work, Shopify payouts; Gusto and ProAdvisor are pluses"},
            {"file": "transcripts/Luma Yoga Collective - Social Media Manager.txt", "role": "Social Media Manager",
             "client": "Luma Yoga Collective", "changed_mind": "Instagram + TikTok -> Instagram + LinkedIn",
             "missing": ["pay rate", "posting volume"],
             "final_needs": "Instagram + LinkedIn, Reels in CapCut, Canva, 2+ years, part-time 25-30h"},
            {"file": "transcripts/Kettle & Kite Home Goods - Customer Support Agent.txt", "role": "Customer Support Agent",
             "client": "Kettle & Kite Home Goods", "changed_mind": "Zendesk -> Gorgias",
             "missing": ["pay rate", "shift"],
             "final_needs": "two agents, Gorgias + Shopify, email and chat, 1+ year e-commerce support, full-time"},
        ],
        "files": files,
        "test_question_types": qtypes,
    }
    (ROOT / "data_key.json").write_text(json.dumps(key, indent=2, ensure_ascii=False), encoding="utf-8")
    (ROOT / "test_questions.json").write_text(json.dumps(questions, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Wrote {len(files)} CV files for {key['person_count']} people, data_key.json, test_questions.json")


if __name__ == "__main__":
    main()
