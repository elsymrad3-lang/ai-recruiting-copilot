# AI Recruiting Copilot

A copilot for a remote-staffing agency: it turns client call notes into job descriptions, finds the best 5 candidates with reasons, and runs email outreach, with a recruiter approving every step in Slack.

**Live portal:** https://subtle-shortbread-fdeb99.netlify.app

Built with **n8n**, **GoHighLevel**, **Slack**, **Supabase** (Postgres) and **Groq** (`openai/gpt-oss-120b`). Portal and scripts built with **Claude Code**.

## How it works

| Step | Where |
|---|---|
| Call transcript → job description → Slack **Approve / Reject** | n8n `5 - Job Description` |
| CVs (PDF/Word) → searchable profiles; newer CV replaces older; duplicates detected; unreadable CVs flagged | `ingest.py` → n8n `1 - CV Intake` |
| Plain-English question → top 5 with score, 2-line reason and source CV, or "No good match" | Slack bot → n8n `4 - Ask Copilot` → `2 - Find Candidates` |
| Candidate **Approve** → GHL contact + pipeline + 3-email sequence that stops on reply | n8n `3 - Candidate Approval` → GHL workflows *Candidate Intake*, *Candidate Outreach*, *Candidate Replied* |
| Reply → Slack alert + database update; GHL stages synced every 15 min | n8n `6 - Reply Handler`, `7 - Stage Sync` |
| 10 test questions with right answers → score out of 10 | n8n `8 - Run Test` |

## Rules the system enforces

- **Nothing reaches a candidate or client without Approve.** Only the Approve branch calls GHL; emails start from the tag it adds. Roles stay Draft until their job description is approved.
- **No invented candidates.** Candidates come only from the database and are filtered by code; an empty result returns "No good match" with the missing requirements.
- **Every recommendation names its CV**, and the AI's supporting quote must appear in that CV's text.
- **Unreadable CVs are flagged** (Slack alert + portal), never skipped.

## Scoring

A candidate must meet every must-have (skills with years, languages, certifications, years in the role, "worked for a CEO/founder"). Score = 70 + up to 20 for nice-to-haves + 10 if the CV was updated in the last 2 years. Top 5, never padded.

## Repository

| Path | What |
|---|---|
| `n8n/` | The 9 exported n8n workflows (import with *⋮ → Import from File*, then pick your own credentials) |
| `portal/index.html` | Recruiter portal: Open Roles, Pipeline, Test (read-only Supabase key) |
| `ingest.py` | Extracts CV text and sends each CV to the CV Intake webhook |
| `send_transcripts.py` | Sends the client transcripts to the Job Description webhook |
| `generate_data.py`, `check_data.py` | Generate the fake data and verify it |
| `cvs/`, `transcripts/` | 60 fake CVs (1 scanned, 3 duplicates, 2 updated) and 5 fake client calls |
| `test_questions.json`, `data_key.json` | The 10 test questions with answers, and the answer key for the data |

## Running the scripts

```
pip install requests python-dotenv pypdf python-docx
copy .env.example .env      # then fill in the two n8n webhook URLs
python ingest.py
python send_transcripts.py
```

All data is fictional. Outreach runs in test mode: emails go to the builder's own test inbox, never to the addresses in the CVs.
