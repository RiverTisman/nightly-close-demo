# Nightly Close — restaurant close-out automation (demo)

A web app I built for the full-service restaurant I manage (≈150–200 staff) that turns the nightly tip-out from a manual copy-and-cross-check job into a few uploads and a review screen. It reads three systems that don't talk to each other — **Toast** (POS), **HotSchedules** (scheduling), and **ADP** (time clock) — cross-checks them against the manager's weekly Excel Tip Sheet, and hands back ready-to-paste entries plus a short list of things a human actually needs to decide.

**This repo is a public demo.** Every employee, EID, dollar amount and file in it is fictional ("Harbor & Vine," Thursday 9/24/26). The parsing and checking code is the same code that runs the real nightly close.

## Try it

Open the live demo, press **Load sample files**, then **Fire**. Two pages:

| Page | Inputs | What comes back |
| --- | --- | --- |
| **Daily Close** | Toast Shift Report CSV, HotSchedules daily roster | Server cash/credit tip entries and cash-sales pastes, bartender pool totals, support-staff roster by role, cut confirmation, and every row that needs a manager's call |
| **Punch Report Check** | ADP Punch Report, the week's Tip Sheet workbook, HotSchedules (optional) | Missing clock-outs with a suggested time from peers, early cuts, people paid in the wrong department, people on the sheet who never punched in, people who punched in but aren't on the sheet |

The sample files deliberately contain the real-world problems the app was built to catch: a name Toast spells differently, a call-in, a trainee, a terminated employee on the schedule, a server under the 4-hour tip-pool minimum, a large-party gratuity, a forgotten clock-out, a wrong department, and a Tip Sheet name that doesn't byte-match its dropdown.

## The design rule behind it

**Never guess a name match.** Excel's `SUMIF` returns `$0` — not an error — when a pasted name differs from the dropdown by a single space. That happened in production and silently zeroed one employee's tip-out for the night. So the app has no fuzzy matching anywhere: names resolve through one mapping file or get flagged, and the cut/role-swap inputs use a pick-from-the-list control so a typo is structurally impossible.

Other decisions in the same spirit:

- Anything ambiguous (clock-ins near the lunch/dinner cutover, call-ins, house shifts, trainees) defaults to "flag for review," never to silently included or excluded.
- Suggested clock-out times are averaged from peers in the same role *and* the same opener/mid/closer shift — and labeled as a suggestion, never written as a punch.
- Stateless: uploads live in a temp folder for one request and are deleted. No database.

## Stack

Python · FastAPI · Jinja2 · openpyxl · vanilla JS · Docker · Render. Built with Claude Code as my pair programmer; I owned the problem definition, the business rules, and verification against real nightly data.

## Run locally

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload       # http://localhost:8000
```

Regenerate the sample files with `python sample_data/generate.py`.

## Deploy (Render, free)

New → Web Service → connect this repo. Render detects the `Dockerfile`; pick the Free instance. Free services sleep after 15 minutes idle and take ~30–50 seconds to wake on the first visit.
