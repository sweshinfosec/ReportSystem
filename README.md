# ReportSystem

A Django app that turns Nessus (or similar) vulnerability scan exports into clean, presentable
Vulnerability Assessment (VA) deliverables: a formatted interim Excel/CSV report, a Round 1 vs
Round 2 (pre- vs post-remediation) comparison, and a final polished Word report.

## Requirements

- Python 3.10+ (developed against 3.13)
- pip

## Installation

```bash
# from the project root (same folder as manage.py)
python3 -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate

pip install -r requirements.txt

python3 manage.py migrate
```

## Running the server

```bash
python3 manage.py runserver
```

Then open **http://127.0.0.1:8000/** in a browser.

## Features / Pages

### Home (`/`)
Landing page for the app — links out to the other tools below.

### Reporting Dashboard (`/reporting-dashboard/`)
A dashboard placeholder view for reporting-related links/status.

### Phase 1 — Interim Report (`/interim-pre/`)
Upload one or more raw scan CSV exports (e.g. Nessus). The app:
- Concatenates and de-duplicates the uploaded scans
- Normalizes the `Risk` column to `Severity` and keeps only valid severities
  (Critical / High / Medium / Low)
- Sorts findings by Severity → Name → CVSS score → Host (numeric IP order)
- Produces a formatted **Interim Report** as both `.xlsx` (with a colour-coded
  severity column and a `Summary` tab showing vulnerability counts by host and by
  finding) and `.csv`

Use this to get a clean baseline report from a single round of scanning, or to
prepare the "before" data for Phase 2 below.

### Phase 2 — Post-Remediation Compare (`/interim-post/`)
Upload your Round 1 (pre-remediation) report plus one or more Round 2
(post-remediation) scan files. The app:
- Runs the Round 2 scans through the same cleanup as Phase 1
- Matches findings between rounds by finding name + host + port
- Marks each Round 1 finding **Open** if it's still present in Round 2, or
  **Fixed & Closed** if it's gone
- Flags anything newly reported in Round 2 (not in Round 1) as a new **Open** finding
- Produces a **Consolidated VA Report** (`.xlsx`/`.csv`) with separate `R1`, `R2`,
  and combined `InterimReport_VA` tabs, sorted by Severity, plus a `Summary` tab

This is the file to hand to Phase 3 when you want a report that shows what's been
fixed and what's still outstanding.

### Phase 3 — Word Report (`/word-report/`)
Upload either the Phase 1 interim report or the Phase 2 consolidated report
(`.xlsx` or `.csv`). The app generates a fully formatted **Word (.docx) VA report**:
cover page, document control/review tables, an executive summary with severity
totals by host, a Vulnerabilities Summary table, and one detailed section per
finding (affected servers, synopsis, description, recommendation, sample
findings, and sign-off fields for customer/Opensource comments).

Where a project or customer name hasn't been supplied yet, the report shows a
`<Project>` / `<customer>` placeholder highlighted in yellow so it's easy to spot
and fill in before the report goes out.

## Project layout

```
ReportSystem/
├── core/                   # Django project settings/urls
├── processor/
│   ├── views.py             # the three phases above
│   ├── utils_interim.py     # Phase 1 report generation
│   ├── utils_compare.py     # Phase 2 Round 1 vs Round 2 comparison
│   ├── utils_word.py        # Phase 3 Word document generation
│   ├── templates/processor/ # HTML templates for each page
│   └── static/               # app CSS
├── static/images/           # logo / certification images used in the Word report
├── media/                    # uploaded scans and generated reports (git-ignored)
└── manage.py
```

## Notes

- Uploaded scans and generated reports are written to `media/`, which is
  git-ignored — nothing there is version controlled.
- `db.sqlite3` is also git-ignored; run `python3 manage.py migrate` to create it
  locally.
