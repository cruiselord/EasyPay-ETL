# NIBSS EasyPay Reconciliation ETL

Turn raw **NIBSS EasyPay / Direct Debit** CSV exports into a single, polished
Excel workbook — with a dashboard summary, a filtered exceptions sheet, and
automatic email delivery. No FTP, no database, no cloud service required.

> **Output example:** `output/NIBSS_Reconciliation_29_07_2026.xlsx`

---

## Contents

- [What it does](#what-it-does)
- [How it works](#how-it-works)
- [Requirements](#requirements)
- [Quick start](#quick-start)
- [Configuration (.env)](#configuration-env)
- [Getting your data](#getting-your-data)
- [Running it](#running-it)
- [Automating it daily (macOS)](#automating-it-daily-macos)
- [The output workbook](#the-output-workbook)
- [Project structure](#project-structure)
- [Troubleshooting](#troubleshooting)

---

## What it does

NIBSS (Nigeria Inter-Bank Settlement System) publishes EasyPay and Direct Debit
reports four times a day as **pipe-delimited (`|`) CSV files**. This tool:

1. **Reads** every CSV from a day's session folders
2. **Cleans & standardises** the columns (fixes NIBSS typos, maps bank codes)
3. **Deduplicates** — a successful payment appears in both reports, so the
   Direct Debit copy is removed to avoid double-counting
4. **Builds** a styled 5-sheet Excel workbook (see [The output workbook](#the-output-workbook))
5. **Emails** the workbook to your team and an executive summary to your MD

Optionally, it can also **download the files itself** from the NIBSS Thin Client
web API (no manual download, no FTP) and **run on a daily schedule**.

---

## How it works

```
NIBSS Thin Client            fetcher.py               reader.py
  (web API)   ──login+list──▶  downloads CSVs  ──▶  parses | CSVs,
   │                            into DD_MM_YYYY/        standardises columns,
   │                            session folders         dedupes Direct Debit
   │                                                          │
   │                                                          ▼
   │                                                    writer.py
   │ (optional) ──────────────────────────────────▶  builds 5-sheet .xlsx
   │                                                        │
   │                                                        ▼
   │                                                     mailer.py
   └────────────────────────────────────────────▶  emails team + MD
```

All processing is **local**. The only external calls are the initial file
download (from NIBSS) and the outgoing email (via SMTP).

---

## Requirements

| Need | Version / detail |
|---|---|
| Python | 3.10+ (tested on 3.14) |
| pip packages | `pandas`, `openpyxl` (see `requirements.txt`) |
| Email (optional) | A Microsoft 365 / Outlook mailbox with SMTP auth enabled |
| Schedule (optional) | macOS (`launchd`) |

---

## Quick start

```bash
# 1. Get the code
git clone https://github.com/<your-org>/EasyPay-ETL.git
cd EasyPay-ETL

# 2. Create a virtual environment and install dependencies
python3 -m venv nibss_etl/venv
source nibss_etl/venv/bin/activate            # Windows: nibss_etl\venv\Scripts\activate
pip install -r requirements.txt

# 3. Configure your credentials
cp .env.example .env
#    ...then open .env and fill in your values (see Configuration below)

# 4. Fetch today's files and build the report
./run.sh
```

The workbook lands in `nibss_etl/output/`.

---

## Configuration (.env)

Everything is configured through environment variables in `.env` (gitignored —
never commit it). Copy the template and fill it in:

```bash
cp .env.example .env
```

| Variable | Required | Purpose |
|---|---|---|
| `NIBSS_USER` | yes (for fetch) | Your NIBSS Thin Client username |
| `NIBSS_PASSWORD` | yes (for fetch) | Your NIBSS Thin Client password |
| `EMAIL_USER` | for email | Sender mailbox (your Outlook address) |
| `EMAIL_PASSWORD` | for email | Microsoft **app password** (not your normal password) |
| `TEAM_EMAIL` | for email | Workbook recipients, comma-separated |
| `MD_EMAIL` | optional | Executive-summary recipient (leave empty to skip the MD email) |
| `NIBSS_RUN_AFTER` | optional | Earliest time the scheduler may run (default `07:00`) |

> **Microsoft app password:** create one at <https://mysignins.microsoft.com/security-info>
> (requires MFA). Your tenant admin must enable SMTP AUTH (Exchange admin →
> Mail flow → SMTP AUTH).

---

## Getting your data

You don't need FTP. The NIBSS Thin Client portal is a front-end over a JSON web
API, and `fetcher.py` talks to it directly.

```bash
# Fetch today's files (or pass a specific date)
python nibss_etl/fetcher.py                    # today
python nibss_etl/fetcher.py 18_08_2026         # a specific day
python nibss_etl/fetcher.py --list             # browse the remote file tree
```

Files land in `<DD_MM_YYYY>/…` at the project root, in the exact session-based
structure the pipeline expects:

```
<DD_MM_YYYY>/
  <YYYYMMDDHHMMSS>/          ← 4 sessions per day
    Easy Pay/
      credit_failed.csv
      credit_successful.csv
      debit_failed.csv
      debit_successful.csv
      name_enquiry_failed.csv
      name_enquiry_successful.csv
    Direct Debit/
      credit_failed.csv
      credit_successful.csv
      debit_failed.csv
      debit_successful.csv
```

Re-running skips files already downloaded (matched by size).

**Prefer to download manually?** Log in at the NIBSS Thin Client, export the
EasyPay / Direct Debit reports for the day, and place the resulting
`<DD_MM_YYYY>/` folder at the project root.

> The NIBSS Thin Client portal is at:
> `https://nibsswebserver.nibss-plc.com.ng/ThinClient/WTM/public/#/login`
> (this URL is baked into `fetcher.py` — change it there if NIBSS moves it).

---

## Running it

```bash
./run.sh                # fetch + process the current day
./run.sh 18_08_2026     # fetch + process a specific day
```

`run.sh` does two things: runs `fetcher.py` (download) then `main.py`
(workbook + email). To skip the download step, run `main.py` directly:

```bash
cd nibss_etl && source venv/bin/activate
python main.py            # most recent date folder
python main.py 18_08_2026 # specific date
```

---

## Automating it daily (macOS)

A small guard script plus a `launchd` agent runs the whole thing without a
human, every morning.

```bash
./install_launchd.sh                  # schedule for 07:00 (default)
NIBSS_RUN_AFTER=06:00 ./install_launchd.sh   # custom time
```

This installs a LaunchAgent (`local.nibss-etl`) that fires `auto_run.sh`:

- **daily at your chosen time** (`StartCalendarInterval`)
- **on login/wake** (`RunAtLoad`)
- **every 15 min** as a catch-up (idempotent — a marker file in `logs/` stops
  duplicate runs)

`auto_run.sh` reconciles **yesterday's** data (the last NIBSS session closes at
23:59, so a full day is only complete after midnight). It only runs when:

1. the clock is at/after the run time,
2. NIBSS is reachable (the Mac has internet), and
3. yesterday's report hasn't already been produced.

> **macOS privacy note:** launchd cannot read `~/Documents` by default. If the
> project lives under `~/Documents`, grant **Full Disk Access** to `/bin/bash`
> (System Settings → Privacy & Security → Full Disk Access). Alternatively move
> the project outside `~/Documents`.

To remove the schedule:

```bash
launchctl unload ~/Library/LaunchAgents/local.nibss-etl.plist
rm ~/Library/LaunchAgents/local.nibss-etl.plist
```

---

## The output workbook

| Sheet | Contents |
|---|---|
| **Summary** | Dashboard: KPI cards, session overview (successful vs failed), top banks by volume, failure analysis by NIP code, channel breakdown, transaction patterns. |
| **EasyPay** | Every EasyPay row in a filterable Excel table (failed rows highlighted red). |
| **DirectDebit** | Same structure as EasyPay, with duplicate successful payments removed. |
| **Exceptions** | Every failed row from both sources, sorted by amount descending, with a `source` column. |
| **Run_Log** | Pipeline metadata: per-file manifest (found/missing, size, row count), dropped columns, schema warnings. |

---

## Project structure

```
.
├── run.sh                    # fetch + process (one command)
├── auto_run.sh               # scheduler guard (time + internet + once-per-day)
├── install_launchd.sh        # install the macOS daily schedule
├── requirements.txt          # pip dependencies
├── .env.example              # copy to .env and fill in
└── nibss_etl/
    ├── fetcher.py            # downloads CSVs from the NIBSS Thin Client web API
    ├── main.py               # CLI entry point (date detection, orchestration)
    └── etl/
        ├── reader.py         # folder walking, CSV parsing, column standardisation
        ├── writer.py         # workbook builder (5 sheets, styling)
        └── mailer.py         # team + MD emails via SMTP
```

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `NIBSS_USER / NIBSS_PASSWORD not set` | `cp .env.example .env` and fill in your credentials. |
| `Emails skipped — missing env vars` | Ensure `EMAIL_USER`, `EMAIL_PASSWORD`, `TEAM_EMAIL` are set in `.env`. |
| `Login failed` from the fetcher | Check your NIBSS username/password; the account must be enabled for the Thin Client. |
| `Failed to save request context` during fetch | Transient NIBSS-side error — the fetcher retries automatically; just re-run. |
| Report has 0 rows | The day's files may be header-only (no transactions), or you ran before NIBSS published the day's sessions. |
| launchd job logs `Operation not permitted` | Grant Full Disk Access to `/bin/bash` (see [Automating it daily](#automating-it-daily-macos)). |
| `ModuleNotFoundError: pandas` | Activate the venv or `pip install -r requirements.txt`. |

---

## Notes & limitations

- **Pipe delimiter** — input CSVs use `|`, not commas. The reader handles this
  automatically; don't re-save files as comma-separated.
- **Name-enquiry files** are excluded from the totals (they are beneficiary
  lookups, not transactions).
- **Bank names** are mapped from NIBSS short codes via a dictionary in
  `writer.py`; unknown codes display as-is.
- **No database** — everything is file-based and runs locally.
