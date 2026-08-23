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
- [Testing without NIBSS access](#testing-without-nibss-access)
- [The output workbook](#the-output-workbook)
- [Project structure](#project-structure)
- [Troubleshooting](#troubleshooting)
- [Notes & limitations](#notes--limitations)

---

## What it does

NIBSS (Nigeria Inter-Bank Settlement System) publishes EasyPay and Direct Debit
reports four times a day as **pipe-delimited (`|`) CSV files**. This tool:

1. **Downloads** the files itself from the NIBSS Thin Client web API — no FTP,
   no manual download — or reads them from a local folder
2. **Reads & cleans** every CSV: standardises column names, fixes NIBSS typos,
   maps bank short-codes to full names
3. **Deduplicates** — a successful payment appears in both reports, so the
   Direct Debit copy is removed to avoid double-counting
4. **Builds** a styled 5-sheet Excel workbook (see
   [The output workbook](#the-output-workbook))
5. **Emails** the workbook to your team and an executive summary to your MD
6. **Schedules** itself to run every morning (macOS `launchd`), with guards so
   it only runs when the Mac is awake and in use (lid open, display on), is
   online, and never duplicates a day's report

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
git clone https://github.com/cruiselord/EasyPay-ETL.git
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

### Emails

On a successful run the pipeline sends up to **two emails** via Outlook SMTP
(`smtp.office365.com:587`):

| Email | Recipient | Content |
|---|---|---|
| Reconciliation report | `TEAM_EMAIL` | The generated `.xlsx` workbook attached |
| Executive summary | `MD_EMAIL` (optional) | HTML summary: total txns, approved, failed, success rate, total/approved/failed value |

If `MD_EMAIL` is empty, only the team email is sent. If any of `EMAIL_USER`,
`EMAIL_PASSWORD`, `TEAM_EMAIL` is missing (or SMTP fails), emails are skipped
with a warning and the workbook is still written.

> **Microsoft app password:** create one at <https://mysignins.microsoft.com/security-info>
> (requires MFA). Your tenant admin must enable SMTP AUTH (Exchange admin →
> Mail flow → SMTP AUTH).

---

## Getting your data

You don't need FTP. The NIBSS Thin Client portal is a front-end over a JSON web
API, and `fetcher.py` talks to it directly.

```bash
python nibss_etl/fetcher.py                    # today
python nibss_etl/fetcher.py 18_08_2026         # a specific day
python nibss_etl/fetcher.py --list             # browse the remote file tree
python nibss_etl/fetcher.py --list --depth 3   # limit tree depth
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

The fetcher retries transient errors automatically and skips files already
downloaded (matched by size), so re-runs are safe and fast.

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
(workbook + email).

To skip the download step, run `main.py` directly:

```bash
cd nibss_etl && source venv/bin/activate
python main.py            # auto-detect the most recent DD_MM_YYYY folder
python main.py 18_08_2026 # process a specific date
```

| Argument | Description |
|---|---|
| `(none)` | Auto-detect the most recent `DD_MM_YYYY` folder at the project root |
| `DD_MM_YYYY` | Process a specific date folder, e.g. `python main.py 29_07_2026` |

---

## Automating it daily (macOS)

A small guard script plus a `launchd` agent runs the whole thing without a
human, every morning.

```bash
./install_launchd.sh                  # schedule for 07:00 (default)
NIBSS_RUN_AFTER=06:00 ./install_launchd.sh   # custom time
```

This installs a LaunchAgent (labelled `com.yourorg.nibss-etl` — `yourorg` is a
placeholder for your organisation's reverse-DNS short-name) that fires
`auto_run.sh`:

- **daily at your chosen time** (`StartCalendarInterval`)
- **on login/wake** (`RunAtLoad`)
- **every 15 min** as a catch-up (idempotent — a marker file in `logs/` stops
  duplicate runs)

`auto_run.sh` reconciles **yesterday's** data (the last NIBSS session closes at
23:59, so a full day is only complete after midnight). It only runs when **all**
of the following are true:

1. the clock is at/after the run time,
2. the Mac is in an **active session** — the lid is **open** and the **display
   is awake** (see [Why we require an active session](#why-we-require-an-active-session)),
3. NIBSS is reachable (the Mac has internet), and
4. yesterday's report hasn't already been produced.

Every trigger is written to `logs/auto_run.log` — including each skip and the
reason — so you can always see exactly why a run did or didn't happen.

### Why we require an active session

Originally the agent fired purely on the clock (`StartCalendarInterval` at
07:00). That's fine on a laptop while you're sitting at it — but with the lid
closed, macOS puts the machine into **clamshell sleep**, and `launchd` only gets
brief **DarkWake** slices (a few seconds of CPU every ~15 minutes). One morning
the job fired during a DarkWake:

- the Python process was frozen mid-download for ~40 minutes, waking only in
  2–9 second bursts;
- with Wi-Fi half-up during each sleep transition, the fetcher's requests to
  NIBSS failed with `HTTP Error 500` and it eventually gave up;
- the same run completed instantly when re-run with the lid open.

The schedule itself was never the problem — it was the job *trying to run while
the machine was asleep*. To fix it, `auto_run.sh` now checks for a genuine login
before doing any work:

- **lid open** — `ioreg` reports `AppleClamshellState = No`
- **display awake** — `powerd` holds the `Prevent sleep while display is on`
  assertion (`pmset -g assertions`)

If either check fails, the run is skipped and logged (`skipping — lid closed` /
`skipping — display asleep`). The 15-minute catch-up keeps retrying cheaply
until you actually log in, then the report runs once for the previous day.

This keeps the whole pipeline **local** — no cloud scheduler, no remote server,
no external dependency beyond the NIBSS fetch and the SMTP email. The trade-off
is deliberate: the report is produced a little later (whenever you first open
the Mac after 07:00) in exchange for never running unattended on a sleeping
machine.

> **macOS privacy note:** launchd cannot read `~/Documents` by default. If the
> project lives under `~/Documents`, grant **Full Disk Access** to `/bin/bash`
> (System Settings → Privacy & Security → Full Disk Access). Alternatively move
> the project outside `~/Documents`.

To remove the schedule:

```bash
launchctl unload ~/Library/LaunchAgents/com.yourorg.nibss-etl.plist
rm ~/Library/LaunchAgents/com.yourorg.nibss-etl.plist
```

---

## Testing without NIBSS access

You can verify the whole pipeline works with dummy data — no NIBSS account
needed:

```bash
# Create a test date folder with one session
mkdir -p test_01_01_2026/20260101000000/Easy\ Pay
mkdir -p test_01_01_2026/20260101000000/Direct\ Debit

# Create a minimal pipe-delimited CSV
cat > "test_01_01_2026/20260101000000/Easy Pay/credit_successful.csv" << 'EOF'
transaction_id|amount|narration|status|beneficiary_bvn|channel_code|nip_response_code|destination_institution_code|session
TXN001|5000.00|Test payment|00|OPY|2|00|000014|0000
EOF

# Run it (emails are skipped unless .env is configured)
python nibss_etl/main.py test_01_01_2026
```

The pipeline handles missing files gracefully — every gap is logged in the
`Run_Log` sheet and whatever is available gets processed.

---

## The output workbook

Five sheets, written to `nibss_etl/output/NIBSS_Reconciliation_<DD_MM_YYYY>.xlsx`.

### 1. Summary (dashboard)

- **KPI cards** — Total Txns, Successful, Failed, Success Rate, Total Value,
  At-Risk (failed) Value, Value Success Rate, Avg Value, and number of distinct
  beneficiary banks.
- **Session overview** — per-session counts and value splits (successful vs
  failed) with labelled sessions: 5:59 AM, 11:59 AM, 5:59 PM, 11:59 PM.
- **Top banks** — by volume, with a pie chart.
- **Failure analysis** — failed transactions grouped by NIP response code
  (e.g. `00` success, `91` routing error, `96` system malfunction,
  `97` timeout).
- **Purpose patterns** — transaction categories derived from the narration.
- **Channel breakdown** — by channel code, with descriptions.

### 2 & 3. EasyPay / DirectDebit

Filterable Excel tables (AutoFilter) with columns reordered to a priority order:

`session → transaction_type → status → amount → transaction_id →
payment_reference → narration → response_status → message → nip_response_code
→ …` (remaining columns alphabetical).

- Failed rows are highlighted red.
- The `amount` column is currency-formatted (₦).
- Header row + columns A–C are frozen (`D2`).

### 4. Exceptions

Every failed row from both sources combined, with an added `source` column
(`EasyPay` / `DirectDebit`), sorted by amount descending, red-highlighted.

### 5. Run_Log

Pipeline metadata: generation timestamp, source folder, per-sheet row counts,
and a per-file manifest (session, source, file, found/missing, row count, size,
modified time, detail). Schema-drift warnings and dropped columns are logged
here too.

### Column handling

The reader applies these transformations automatically:

| Rule | Effect |
|---|---|
| Original `status` column | renamed to `response_status` |
| Filename (`_failed` / `_successful`) | becomes the canonical `status` column |
| `debit_*` columns | renamed to `originator_*` for a shared schema |
| `initiatior_*` typo | corrected to `initiator_*` |
| Literal `"null"` cells | replaced with empty |
| 100 %-null columns | dropped and logged |

---

## Project structure

```
.
├── run.sh                    # fetch + process (one command)
├── auto_run.sh               # scheduler guard (time + active-session + internet + once-per-day)
├── install_launchd.sh        # install the macOS daily schedule
├── requirements.txt          # pip dependencies
├── .env.example              # copy to .env and fill in
└── nibss_etl/
    ├── fetcher.py            # downloads CSVs from the NIBSS Thin Client web API
    ├── main.py               # CLI entry point (date detection, orchestration)
    └── etl/
        ├── reader.py         # folder walking, CSV parsing, column standardisation
        ├── writer.py         # workbook builder (5 sheets, styling, charts)
        └── mailer.py         # team + MD emails via SMTP
```

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `NIBSS_USER / NIBSS_PASSWORD not set` | `cp .env.example .env` and fill in your credentials. |
| `Emails skipped — missing env vars` | Ensure `EMAIL_USER`, `EMAIL_PASSWORD`, `TEAM_EMAIL` are set in `.env`. |
| `Login failed` from the fetcher | Check your NIBSS username/password; the account must be enabled for the Thin Client. |
| `TOTP/OTP step not supported` | Your NIBSS account requires 2FA — the fetcher can't automate that yet. |
| `Failed to save request context` during fetch | Transient NIBSS-side error — the fetcher retries automatically; just re-run. |
| Report has 0 rows | The day's files may be header-only (no transactions), or you ran before NIBSS published the day's sessions. |
| launchd job logs `Operation not permitted` | Grant Full Disk Access to `/bin/bash` (see [Automating it daily](#automating-it-daily-macos)). |
| launchd job logs `skipping — lid closed` / `skipping — display asleep` | Expected — the job waits for an active session. Open the lid and log in; it runs within 15 minutes (see [Why we require an active session](#why-we-require-an-active-session)). |
| `ModuleNotFoundError: pandas` | Activate the venv or `pip install -r requirements.txt`. |

---

## Notes & limitations

- **Pipe delimiter** — input CSVs use `|`, not commas. The reader handles this
  automatically; don't re-save files as comma-separated.
- **Name-enquiry files** are excluded from the totals (they are beneficiary
  lookups, not transactions).
- **Bank names** are mapped from NIBSS short codes via a dictionary in
  `writer.py` (`BANK_NAME_MAP`); unknown codes display as-is.
- **Session naming** — the pipeline expects NIBSS's `YYYYMMDDHHMMSS` session
  folders and the standard 4 sessions per day.
- **No database** — everything is file-based and runs locally.
- **No license** — internal tool; add one if you make the repo public.
