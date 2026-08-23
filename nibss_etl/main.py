"""
NIBSS EasyPay & Direct Debit reconciliation ETL pipeline.

Usage
-----
    python main.py                            # process today's data
    python main.py 29_07_2026                 # process a specific date folder

The pipeline reads pipe-delimited CSV files from the NIBSS folder structure,
builds a single 5-sheet workbook, and writes it to ``output/``.
"""

import argparse
import os
import sys
from pathlib import Path
from datetime import datetime
from etl.reader import load_all_data
from etl.writer import build_workbook
from etl.mailer import send_emails


def _load_dotenv() -> None:
    """Load EMAIL_* vars from .env in the project root, if present (run.sh does this too)."""
    env_path = Path(__file__).resolve().parent.parent / ".env"
    if not env_path.exists():
        return
    for line in env_path.read_text().splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            key, _, value = line.partition("=")
            os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def _most_recent_folder() -> str:
    """
    Scan all DD_MM_YYYY folders under the project root and return
    the one with the most recent modification time.
    """
    import re

    root = Path(__file__).resolve().parent.parent
    pattern = re.compile(r"^\d{2}_\d{2}_\d{4}$")
    candidates = [(d.stat().st_mtime, d.name) for d in root.iterdir() if d.is_dir() and pattern.match(d.name)]
    if not candidates:
        return datetime.now().strftime("%d_%m_%Y")
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]


def _print_summary(output_path: str, data: dict, run_log: list):
    """Print a brief run summary to stdout."""
    ep_rows = len(data.get("easy_pay", []))
    dd_rows = len(data.get("direct_debit", []))
    files_found = sum(1 for e in run_log if e.get("found"))
    files_total = len(run_log)

    print(f"✓  Workbook written → {output_path}")
    print(f"   EasyPay:    {ep_rows} rows")
    print(f"   DirectDebit: {dd_rows} rows")
    print(f"   Files:       {files_found}/{files_total} found")
    print(f"   Run_Log:     {len(run_log)} entries")


def main():
    parser = argparse.ArgumentParser(
        description="NIBSS EasyPay & Direct Debit reconciliation ETL"
    )
    parser.add_argument(
        "date_folder",
        nargs="?",
        default=None,
        help="Date folder name in DD_MM_YYYY format (default: today or most recent)",
    )
    args = parser.parse_args()

    _load_dotenv()

    project_root = Path(__file__).resolve().parent.parent
    date_folder = args.date_folder or _most_recent_folder()
    root_path = str(project_root / date_folder)

    print(f"Reading: {root_path}")

    data, run_log = load_all_data(root_path)

    if data["easy_pay"].empty and data["direct_debit"].empty:
        print("No data found — nothing to write.")
        sys.exit(1)

    output_name = f"NIBSS_Reconciliation_{date_folder}.xlsx"
    output_path = str(project_root / "nibss_etl" / "output" / output_name)

    build_workbook(data, run_log, root_path, output_path)
    _print_summary(output_path, data, run_log)
    if send_emails(data, output_path, date_folder) is False:
        sys.exit(2)


if __name__ == "__main__":
    main()
