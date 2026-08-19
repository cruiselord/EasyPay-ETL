#!/bin/bash
# NIBSS ETL — single entry point.
# 1. Fetches the day's CSV exports from the NIBSS Thin Client web API (no FTP).
# 2. Builds the reconciliation workbook and sends the report emails.
# Email + portal credentials are loaded from .env in the project root.
set -e
cd "$(dirname "$0")"
if [ -f .env ]; then
    set -a
    . ./.env
    set +a
fi
cd nibss_etl
source venv/bin/activate

# Fetch today's folder (or the date passed in) from the Thin Client API.
python fetcher.py "$@"

# Build the workbook + send emails for the fetched data.
python main.py "$@"
