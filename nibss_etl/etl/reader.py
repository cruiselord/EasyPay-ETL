"""
Reader module for the NIBSS ETL pipeline.

Responsible for:
  1. Discovering session folders inside a date root folder
  2. Reading pipe-delimited CSV files and cleaning them
  3. Combining all CSVs across sessions into unified DataFrames

Designed so that folder discovery can be replaced with an FTP fetch step
without touching the CSV cleaning or DataFrame construction logic.
"""

from pathlib import Path
from datetime import datetime
import pandas as pd


# Files expected inside each session's Easy Pay folder.
# Name-enquiry files are deliberately excluded: they are beneficiary
# lookups, not transactions, and would inflate every total.
EASYPAY_FILES = [
    "credit_failed.csv",
    "credit_successful.csv",
    "debit_failed.csv",
    "debit_successful.csv",
]

# Files expected inside each session's Direct Debit folder
DIRECTDEBIT_FILES = [
    "credit_failed.csv",
    "credit_successful.csv",
    "debit_failed.csv",
    "debit_successful.csv",
]

# Map debit-column names to originator-column names so credit & debit
# rows share the same schema when combined into a single table.
DEBIT_TO_ORIGINATOR = {
    "debit_account_name": "originator_account_name",
    "debit_account_number": "originator_account_number",
    "debit_bvn": "originator_bvn",
    "debit_kyc_level": "originator_kyc_level",
}

# Known NIBSS typos in column headers — fix silently.
COLUMN_TYPO_FIXES = {
    "initiatior_account_name": "initiator_account_name",
    "initiatior_account_number": "initiator_account_number",
}

# Sort-order mappings for explicit row ordering (not implicit from file-read order).
_SESSION_ORDER = {"0559": 0, "1159": 1, "1759": 2, "2359": 3}
_TYPE_ORDER = {"credit": 0, "debit": 1}
_STATUS_ORDER = {"failed": 0, "successful": 1}


# ---------------------------------------------------------------------------
# Session discovery
# ---------------------------------------------------------------------------

def discover_sessions(root_path: str) -> list[dict]:
    """
    Walk *root_path* and find session folders by their YYYYMMDDHHMMSS name.

    Returns a list sorted by timestamp ascending.  Each dict contains:
        path             – Path to the session folder
        session          – human-readable label e.g. "0559", "1159"
        easy_pay_path    – Path to "Easy Pay/" (or None)
        direct_debit_path – Path to "Direct Debit/" (or None)

    Returns an empty list when *root_path* is invalid or contains no
    recognisable session folders — never raises.
    """
    root = Path(root_path)
    if not root.is_dir():
        return []

    sessions = []
    for entry in sorted(root.iterdir()):
        if not entry.is_dir() or entry.name.startswith("."):
            continue
        if len(entry.name) != 14 or not entry.name.isdigit():
            continue

        # Derive label from the HHMM portion of the 14-digit timestamp.
        hhmm = entry.name[8:12]
        label = {"0559": "0559", "1159": "1159",
                 "1759": "1759", "2359": "2359"}.get(hhmm, hhmm)

        easy_pay = entry / "Easy Pay"
        direct_debit = entry / "Direct Debit"

        if not easy_pay.is_dir() and not direct_debit.is_dir():
            continue

        sessions.append({
            "path": entry,
            "session": label,
            "easy_pay_path": easy_pay if easy_pay.is_dir() else None,
            "direct_debit_path": direct_debit if direct_debit.is_dir() else None,
        })

    return sessions


# ---------------------------------------------------------------------------
# Low-level CSV reading and cleaning
# ---------------------------------------------------------------------------

def _standardize_column_names(columns: list) -> list:
    """Lowercase, strip whitespace, replace spaces with underscores, fix typos."""
    cleaned = []
    for col in columns:
        col = col.strip().lower().replace(" ", "_")
        col = COLUMN_TYPO_FIXES.get(col, col)
        cleaned.append(col)
    return cleaned


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Rename debit-columns to originator-columns so credit and debit
    DataFrames share a common schema.  Drops any duplicate columns
    that arise after the rename.
    """
    df = df.rename(columns=DEBIT_TO_ORIGINATOR)
    df = df.loc[:, ~df.columns.duplicated()]
    return df


def _rename_original_status(df: pd.DataFrame) -> pd.DataFrame:
    """
    The CSV files contain their own *status* column (e.g. "COMPLETED").
    We prepend our own filename-derived *status* column later, so rename
    the original one to *response_status* to avoid a name clash.
    """
    if "status" in df.columns:
        df = df.rename(columns={"status": "response_status"})
    return df


def _drop_junk_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Some NIBSS CSVs have extra trailing fields (literal "null" columns).
    Drop any column whose name is empty, "null", "none", or starts with
    "unnamed".
    """
    junk = {"", "null", "none"}
    cols_to_drop = [
        c for c in df.columns
        if c.strip().lower() in junk or c.lower().startswith("unnamed")
    ]
    if cols_to_drop:
        df = df.drop(columns=cols_to_drop)
    return df


def read_csv(filepath: Path) -> pd.DataFrame | None:
    """
    Read a pipe-delimited CSV, clean it, and return a DataFrame.

    Cleaning steps:
      * Strip whitespace from headers and all string cells
      * Lowercase + snake_case + fix-typo headers
      * Debit columns → originator columns
      * Original ``status`` → ``response_status``
      * Remove junk trailing columns
      * Parse ``amount`` to float
      * Replace literal ``"null"`` strings with ``pd.NA``
      * Keep ``transaction_location`` as raw string (lat/lon split removed —
        values are always placeholder GPS zeros)

    Returns ``None`` when the file is missing, empty, or unreadable
    (never raises).
    """
    if not filepath.is_file() or filepath.stat().st_size == 0:
        return None

    try:
        df = pd.read_csv(
            filepath,
            delimiter="|",
            dtype=str,
            keep_default_na=False,
            encoding="utf-8",
        )
    except Exception:
        return None

    if df.empty:
        return None

    # Standardise headers
    df.columns = _standardize_column_names(df.columns.tolist())

    # Strip whitespace from every string cell
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x)

    # Replace literal "null" with pandas NA
    df = df.replace("null", pd.NA)

    # Remove extra trailing columns (e.g. name_enquiry "null|null")
    df = _drop_junk_columns(df)

    # Normalise debit → originator
    df = _normalize_columns(df)

    # Rename original status column so it doesn't clash with prepended one
    df = _rename_original_status(df)

    # Parse amount column (strip any stray non-numeric chars)
    if "amount" in df.columns:
        df["amount"] = (
            df["amount"]
            .str.replace(r"[^\d.]", "", regex=True)
            .replace("", pd.NA)
            .astype(float)
        )

    # NOTE: transaction_location is *not* split into lat/lon anymore.
    # Those values were always "00.00000,-00.00000" (placeholder zeros)
    # and never carried real coordinate data.  Keep the raw string column
    # for reference; the dead-column audit in the writer will flag it if
    # it turns out to be all-NaN (unlikely since every row has a value).

    return df


# ---------------------------------------------------------------------------
# Filename classification helpers
# ---------------------------------------------------------------------------

def _derive_transaction_type(filename: str) -> str:
    """Return ``credit``, ``debit``, or ``name_enquiry`` based on *filename*."""
    name = filename.lower()
    if name.startswith("name_enquiry"):
        return "name_enquiry"
    if name.startswith("credit"):
        return "credit"
    if name.startswith("debit"):
        return "debit"
    return name.replace("_failed.csv", "").replace("_successful.csv", "")


def _derive_status(filename: str) -> str:
    """Return ``successful`` or ``failed`` based on the *filename* suffix."""
    name = filename.lower()
    if "_failed.csv" in name:
        return "failed"
    if "_successful.csv" in name:
        return "successful"
    return name


# ---------------------------------------------------------------------------
# Schema drift detection
# ---------------------------------------------------------------------------

def _log_schema_drift(frames: list[pd.DataFrame], filenames: list[str],
                      session_label: str, source: str, run_log: list):
    """
    Compare column sets across source-file DataFrames within one session.
    Log any column that appears in only some of the files so silent
    all-null columns in the union don't go unnoticed.
    """
    if len(frames) < 2:
        return

    col_sets = dict(zip(filenames, [set(df.columns) for df in frames]))
    all_cols = set.union(*col_sets.values())

    for fname, cols in col_sets.items():
        missing = sorted(all_cols - cols)
        if missing:
            run_log.append({
                "session": session_label,
                "source": source,
                "file": f"(schema drift: {fname})",
                "found": True,
                "rows": 0,
                "detail": f"Missing vs sibling files: {missing}",
            })


# ---------------------------------------------------------------------------
# Per-source-per-session loader
# ---------------------------------------------------------------------------

def load_source(session_info: dict, source: str, run_log: list) -> pd.DataFrame:
    """
    Load every CSV for *source* (``"EasyPay"`` or ``"DirectDebit"``) from
    one session and return a single DataFrame with prepended columns:

        session, transaction_type, status

    Each expected file is logged to *run_log* regardless of whether it was
    found, empty, or contained data.  Returns an empty DataFrame when no
    files were readable.
    """
    if source == "EasyPay":
        folder = session_info["easy_pay_path"]
        expected = EASYPAY_FILES
    else:
        folder = session_info["direct_debit_path"]
        expected = DIRECTDEBIT_FILES

    label = session_info["session"]
    frames: list[pd.DataFrame] = []
    file_order: list[str] = []  # parallel to frames for schema-drift logging

    if folder is None:
        run_log.append({
            "session": label,
            "source": source,
            "file": "(folder missing)",
            "found": False,
            "rows": 0,
            "size_kb": None,
            "modified": None,
            "detail": None,
        })
        return pd.DataFrame()

    for fname in expected:
        filepath = folder / fname

        # Capture file-system metadata before reading
        size_kb = None
        modified = None
        if filepath.is_file():
            try:
                stat = filepath.stat()
                size_kb = round(stat.st_size / 1024, 1)
                modified = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            except OSError:
                pass

        df = read_csv(filepath)
        txn_type = _derive_transaction_type(fname)
        status = _derive_status(fname)

        if df is not None and not df.empty:
            df.insert(0, "status", status)
            df.insert(0, "transaction_type", txn_type)
            df.insert(0, "session", label)
            frames.append(df)
            file_order.append(fname)
            run_log.append({
                "session": label,
                "source": source,
                "file": fname,
                "found": True,
                "rows": len(df),
                "size_kb": size_kb,
                "modified": modified,
                "detail": None,
            })
        else:
            run_log.append({
                "session": label,
                "source": source,
                "file": fname,
                "found": False,
                "rows": 0,
                "size_kb": size_kb,
                "modified": modified,
                "detail": "empty" if (filepath.is_file() and filepath.stat().st_size > 0) else "missing/empty",
            })

    # Check for column mismatches across source files before merging
    if len(frames) > 1:
        _log_schema_drift(frames, file_order, label, source, run_log)

    return pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()


# ---------------------------------------------------------------------------
# Explicit sort
# ---------------------------------------------------------------------------

def _sort_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Sort rows deterministically: session (0559 → 2359), transaction_type
    (credit → debit → name_enquiry), status (failed before successful).
    """
    if df.empty:
        return df
    df = df.copy()
    df["_sort_s"] = df["session"].map(lambda x: _SESSION_ORDER.get(x, 99))
    df["_sort_t"] = df["transaction_type"].map(lambda x: _TYPE_ORDER.get(x, 99))
    df["_sort_st"] = df["status"].map(lambda x: _STATUS_ORDER.get(x, 99))
    df = df.sort_values(["_sort_s", "_sort_t", "_sort_st"])
    return df.drop(columns=["_sort_s", "_sort_t", "_sort_st"]).reset_index(drop=True)


# ---------------------------------------------------------------------------
# Cross-source deduplication
# ---------------------------------------------------------------------------

def _dedup_against_easypay(easy_pay: pd.DataFrame, direct_debit: pd.DataFrame) -> pd.DataFrame:
    """
    Drop DirectDebit rows that are the same successful payment as an
    EasyPay row.  NIBSS reports every successful outward payment twice —
    once as an EasyPay credit and once as a DirectDebit debit — so both
    copies are identical (same transaction_id, amount, beneficiary).

    The EasyPay row is kept (richer data).  Failed rows and rows with a
    missing/empty transaction_id are never touched — a row that can't be
    matched on a valid ID is kept rather than risk losing a transaction.
    """
    if direct_debit.empty or easy_pay.empty:
        return direct_debit
    if "transaction_id" not in easy_pay.columns or "transaction_id" not in direct_debit.columns:
        return direct_debit

    ep_ids = set(easy_pay.loc[
        (easy_pay["status"] == "successful")
        & (easy_pay["transaction_id"].notna())
        & (easy_pay["transaction_id"].astype(str).str.strip() != ""),
        "transaction_id",
    ])
    if not ep_ids:
        return direct_debit

    dd_ids = direct_debit["transaction_id"].astype(str).str.strip()
    valid = dd_ids.notna() & (dd_ids != "") & (dd_ids != "nan")
    is_dupe = valid & direct_debit["transaction_id"].isin(ep_ids)
    return direct_debit.loc[~is_dupe].reset_index(drop=True)


# ---------------------------------------------------------------------------
# Top-level loader
# ---------------------------------------------------------------------------

def load_all_data(root_path: str) -> tuple:
    """
    Entry point for data loading.

    Discovers sessions inside *root_path*, reads all CSVs, sorts rows
    deterministically, and returns a tuple ``(data, run_log)`` where::

        data    = {"easy_pay": DataFrame, "direct_debit": DataFrame}
        run_log = list[dict]   # one entry per expected file + schema drift

    Both DataFrames contain every row across all four sessions, with
    the prepended columns *session*, *transaction_type*, and *status*,
    sorted by session → type → status.  DirectDebit rows duplicating an
    EasyPay successful payment are removed (see _dedup_against_easypay).
    """
    sessions = discover_sessions(root_path)
    if not sessions:
        return {"easy_pay": pd.DataFrame(), "direct_debit": pd.DataFrame()}, []

    run_log: list[dict] = []
    ep_frames: list[pd.DataFrame] = []
    dd_frames: list[pd.DataFrame] = []

    for session in sessions:
        ep = load_source(session, "EasyPay", run_log)
        if not ep.empty:
            ep_frames.append(ep)

        dd = load_source(session, "DirectDebit", run_log)
        if not dd.empty:
            dd_frames.append(dd)

    data = {
        "easy_pay": _sort_dataframe(
            pd.concat(ep_frames, ignore_index=True, sort=False)
        ) if ep_frames else pd.DataFrame(),
        "direct_debit": _sort_dataframe(
            pd.concat(dd_frames, ignore_index=True, sort=False)
        ) if dd_frames else pd.DataFrame(),
    }

    # The same successful payment appears in both reports (EasyPay credit +
    # DirectDebit debit); drop the DirectDebit copy so totals aren't doubled.
    data["direct_debit"] = _sort_dataframe(
        _dedup_against_easypay(data["easy_pay"], data["direct_debit"])
    )

    return data, run_log
