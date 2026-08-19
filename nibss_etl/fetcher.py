"""
Fetcher — downloads NIBSS EasyPay / Direct Debit CSV exports from the Thin
Client web API.  No FTP, no browser: it logs in and talks to the same JSON
endpoints the Angular portal uses.

Usage
-----
    python fetcher.py --list                 # print the remote file tree
    python fetcher.py 18_08_2026             # download a specific date
    python fetcher.py                        # download today's date

Credentials are read from NIBSS_USER / NIBSS_PASSWORD in the project-root
.env (or the environment).  The password is never logged or printed.
"""

from __future__ import annotations

import argparse
import http.cookiejar
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
from collections import deque
from datetime import datetime, timezone
from pathlib import Path

BASE_URL = "https://nibsswebserver.nibss-plc.com.ng/ThinClient/WtmApiService.asmx"
THINCLIENT_BASE = "https://nibsswebserver.nibss-plc.com.ng/ThinClient/"


def _load_env() -> None:
    """Load NIBSS_* vars from the project-root .env if not already set."""
    env_path = Path(__file__).resolve().parent.parent / ".env"
    if not env_path.exists():
        return
    for line in env_path.read_text().splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            key, _, value = line.partition("=")
            os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def _build_opener() -> urllib.request.OpenerDirector:
    """Return an opener that persists the ASP.NET session cookie across calls."""
    jar = http.cookiejar.CookieJar()
    return urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))


def _post(opener: urllib.request.OpenerDirector, method: str, payload: dict,
          retries: int = 6) -> dict:
    """POST JSON to an ASMX method, retrying transient network/server errors."""
    last: Exception | None = None
    for attempt in range(retries):
        try:
            req = urllib.request.Request(
                f"{BASE_URL}/{method}",
                data=json.dumps(payload).encode("utf-8"),
                headers={"Content-Type": "application/json", "Accept": "application/json"},
            )
            with opener.open(req, timeout=60) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            if exc.code < 500:
                raise
            last = exc
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            last = exc
        time.sleep(1.0 * (attempt + 1))
    raise RuntimeError(f"{method}: request failed after {retries} attempts: {last}")


def _login(opener: urllib.request.OpenerDirector, user: str, password: str) -> None:
    """Authenticate against AuthUser; raises on failure."""
    body = _post(opener, "AuthUser", {"user": user, "pass": password, "rememberMe": True, "lang": "en-US"})
    d = body.get("d") or {}
    if not d.get("IsOk"):
        raise RuntimeError(f"Login failed: {d.get('ErrorMessage')}")
    if d.get("TOTPSetupMode"):
        raise RuntimeError("This account requires a TOTP/OTP step — not yet supported.")


def _list_folder(opener: urllib.request.OpenerDirector, path: str | None) -> list[dict]:
    """List the contents of a folder.  *path* of None means the tree root."""
    if path is None:
        body = _post(opener, "GetFileTree", {})
    else:
        body = _post(opener, "GetFileSubTree", {"subFolderPath": path})
    d = body.get("d") or {}
    if not d.get("IsOk"):
        raise RuntimeError(d.get("ErrorMessage") or "file listing failed")
    return d.get("FilesList") or []


def _full_path(item: dict) -> str:
    """Reconstruct an item's server path from RelativePath + Name."""
    name = item.get("Name") or ""
    rel = item.get("RelativePath")
    if not rel or rel == "/":
        return "/" + name
    return rel.rstrip("/") + "/" + name


def _iso(value: str | None) -> str | None:
    """Convert a .NET ``/Date(ms)/`` string to ISO-8601 (UTC, ms precision)."""
    if not value:
        return value
    m = re.search(r"(\d+)", value)
    if not m:
        return value
    ts = datetime.fromtimestamp(int(m.group(1)) / 1000, tz=timezone.utc)
    return ts.strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + "Z"


def _download_file(opener: urllib.request.OpenerDirector, item: dict, dest: Path,
                   attempts: int = 4) -> bool:
    """Resolve *item*'s download link and save it to *dest*.  Returns success."""
    name = item.get("Name") or ""
    file_info = {
        "Name": name,
        "Created": _iso(item.get("Created")),
        "Size": item.get("Size"),
        "FileType": item.get("FileType"),
        "FullPath": _full_path(item),
        "IsFolder": item.get("IsFolder", False),
        "RelativePath": item.get("RelativePath"),
        "SessionStatGuid": item.get("SessionStatGuid"),
        "LastModified": _iso(item.get("LastModified")),
    }
    last_error: str | None = None
    for _ in range(attempts):
        body = _post(opener, "GetDownloadingFileInfo", {"fileInfo": file_info})
        d = body.get("d") or {}
        if not d.get("IsOk"):
            last_error = d.get("ErrorMessage")
            time.sleep(2)
            continue
        link = d.get("DownloadLink") or ""
        url = link if link.startswith("http") else THINCLIENT_BASE + link.lstrip("/")
        try:
            with opener.open(urllib.request.Request(url), timeout=120) as resp:
                dest.parent.mkdir(parents=True, exist_ok=True)
                dest.write_bytes(resp.read())
            return True
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            last_error = str(exc)
            time.sleep(2)
    print(f"  ! {name}: {last_error}")
    return False


def _find_date_folder(opener: urllib.request.OpenerDirector, date: str,
                      max_depth: int = 8) -> dict | None:
    """Find the server folder matching *date* (DD_MM_YYYY); root-first."""
    target_dd = re.sub(r"\D", "", date)
    target_yy = date[6:10] + date[3:5] + date[0:2]

    def matches(name: str) -> bool:
        return re.sub(r"\D", "", name or "") in (target_dd, target_yy)

    root = _list_folder(opener, None)
    for item in root:
        if item.get("IsFolder") and matches(item.get("Name") or ""):
            return item

    # Fallback: the date folder may be nested below other folders.
    queue: deque[tuple[dict, int]] = deque((i, 1) for i in root if i.get("IsFolder"))
    while queue:
        item, depth = queue.popleft()
        if matches(item.get("Name") or ""):
            return item
        if depth >= max_depth:
            continue
        try:
            children = _list_folder(opener, _full_path(item))
        except RuntimeError as exc:
            print(f"  ! {item.get('Name')}: {exc}")
            continue
        queue.extend((c, depth + 1) for c in children if c.get("IsFolder"))
    return None


def _list_tree(opener: urllib.request.OpenerDirector, max_depth: int) -> None:
    """Print the remote tree (folders end in '/'), bounded by *max_depth*."""

    def walk(items: list[dict], depth: int, prefix: str) -> None:
        for item in items:
            name = item.get("Name") or ""
            suffix = "/" if item.get("IsFolder") else f"  [{item.get('Size')}]"
            print(f"{prefix}{name}{suffix}")
            if item.get("IsFolder") and depth < max_depth:
                try:
                    walk(_list_folder(opener, _full_path(item)), depth + 1, prefix + "  ")
                except RuntimeError as exc:
                    print(f"{prefix}  ! {exc}")

    walk(_list_folder(opener, None), 0, "")


def _fetch(opener: urllib.request.OpenerDirector, project_root: Path, date: str) -> int:
    """Download every CSV under the server folder matching *date*."""
    print(f"Looking for date folder {date} …")
    date_item = _find_date_folder(opener, date)
    if date_item is None:
        print(f"No folder matching {date} found on the server. Run --list to inspect naming.")
        return 1

    out_root = project_root / date
    downloaded = 0
    stack: list[tuple[dict, Path]] = [(date_item, Path("."))]
    while stack:
        item, rel = stack.pop()
        if item.get("IsFolder"):
            try:
                children = _list_folder(opener, _full_path(item))
            except RuntimeError as exc:
                print(f"  ! {item.get('Name')}: {exc}")
                continue
            stack.extend((child, rel / (child.get("Name") or "")) for child in children)
        else:
            name = item.get("Name") or ""
            if name.lower().endswith(".csv"):
                dest = out_root / rel
                if dest.exists() and str(dest.stat().st_size) == str(item.get("Size")):
                    continue
                if _download_file(opener, item, dest):
                    downloaded += 1
                    print(f"  ✓ {rel}")

    print(f"Downloaded {downloaded} file(s) to {out_root}/")
    return 0


def main() -> None:
    parser = argparse.ArgumentParser(description="NIBSS Thin Client CSV fetcher")
    parser.add_argument("date", nargs="?", default=None, help="Date folder in DD_MM_YYYY format")
    parser.add_argument("--list", action="store_true", help="List the remote file tree and exit")
    parser.add_argument("--depth", type=int, default=8, help="Max depth for --list")
    args = parser.parse_args()

    _load_env()
    user = os.environ.get("NIBSS_USER", "").strip()
    password = os.environ.get("NIBSS_PASSWORD", "").strip()
    if not user or not password:
        print("NIBSS_USER / NIBSS_PASSWORD not set in .env", file=sys.stderr)
        sys.exit(1)

    opener = _build_opener()
    _login(opener, user, password)
    print("Logged in OK")

    if args.list:
        _list_tree(opener, args.depth)
        return

    date = args.date or datetime.now().strftime("%d_%m_%Y")
    project_root = Path(__file__).resolve().parent.parent
    sys.exit(_fetch(opener, project_root, date))


if __name__ == "__main__":
    main()
