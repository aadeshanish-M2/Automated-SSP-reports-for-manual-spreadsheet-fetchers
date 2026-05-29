#!/usr/bin/env python3
"""
MGID Publisher Reports API → Google Sheets daily revenue sync.

Pulls month-to-date data grouped by Date + Website via MGID's
v2 website-custom-report endpoint, computes CPM = Revenue/Clicks*1000,
and replaces the destination sheet contents (MTD-only mode).

If the API response is empty, missing expected fields, or stale, the sheet
is left untouched.
"""

import json
import os
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ── Configuration ─────────────────────────────────────────────────────────────

SCRIPT_DIR     = Path(__file__).parent
CONFIG_PATH    = SCRIPT_DIR / "config.json"
SA_PATH        = SCRIPT_DIR / "service_account.json"
SPREADSHEET_ID = os.environ.get("SHEETS_SPREADSHEET_ID") or "1cUNyk4Zrx2ZTZCjVl0O97zcNqwefr3yofJCz88HJpEc"
SHEET_NAME     = "Sheet1"
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]

API_BASE       = "https://api.mgid.com/v2/pub/account"
DATE_INTERVAL  = "thisMonth"      # MGID preset = month-to-date
DIMENSIONS     = "date,website"
METRICS        = "clicks,revenue"

MAX_ALLOWED_AGE_DAYS = 5

HEADER = ["Date", "Website", "Clicks", "Revenue", "CPM"]


# ── Helpers ───────────────────────────────────────────────────────────────────

def log(msg: str) -> None:
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


_REQUIRED_KEYS = ("mgid_client_id", "mgid_api_token")


def load_config() -> dict:
    """Credentials: env vars first (UPPER_CASE), then config.json fallback."""
    cfg = {}
    for k in _REQUIRED_KEYS:
        v = os.environ.get(k.upper())
        if v:
            cfg[k] = v
    if CONFIG_PATH.exists():
        try:
            file_cfg = json.loads(CONFIG_PATH.read_text())
            for k in _REQUIRED_KEYS:
                if not cfg.get(k) and file_cfg.get(k):
                    cfg[k] = file_cfg[k]
        except Exception:
            pass
    missing = [k for k in _REQUIRED_KEYS if not cfg.get(k)]
    if missing:
        env_names = ", ".join(k.upper() for k in missing)
        sys.exit(
            f"ERROR: Missing credentials {missing}. "
            f"Set env vars [{env_names}] or fill config.json."
        )
    return cfg



def load_service_account():
    """Service account: env var GOOGLE_SERVICE_ACCOUNT_JSON first, then file."""
    sa_json_str = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if sa_json_str:
        try:
            sa_info = json.loads(sa_json_str)
            return service_account.Credentials.from_service_account_info(
                sa_info, scopes=SCOPES
            )
        except Exception as e:
            sys.exit(f"ERROR: GOOGLE_SERVICE_ACCOUNT_JSON env var invalid: {e}")
    if not SA_PATH.exists():
        sys.exit(f"ERROR: {SA_PATH} not found and GOOGLE_SERVICE_ACCOUNT_JSON not set.")
    sa = json.loads(SA_PATH.read_text())
    if not sa.get("client_email") or not sa.get("private_key"):
        sys.exit("ERROR: service_account.json is not populated.")
    return service_account.Credentials.from_service_account_file(
        str(SA_PATH), scopes=SCOPES
    )


# ── Step 1 — Fetch report from MGID API ───────────────────────────────────────

def fetch_report(client_id: str, token: str) -> list[dict]:
    url = f"{API_BASE}/{client_id}/website-custom-report"
    params = {
        "dateInterval": DATE_INTERVAL,
        "dimensions":   DIMENSIONS,
        "metrics":      METRICS,
    }
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept":        "application/json",
    }
    log(f"Requesting MGID report ({DATE_INTERVAL}, dims={DIMENSIONS}, metrics={METRICS})…")
    r = requests.get(url, params=params, headers=headers, timeout=60)
    if r.status_code != 200:
        sys.exit(f"ERROR: MGID API failed ({r.status_code}): {r.text[:400]}")
    try:
        data = r.json()
    except ValueError:
        sys.exit(f"ERROR: MGID response was not valid JSON: {r.text[:300]}")

    # The API can return either a list of rows directly, or {"data": [...]},
    # or {"errors": [...]} — handle each.
    if isinstance(data, dict):
        if data.get("errors"):
            sys.exit(f"ERROR: MGID returned errors: {data['errors']}")
        if isinstance(data.get("data"), list):
            data = data["data"]
        else:
            # Sometimes the payload is wrapped in {"result": [...]} or similar.
            for k in ("result", "items", "rows"):
                if isinstance(data.get(k), list):
                    data = data[k]
                    break
            else:
                sys.exit(f"ERROR: Unexpected MGID response shape: {str(data)[:300]}")

    if not isinstance(data, list):
        sys.exit(f"ERROR: Expected list of rows from MGID, got {type(data).__name__}.")

    log(f"MGID returned {len(data)} rows.")
    return data


# ── Step 2 — Validate & shape ─────────────────────────────────────────────────

def process_rows(rows: list[dict]) -> pd.DataFrame:
    if not rows:
        sys.exit("ERROR: MGID returned 0 rows — aborting to protect the sheet.")

    df = pd.DataFrame(rows)
    df.columns = [c.strip() for c in df.columns]

    # Field-name candidates — MGID's exact casing isn't documented.
    def find_col(*cands):
        norm = {c.lower(): c for c in df.columns}
        for c in cands:
            if c.lower() in norm:
                return norm[c.lower()]
        return None

    date_col    = find_col("date", "day")
    website_col = find_col("website", "domain", "site")
    clicks_col  = find_col("clicks")
    rev_col     = find_col("revenue", "earnings")

    missing = [name for name, col in [
        ("date", date_col), ("website", website_col),
        ("clicks", clicks_col), ("revenue", rev_col),
    ] if col is None]
    if missing:
        sys.exit(
            f"ERROR: Missing expected fields {missing} in MGID response.\n"
            f"       Got columns: {list(df.columns)}"
        )

    df = df.dropna(subset=[website_col, date_col])
    parsed = pd.to_datetime(df[date_col], errors="coerce")
    bad = parsed.isna().sum()
    if bad:
        log(f"WARNING: {bad} row(s) had unparseable dates and will be dropped.")
    df = df[parsed.notna()].copy()
    df["__date"] = pd.to_datetime(df[date_col], errors="coerce")

    if df.empty:
        sys.exit("ERROR: No valid rows after date parsing — aborting.")

    newest = df["__date"].max()
    age_days = (datetime.now() - newest).days
    if age_days > MAX_ALLOWED_AGE_DAYS:
        sys.exit(
            f"ERROR: Newest date is {newest.date()} ({age_days} days ago). "
            f"Expected within {MAX_ALLOWED_AGE_DAYS} days — aborting."
        )

    # MGID's thisMonth preset already returns MTD, but defensively re-filter
    # in case the preset behaviour ever shifts.
    first_of_month = pd.Timestamp(datetime.now().replace(day=1).date())
    before = len(df)
    df = df[df["__date"] >= first_of_month]
    if before != len(df):
        log(f"Filtered to MTD ({first_of_month.date()} onward): kept {len(df)}, dropped {before - len(df)}.")

    clicks  = pd.to_numeric(df[clicks_col], errors="coerce").fillna(0)
    revenue = pd.to_numeric(df[rev_col],    errors="coerce").fillna(0)

    out = pd.DataFrame({
        "Date":    df["__date"].dt.strftime("%Y-%m-%d"),
        "Website": df[website_col].astype(str).str.strip(),
        "Clicks":  clicks.astype(int),
        "Revenue": revenue.round(4),
        # CPM = Revenue / Clicks * 1000 (per spec); 0 when Clicks == 0.
        "CPM":     (revenue.where(clicks > 0, 0) / clicks.where(clicks > 0, 1) * 1000).round(4),
    })
    out = out[~out["Website"].str.lower().isin(["", "nan", "none"])]
    if out.empty:
        sys.exit("ERROR: No valid rows after cleaning — aborting.")

    log(f"Valid: {len(out)} rows, dates {out['Date'].min()} → {out['Date'].max()}")
    return out


# ── Step 3 — Replace sheet contents ───────────────────────────────────────────

def write_sheet(df: pd.DataFrame, creds) -> None:
    log("Connecting to Google Sheets…")
    service = build("sheets", "v4", credentials=creds, cache_discovery=False)

    rows = [
        [r["Date"], r["Website"], r["Clicks"], r["Revenue"], r["CPM"]]
        for _, r in df.iterrows()
    ]
    rows.sort(key=lambda r: (r[0], r[1]))
    rows.sort(key=lambda r: r[0], reverse=True)

    final_rows = [HEADER] + rows
    log(f"Writing {len(rows)} MTD rows to sheet (replacing existing)…")
    service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID, range=SHEET_NAME, body={},
    ).execute()
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!A1",
        valueInputOption="USER_ENTERED",
        body={"values": final_rows},
    ).execute()
    log("Sheet updated successfully.")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    log("=== MGID API → Google Sheets daily sync ===")
    cfg   = load_config()
    creds = load_service_account()

    rows = fetch_report(cfg["mgid_client_id"], cfg["mgid_api_token"])
    df   = process_rows(rows)
    write_sheet(df, creds)
    log("=== Done. ===")


if __name__ == "__main__":
    main()
