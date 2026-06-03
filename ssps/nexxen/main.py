#!/usr/bin/env python3
"""
Nexxen Detailed Reporting API → Google Sheets daily revenue sync.

Pulls the last 30 days, grouped by Date + Site/App Bundle, via the official
Nexxen Reports API (ssp-api.nexxen.com). No browser, no captcha.

The script upserts rows into the destination sheet keyed on (Date, Site/App
Bundle): same date+site combos get overwritten with the fresh numbers, older
rows outside the 30-day window are left as-is, newest dates on top.

If anything looks wrong (empty file, missing columns, super-old data) the sheet
is left untouched.
"""

import json
import os
import sys
import time
from datetime import datetime
from io import StringIO
from pathlib import Path

import pandas as pd
import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ── Configuration ─────────────────────────────────────────────────────────────

SCRIPT_DIR     = Path(__file__).parent
CONFIG_PATH    = SCRIPT_DIR / "config.json"
SA_PATH        = SCRIPT_DIR / "service_account.json"
SPREADSHEET_ID = os.environ.get("SHEETS_SPREADSHEET_ID") or "1y6i7-L3qoE7TIsdbaEmV-8D0euVdb3JgHanoLb7U-Jo"
SHEET_NAME     = "Sheet1"
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]

API_BASE       = "https://ssp-api.nexxen.com"
AUTH_URL       = f"{API_BASE}/ctrl/auth"
DETAILED_URL   = f"{API_BASE}/ctrl/api/insights/detailed"
DOWNLOAD_URL   = f"{API_BASE}/ctrl/api/download/fileURL"

DATE_RANGE     = "LAST_30_DAYS"
GROUP_BY       = ["EVENT_TIME", "DOMAIN_BUNDLE"]

# Column-name candidates — actual API CSV headers may differ slightly.
DATE_COL_CANDIDATES = ["Date", "Event Time", "EVENT_TIME", "Day"]
SITE_COL_CANDIDATES = ["Site/App Bundle", "Domain Bundle", "DOMAIN_BUNDLE",
                       "Site / App Bundle", "Site/App bundle", "Site", "Bundle"]
IMPR_COL_CANDIDATES = ["Impressions", "IMPRESSIONS", "Imps"]
REV_COL_CANDIDATES  = ["Est. Earnings (USD)", "Est Earnings (USD)", "EST_EARNING",
                       "Earnings (USD)", "Revenue"]
ECPM_COL_CANDIDATES = ["eCPM (USD)", "eCPM(USD)", "ECPM", "eCPM"]

MAX_ALLOWED_AGE_DAYS = 5      # abort if newest date in CSV older than this
DOWNLOAD_TIMEOUT_S   = 600    # give the report up to 10 min to land in S3
DOWNLOAD_POLL_S      = 20     # poll the download endpoint every N seconds

HEADER = ["Domain", "Date", "Revenue", "Impression", "CPM"]


def _normalize_and_aggregate(rows):
    """
    Collapse URL paths AND subdomains in Domain values to their registered
    root domain (e.g. https://abc.jugantor.com/sports → jugantor.com), then
    sum Impression + Revenue across same (Domain, Date) groups. CPM recomputed
    from totals (Revenue/Impression * 1000) — averaging CPMs would be wrong.

    Uses the Public Suffix List via tldextract so multi-part TLDs like
    .co.uk, .com.br, .gov.in resolve correctly.
    """
    from urllib.parse import urlparse
    try:
        import tldextract
        # suffix_list_urls=() forces use of bundled PSL snapshot only — no
        # network call on first run, important for stateless cloud containers.
        _extract = tldextract.TLDExtract(suffix_list_urls=(), cache_dir=False)
    except ImportError:
        _extract = None

    def root(s):
        s = str(s).strip()
        if "://" in s:
            try:
                s = urlparse(s).netloc or s
            except Exception:
                pass
        s = s.lower().split("/")[0].split(":")[0]
        if not s:
            return ""
        if _extract is not None:
            ext = _extract(s)
            if ext.domain and ext.suffix:
                return f"{ext.domain}.{ext.suffix}"
        # Fallback: strip only "www."
        if s.startswith("www."):
            s = s[4:]
        return s

    bucket = {}
    for r in rows:
        domain = root(r[0])
        date   = str(r[1]).strip()
        try:
            rev = float(str(r[2]).lstrip("$").replace(",", ""))
        except Exception:
            rev = 0.0
        try:
            imp = int(float(str(r[3])))
        except Exception:
            imp = 0
        k = (domain, date)
        if k in bucket:
            bucket[k]["rev"] += rev
            bucket[k]["imp"] += imp
        else:
            bucket[k] = {"rev": rev, "imp": imp}

    new_rows = []
    for (domain, date), v in bucket.items():
        cpm = round(v["rev"] / v["imp"] * 1000, 4) if v["imp"] > 0 else 0.0
        new_rows.append([domain, date, f"${v['rev']:.2f}", v["imp"], cpm])
    return new_rows

# ── Helpers ───────────────────────────────────────────────────────────────────

def log(msg: str) -> None:
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


_REQUIRED_KEYS = ("nexxen_api_username", "nexxen_api_password", "nexxen_api_key")


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



def first_matching_column(df_cols, candidates):
    norm = {c.strip().lower(): c for c in df_cols}
    for cand in candidates:
        if cand.strip().lower() in norm:
            return norm[cand.strip().lower()]
    return None


# ── Step 1 — Talk to the Nexxen API ───────────────────────────────────────────

def login(username: str, password: str) -> str:
    log("Logging into Nexxen API…")
    r = requests.post(AUTH_URL, data={"username": username, "password": password}, timeout=30)
    if r.status_code != 200:
        sys.exit(f"ERROR: Auth failed ({r.status_code}): {r.text[:300]}")
    token = r.json().get("access_token")
    if not token:
        sys.exit(f"ERROR: Auth response had no access_token: {r.text[:300]}")
    return token


def request_detailed_report(token: str, api_key: str) -> str:
    log(f"Requesting Detailed report ({DATE_RANGE}, groupBy={GROUP_BY})…")
    r = requests.post(
        DETAILED_URL,
        params={"dateRange": DATE_RANGE, "apiKey": api_key},
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type":  "application/json",
        },
        json={"groupBy": GROUP_BY, "currency": "USD"},
        timeout=60,
    )
    if r.status_code != 200:
        sys.exit(f"ERROR: Report request failed ({r.status_code}): {r.text[:400]}")
    # Response is the bare filename string (sometimes wrapped in quotes).
    filename = r.text.strip().strip('"')
    if not filename.endswith(".csv"):
        sys.exit(f"ERROR: Unexpected report-request response: {r.text[:300]}")
    log(f"Report queued: {filename}")
    return filename


def fetch_csv(token: str, api_key: str, filename: str) -> pd.DataFrame:
    log("Polling download endpoint…")
    deadline = time.time() + DOWNLOAD_TIMEOUT_S
    s3_url = ""
    while time.time() < deadline:
        r = requests.get(
            DOWNLOAD_URL,
            params={"fileName": filename, "apiKey": api_key},
            headers={"Authorization": f"Bearer {token}"},
            timeout=30,
        )
        if r.status_code != 200:
            sys.exit(f"ERROR: Download URL request failed ({r.status_code}): {r.text[:300]}")
        s3_url = r.json().get("S3_FILE_URL", "")
        if s3_url:
            break
        log(f"  not ready yet, waiting {DOWNLOAD_POLL_S}s…")
        time.sleep(DOWNLOAD_POLL_S)
    if not s3_url:
        sys.exit("ERROR: Report still not available after 10 minutes — aborting.")

    log("Downloading CSV from S3…")
    r = requests.get(s3_url, timeout=120)
    if r.status_code != 200:
        sys.exit(f"ERROR: S3 download failed ({r.status_code}).")
    return pd.read_csv(StringIO(r.text))


# ── Step 2 — Validate & shape ─────────────────────────────────────────────────

def process_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        sys.exit("ERROR: Downloaded CSV is empty — aborting.")
    df.columns = [c.strip() for c in df.columns]

    date_col = first_matching_column(df.columns, DATE_COL_CANDIDATES)
    site_col = first_matching_column(df.columns, SITE_COL_CANDIDATES)
    impr_col = first_matching_column(df.columns, IMPR_COL_CANDIDATES)
    rev_col  = first_matching_column(df.columns, REV_COL_CANDIDATES)
    ecpm_col = first_matching_column(df.columns, ECPM_COL_CANDIDATES)

    missing = [name for name, col in [
        ("Date",            date_col),
        ("Site/App Bundle", site_col),
        ("Impressions",     impr_col),
        ("Revenue",         rev_col),
        ("eCPM",            ecpm_col),
    ] if col is None]
    if missing:
        sys.exit(
            f"ERROR: Could not find columns {missing} in the CSV.\n"
            f"       Found columns: {list(df.columns)}\n"
            "       Aborting to protect the sheet."
        )

    df = df.dropna(subset=[site_col, date_col])
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
            f"ERROR: Newest date in CSV is {newest.date()} ({age_days} days ago).\n"
            f"       Expected data within the last {MAX_ALLOWED_AGE_DAYS} days — aborting."
        )

    # Filter to month-to-date (1st of current month → today).
    first_of_month = pd.Timestamp(datetime.now().replace(day=1).date())
    _pre_mtd_df = df.copy()
    before = len(df)
    df = df[df["__date"] >= first_of_month]
    log(f"Filtered to MTD ({first_of_month.date()} onward): kept {len(df)}, dropped {before - len(df)}.")
    if df.empty:
        log("WARNING: 0 rows match current-month filter — falling back to full report (likely a month-boundary day, MTD data not available yet).")
        df = _pre_mtd_df

    out = pd.DataFrame({
        "Date":            df["__date"].dt.strftime("%Y-%m-%d"),
        "Site/App Bundle": df[site_col].astype(str).str.strip(),
        "Impressions":     pd.to_numeric(df[impr_col], errors="coerce").fillna(0),
        "Revenue":         pd.to_numeric(df[rev_col],  errors="coerce").fillna(0),
        "eCPM":            pd.to_numeric(df[ecpm_col], errors="coerce").fillna(0),
    })
    out = out[~out["Site/App Bundle"].str.lower().isin(["", "nan", "none"])]
    if out.empty:
        sys.exit("ERROR: No valid rows after cleaning — aborting.")

    log(f"CSV valid: {len(out)} rows, dates {out['Date'].min()} → {out['Date'].max()}")
    return out


# ── Step 3 — Upsert into Google Sheets ────────────────────────────────────────

def upsert_to_sheet(df: pd.DataFrame, creds) -> None:
    """MTD-only mode: replace the sheet contents with the freshly-pulled MTD rows."""
    log("Connecting to Google Sheets…")
    service = build("sheets", "v4", credentials=creds, cache_discovery=False)

    rows = [
        [
            str(r["Site/App Bundle"]).strip(),
            str(r["Date"]).strip(),
            f"${float(pd.to_numeric(r['Revenue'], errors='coerce') or 0):.2f}",
            int(pd.to_numeric(r["Impressions"], errors="coerce") or 0),
            float(pd.to_numeric(r["eCPM"], errors="coerce") or 0),
        ]
        for _, r in df.iterrows()
    ]
    rows = _normalize_and_aggregate(rows)
    rows.sort(key=lambda r: (r[0], r[1]))     # Date asc, Site asc
    rows.sort(key=lambda r: r[0], reverse=True)  # then Date desc (stable)

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
    log("=== Nexxen API → Google Sheets daily sync ===")
    cfg   = load_config()
    creds = load_service_account()

    token    = login(cfg["nexxen_api_username"], cfg["nexxen_api_password"])
    filename = request_detailed_report(token, cfg["nexxen_api_key"])
    df       = fetch_csv(token, cfg["nexxen_api_key"], filename)
    df       = process_df(df)
    upsert_to_sheet(df, creds)

    log("=== Done. ===")


if __name__ == "__main__":
    main()
