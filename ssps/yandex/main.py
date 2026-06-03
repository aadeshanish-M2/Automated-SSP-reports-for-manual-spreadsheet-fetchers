#!/usr/bin/env python3
"""
Yandex Partner Statistics API → Google Sheets daily sync.

Pulls month-to-date data grouped by Date + Domain via the official Yandex
Partner Statistics API (partner.yandex.ru/api/statistics2/get.json) and
replaces the destination sheet contents (MTD-only mode).

If the response is empty, missing fields, or stale, the sheet is left
untouched.
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
SPREADSHEET_ID = os.environ.get("SHEETS_SPREADSHEET_ID") or "1p8t2bda942GOIRWGx6SGYwhRRLiyV5fkumXuw5TULjQ"
SHEET_NAME     = "Sheet1"
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]

API_URL        = "https://partner.yandex.ru/api/statistics2/get.json"
DATE_PERIOD    = "thismonth"      # MTD preset
IMPR_FIELD     = "impressions"
REV_FIELD      = "partner_wo_nds"
ECPM_FIELD     = "ecpm_partner_wo_nds"

MAX_ALLOWED_AGE_DAYS = 5
HEADER = ["Domain", "Date", "Revenue", "Impression", "CPM"]


def _normalize_and_aggregate(rows):
    """
    Collapse URL / subdomain Domain values to their root domain, then sum
    Impression + Revenue across same (Domain, Date) groups. CPM is recomputed
    from totals (Revenue/Impression * 1000) — averaging CPMs would be wrong.
    """
    from urllib.parse import urlparse

    def root(s):
        s = str(s).strip()
        if "://" in s:
            try:
                s = urlparse(s).netloc or s
            except Exception:
                pass
        s = s.lower().split("/")[0].split(":")[0]
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


_REQUIRED_KEYS = ("yandex_oauth_token",)


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
    return service_account.Credentials.from_service_account_file(str(SA_PATH), scopes=SCOPES)


# ── Step 1 — Fetch report from Yandex API ─────────────────────────────────────

def fetch_report(token: str) -> list[dict]:
    log(f"Requesting Yandex report (period={DATE_PERIOD}, dims=date+domain, "
        f"metrics={IMPR_FIELD},{REV_FIELD},{ECPM_FIELD})…")
    # Note: 'field' appears multiple times in the URL — requests handles a list value.
    params = [
        ("lang", "en"),
        ("period", DATE_PERIOD),
        ("dimension_field", "date|day"),
        ("entity_field", "domain"),
        ("field", IMPR_FIELD),
        ("field", REV_FIELD),
        ("field", ECPM_FIELD),
        ("currency", "USD"),     # convert RUB → USD at Yandex's rate
        ("limit", "10000"),
    ]
    headers = {"Authorization": f"OAuth {token}"}
    r = requests.get(API_URL, params=params, headers=headers, timeout=60)
    if r.status_code != 200:
        sys.exit(f"ERROR: Yandex API failed ({r.status_code}): {r.text[:400]}")
    try:
        data = r.json()
    except ValueError:
        sys.exit(f"ERROR: Yandex response not valid JSON: {r.text[:300]}")

    if "errors" in data:
        sys.exit(f"ERROR: Yandex returned errors: {data['errors']}")

    points = data.get("data", {}).get("points", [])
    log(f"Yandex returned {len(points)} data points.")
    return points


# ── Step 2 — Validate & shape ─────────────────────────────────────────────────

def process_points(points: list[dict]) -> pd.DataFrame:
    if not points:
        sys.exit("ERROR: Yandex returned 0 data points — aborting.")

    rows = []
    for p in points:
        dims = p.get("dimensions", {}) or {}
        measures = p.get("measures", [{}])
        m = measures[0] if measures else {}
        date_val = dims.get("date")
        # date can be a list like ["2026-05-28"] or just a string.
        if isinstance(date_val, list):
            date_val = date_val[0] if date_val else None
        rows.append({
            "Date":        date_val,
            "Domain":      dims.get("domain", ""),
            "Impressions": m.get(IMPR_FIELD, 0),
            "Revenue":     m.get(REV_FIELD, 0),
            "eCPM":        m.get(ECPM_FIELD, 0),
        })

    df = pd.DataFrame(rows)
    df = df.dropna(subset=["Date", "Domain"])
    df = df[df["Domain"].astype(str).str.strip() != ""]
    if df.empty:
        sys.exit("ERROR: No valid rows after cleaning — aborting.")

    parsed = pd.to_datetime(df["Date"], errors="coerce")
    bad = parsed.isna().sum()
    if bad:
        log(f"WARNING: {bad} row(s) had unparseable dates and will be dropped.")
    df = df[parsed.notna()].copy()
    df["__date"] = pd.to_datetime(df["Date"], errors="coerce")
    if df.empty:
        sys.exit("ERROR: No valid rows after date parsing — aborting.")

    newest = df["__date"].max()
    age_days = (datetime.now() - newest).days
    if age_days > MAX_ALLOWED_AGE_DAYS:
        sys.exit(
            f"ERROR: Newest date is {newest.date()} ({age_days} days ago). "
            f"Expected within {MAX_ALLOWED_AGE_DAYS} days — aborting."
        )

    # Defensive MTD filter (the API preset is already "thismonth").
    first_of_month = pd.Timestamp(datetime.now().replace(day=1).date())
    _pre_mtd_df = df.copy()
    before = len(df)
    df = df[df["__date"] >= first_of_month]
    if before != len(df):
        log(f"Filtered to MTD ({first_of_month.date()} onward): kept {len(df)}, dropped {before - len(df)}.")

    out = pd.DataFrame({
        "Date":        df["__date"].dt.strftime("%Y-%m-%d"),
        "Domain":      df["Domain"].astype(str).str.strip(),
        "Impressions": pd.to_numeric(df["Impressions"], errors="coerce").fillna(0).astype(int),
        "Revenue":     pd.to_numeric(df["Revenue"],     errors="coerce").fillna(0).round(4),
        "eCPM":        pd.to_numeric(df["eCPM"],        errors="coerce").fillna(0).round(4),
    })

    log(f"Valid: {len(out)} rows, dates {out['Date'].min()} → {out['Date'].max()}")
    return out


# ── Step 3 — Replace sheet contents ───────────────────────────────────────────

def write_sheet(df: pd.DataFrame, creds) -> None:
    if SPREADSHEET_ID == "SPREADSHEET_ID_PLACEHOLDER":
        sys.exit(
            "ERROR: SPREADSHEET_ID not configured. Edit main.py and replace "
            "SPREADSHEET_ID_PLACEHOLDER with the Yandex Google Sheet ID."
        )
    log("Connecting to Google Sheets…")
    service = build("sheets", "v4", credentials=creds, cache_discovery=False)

    rows = [
        [
            str(r["Domain"]).strip(),
            str(r["Date"]).strip(),
            f"${float(pd.to_numeric(r['Revenue'], errors='coerce') or 0):.2f}",
            int(pd.to_numeric(r["Impressions"], errors="coerce") or 0),
            float(pd.to_numeric(r["eCPM"], errors="coerce") or 0),
        ]
        for _, r in df.iterrows()
    ]
    rows = _normalize_and_aggregate(rows)
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
    log("=== Yandex API → Google Sheets daily sync ===")
    cfg   = load_config()
    creds = load_service_account()

    points = fetch_report(cfg["yandex_oauth_token"])
    df     = process_points(points)
    write_sheet(df, creds)
    log("=== Done. ===")


if __name__ == "__main__":
    main()
