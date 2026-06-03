#!/usr/bin/env python3
"""
Insticator (publisher.insticator.com) → Google Sheets daily sync.

Logs in via Playwright, opens Report Builder, sets MTD date range +
Day/Domain dimensions + Revenue/Impressions/Avg eCPM metrics, downloads
the CSV, filters to MTD, and replaces the destination sheet.
"""

import json
import os
import re
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# ── Configuration ─────────────────────────────────────────────────────────────

SCRIPT_DIR     = Path(__file__).parent
CONFIG_PATH    = SCRIPT_DIR / "config.json"
SA_PATH        = SCRIPT_DIR / "service_account.json"
SPREADSHEET_ID = os.environ.get("SHEETS_SPREADSHEET_ID") or "1qiDKWM9I-R7DFI2KLD6fwV-HGOIz0b6tbPR20DJgTEQ"
SHEET_NAME     = "Sheet1"
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]

LOGIN_URL = "https://publisher.insticator.com/auth/signin"

# Column-name candidates — refine on first run once headers are known.
DATE_COL_CANDIDATES = ["Day", "Date", "day", "date"]
SITE_COL_CANDIDATES = ["Domain", "Site", "Website", "domain"]
IMPR_COL_CANDIDATES = ["Total Impressions", "Impressions", "impressions"]
REV_COL_CANDIDATES  = ["Net Revenue ($)", "Net Revenue", "Revenue", "Earnings"]
ECPM_COL_CANDIDATES = ["CPM", "Avg CPM", "Average CPM", "Average eCPM", "Avg eCPM", "eCPM"]

MAX_ALLOWED_AGE_DAYS = 5
HEADER = ["Domain", "Date", "Revenue", "Impression", "CPM"]


# ── Helpers ───────────────────────────────────────────────────────────────────

def log(msg: str) -> None:
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


_REQUIRED_KEYS = ("insticator_username", "insticator_password")


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


# ── Step 1 — Drive the dashboard via Playwright ───────────────────────────────

def download_insticator_csv(username: str, password: str) -> Path:
    log("Launching browser…")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={"width": 1600, "height": 1000},
            accept_downloads=True,
        )
        page = context.new_page()

        log("Navigating to Insticator login…")
        page.goto(LOGIN_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(4000)
        try:
            page.wait_for_selector(
                'input[type="email"], input[name="email"], input[name="username"], input[type="text"]',
                timeout=30_000,
            )
            page.fill(
                'input[type="email"], input[name="email"], input[name="username"], input[type="text"]',
                username,
            )
            page.fill('input[type="password"]', password)
            page.click(
                'button[type="submit"], button:has-text("Sign in"), '
                'button:has-text("Log in"), button:has-text("Login")'
            )
            page.wait_for_url(lambda u: "signin" not in u.lower() and "login" not in u.lower(),
                              timeout=30_000)
            page.wait_for_timeout(5000)
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: Login failed (timeout). Check credentials.\nDetail: {e}")

        log("Logged in. Opening Report Builder…")
        try:
            page.get_by_text("Report Builder", exact=True).first.click(timeout=15_000)
            page.wait_for_timeout(6000)
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: Could not open Report Builder.\nDetail: {e}")

        # Insticator's saved reports keep a fixed date range, so we create a
        # fresh report each run with a unique name and download it.
        report_name = f"MTD-auto-{datetime.now().strftime('%Y%m%d-%H%M%S')}"

        log("Clicking Create New Report…")
        try:
            page.locator('button:has-text("Create New Report")').first.click()
            page.wait_for_timeout(6000)
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: Could not click Create New Report.\nDetail: {e}")

        log("Setting custom date range (1st of month → latest enabled day)…")
        # Insticator's Period dropdown has no MTD preset. Leave Period on
        # its default "Custom" and explicitly pick 1st-of-month → latest
        # selectable date (Insticator has a ~2-day reporting lag, so today
        # and often yesterday are greyed out).
        try:
            from datetime import date, timedelta
            today = date.today()
            first = date(today.year, today.month, 1)

            # Open the date-range field by clicking its visible value text.
            page.locator(
                'text=/\\w+ \\d+, \\d+ - \\w+ \\d+, \\d+/'
            ).first.click(timeout=10_000)
            page.wait_for_timeout(1500)

            from_aria = first.strftime("%a %b %d %Y")
            page.click(
                f'.DayPicker-Day[aria-label="{from_aria}"][aria-disabled="false"]',
                timeout=5000,
            )
            page.wait_for_timeout(400)

            # Walk back from today until we find a clickable end-date cell.
            end_clicked = None
            for offset in range(0, 6):
                cand = today - timedelta(days=offset)
                if cand < first:
                    break
                aria = cand.strftime("%a %b %d %Y")
                sel  = f'.DayPicker-Day[aria-label="{aria}"][aria-disabled="false"]'
                if page.locator(sel).count() > 0:
                    page.click(sel, timeout=3000)
                    end_clicked = cand
                    break
            if end_clicked is None:
                raise RuntimeError("no enabled end-date found in current month")
            log(f"  date range: {first.isoformat()} → {end_clicked.isoformat()}")
            page.wait_for_timeout(400)
            page.locator('button:has-text("Apply")').first.click(timeout=5000)
            page.wait_for_timeout(1500)
        except Exception as e:
            log(f"WARNING: date-range picker: {e}")

        log("Checking dimension: Domain/App  (Day is pre-selected)…")
        try:
            page.click('[id="checkbox-list-label-Domain/App"]', timeout=5000)
            page.wait_for_timeout(500)
        except Exception as e:
            log(f"WARNING: Domain/App: {e}")

        log("Checking metrics: Impressions, Avg CPM  (Revenue is pre-selected)…")
        for metric in ("Impressions", "Avg CPM"):
            try:
                page.click(f'[id="checkbox-list-label-{metric}"]', timeout=5000)
                page.wait_for_timeout(500)
            except Exception as e:
                log(f"WARNING: {metric}: {e}")

        log(f"Naming the report '{report_name}' and saving…")
        try:
            page.locator('input[placeholder*="name" i]').first.fill(report_name)
            page.wait_for_timeout(500)
            page.locator('button:has-text("Save & Close")').first.click(timeout=15_000)
            page.wait_for_timeout(8000)
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: Could not save the report.\nDetail: {e}")

        log(f"Finding the '{report_name}' row and clicking its Download icon…")
        try:
            # The name cell's parent IS the row container that holds all columns
            # including the icon-actions cell (edit / download / trash spans).
            name_cell = page.get_by_text(report_name, exact=True).first
            row = name_cell.locator('xpath=..')
            download_icon = row.locator('span.name-download').first
            with page.expect_download(timeout=180_000) as dl_info:
                download_icon.click(timeout=15_000)
            download = dl_info.value
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: CSV download did not start in time.\nDetail: {e}")

        tmp = Path(tempfile.mktemp(suffix=".csv"))
        download.save_as(str(tmp))
        browser.close()

    log(f"CSV downloaded → {tmp}")
    return tmp


# ── Step 2 — Validate & shape ─────────────────────────────────────────────────

def process_csv(csv_path: Path) -> pd.DataFrame:
    log("Reading CSV…")
    text = csv_path.read_text(errors="replace")
    lines = text.splitlines()
    skip = 0
    for i, line in enumerate(lines):
        low = line.lstrip().lower()
        if low.startswith("date,") or low.startswith("day,") or low.startswith('"date"') or low.startswith('"day"'):
            skip = i
            break
    try:
        from io import StringIO
        df = pd.read_csv(StringIO("\n".join(lines[skip:])))
    except Exception as e:
        sys.exit(f"ERROR: Could not parse CSV.\nDetail: {e}")

    if df.empty:
        sys.exit("ERROR: Downloaded CSV is empty — aborting.")

    df.columns = [c.strip() for c in df.columns]

    date_col = first_matching_column(df.columns, DATE_COL_CANDIDATES)
    site_col = first_matching_column(df.columns, SITE_COL_CANDIDATES)
    impr_col = first_matching_column(df.columns, IMPR_COL_CANDIDATES)
    rev_col  = first_matching_column(df.columns, REV_COL_CANDIDATES)
    ecpm_col = first_matching_column(df.columns, ECPM_COL_CANDIDATES)

    missing = [name for name, col in [
        ("Date",        date_col),
        ("Domain",      site_col),
        ("Impressions", impr_col),
        ("Revenue",     rev_col),
        ("eCPM",        ecpm_col),
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
            f"ERROR: Newest date is {newest.date()} ({age_days} days ago). "
            f"Expected within {MAX_ALLOWED_AGE_DAYS} days — aborting."
        )

    first_of_month = pd.Timestamp(datetime.now().replace(day=1).date())
    _pre_mtd_df = df.copy()
    before = len(df)
    df = df[df["__date"] >= first_of_month]
    if before != len(df):
        log(f"Filtered to MTD ({first_of_month.date()} onward): kept {len(df)}, dropped {before - len(df)}.")
    if df.empty:
        log("WARNING: 0 rows match current-month filter — falling back to full report (likely a month-boundary day, MTD data not available yet).")
        df = _pre_mtd_df

    out = pd.DataFrame({
        "Date":        df["__date"].dt.strftime("%Y-%m-%d"),
        "Domain":      df[site_col].astype(str).str.strip(),
        "Impressions": pd.to_numeric(df[impr_col], errors="coerce").fillna(0),
        "Revenue":     pd.to_numeric(df[rev_col],  errors="coerce").fillna(0),
        "eCPM":        pd.to_numeric(df[ecpm_col], errors="coerce").fillna(0),
    })
    out = out[~out["Domain"].str.lower().isin(["", "nan", "none"])]
    if out.empty:
        sys.exit("ERROR: No valid rows after cleaning — aborting.")

    log(f"Valid: {len(out)} rows, dates {out['Date'].min()} → {out['Date'].max()}")
    return out


# ── Step 3 — Replace sheet contents ───────────────────────────────────────────

def write_sheet(df: pd.DataFrame, creds) -> None:
    if SPREADSHEET_ID == "SPREADSHEET_ID_PLACEHOLDER":
        sys.exit(
            "ERROR: SPREADSHEET_ID not configured. Edit main.py and replace "
            "SPREADSHEET_ID_PLACEHOLDER with the Insticator Google Sheet ID."
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
    log("=== Insticator → Google Sheets daily sync ===")
    cfg   = load_config()
    creds = load_service_account()

    csv_path = download_insticator_csv(cfg["insticator_username"], cfg["insticator_password"])
    df       = process_csv(csv_path)
    write_sheet(df, creds)

    try: csv_path.unlink()
    except Exception: pass
    log("=== Done. ===")


if __name__ == "__main__":
    main()
