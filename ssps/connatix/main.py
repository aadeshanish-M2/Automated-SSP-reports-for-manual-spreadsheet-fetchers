#!/usr/bin/env python3
"""
Connatix (elements.connatix.com) → Google Sheets daily sync.

Logs into the dashboard via Playwright, opens Reporting, refreshes the
pre-saved "MTD" report, waits for it to leave "Pending" status, downloads
the resulting CSV, filters to MTD and replaces the destination sheet.
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
SPREADSHEET_ID = os.environ.get("SHEETS_SPREADSHEET_ID") or "1XecZBeZzgndqIZdTso7SVdCaZGJpC7J2e9-3ozOS7CE"
SHEET_NAME     = "Sheet1"
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]

REPORTS_URL = "https://elements.connatix.com/reports"
REPORT_NAME = "MTD"

# CSV column-name candidates — refine on first run once headers are known.
DATE_COL_CANDIDATES = ["Date", "Day", "date", "day"]
SITE_COL_CANDIDATES = ["Property", "Site", "Domain", "Website", "Site Name", "site", "domain"]
IMPR_COL_CANDIDATES = ["Ad Impressions", "Impressions", "Impression", "Imps", "impressions"]
REV_COL_CANDIDATES  = ["Publisher Total Revenue ($)", "Revenue", "Net Revenue",
                       "Gross Revenue", "Earnings", "revenue"]
ECPM_COL_CANDIDATES = ["CPM ($)", "eCPM", "ECPM", "Effective CPM", "CPM",
                       "Net eCPM", "Gross eCPM"]

MAX_ALLOWED_AGE_DAYS = 5
HEADER = ["Date", "Site", "Impressions", "Revenue", "eCPM"]


# ── Helpers ───────────────────────────────────────────────────────────────────

def log(msg: str) -> None:
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


_REQUIRED_KEYS = ("connatix_username", "connatix_password")


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

def download_connatix_csv(username: str, password: str) -> Path:
    log("Launching browser…")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={"width": 1600, "height": 1000},
            accept_downloads=True,
        )
        page = context.new_page()

        log("Navigating to Connatix Reporting…")
        page.goto(REPORTS_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(4000)

        # ── Login (best-guess selectors, may need adjusting on first real run)
        try:
            page.wait_for_selector(
                'input[type="email"], input[name="email"], input[name="username"], '
                'input[type="text"]',
                timeout=30_000,
            )
            page.fill(
                'input[type="email"], input[name="email"], input[name="username"], '
                'input[type="text"]',
                username,
            )
            page.fill('input[type="password"]', password)
            page.click(
                'button[type="submit"], button:has-text("Sign in"), '
                'button:has-text("Log in"), button:has-text("Login")'
            )
            page.wait_for_url(lambda u: "login" not in u.lower(), timeout=30_000)
            page.wait_for_timeout(5000)
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: Login failed (timeout). Check credentials.\nDetail: {e}")

        log("Logged in. Opening Reporting tab…")
        try:
            page.locator(
                'a:has-text("Reporting"), button:has-text("Reporting"), '
                '[role="menuitem"]:has-text("Reporting")'
            ).first.click()
            page.wait_for_timeout(4000)
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: Could not open Reporting.\nDetail: {e}")

        log(f"Opening '{REPORT_NAME}' report…")
        try:
            page.locator(f':text("{REPORT_NAME}")').first.click()
            page.wait_for_timeout(5000)
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: Could not open the {REPORT_NAME} report.\nDetail: {e}")

        # Connatix disables Refresh when the report was refreshed very recently.
        # If enabled → trigger a refresh, wait for the report list to show
        # Completed, then re-open. If disabled → skip and use the existing
        # data (which is fresh enough).
        refresh_btn = page.locator('button:has-text("Refresh")').first
        try:
            is_disabled = refresh_btn.is_disabled(timeout=5_000)
        except Exception:
            is_disabled = True

        if is_disabled:
            log("Refresh button is disabled (report was refreshed very recently) — "
                "using the existing data.")
        else:
            log("Triggering Refresh…")
            refresh_btn.click()
            page.wait_for_timeout(3000)

            log("Waiting for refresh to complete (can take a few minutes)…")
            completed_re = re.compile(r"(completed|success|done|ready)", re.I)
            pending_re   = re.compile(r"(pending|processing|running|queued)", re.I)
            import time
            start = time.time()
            poll_timeout_s = 900   # 15 minutes
            poll_every_s   = 15
            last_status = ""
            while time.time() - start < poll_timeout_s:
                try:
                    row = page.locator(
                        f'tr:has-text("{REPORT_NAME}"), '
                        f'[role="row"]:has-text("{REPORT_NAME}")'
                    ).first
                    row_text = row.inner_text(timeout=5_000)
                except Exception:
                    row_text = ""
                if completed_re.search(row_text) and not pending_re.search(row_text):
                    log("  → status: Completed.")
                    break
                if row_text != last_status:
                    snippet = " ".join(row_text.split())[:80]
                    log(f"  status: {snippet}")
                    last_status = row_text
                page.wait_for_timeout(poll_every_s * 1000)
            else:
                browser.close()
                sys.exit("ERROR: Report did not reach Completed status in 15 minutes.")

            log(f"Re-opening '{REPORT_NAME}' for download…")
            try:
                page.locator(f':text("{REPORT_NAME}")').first.click()
                page.wait_for_timeout(5000)
            except PlaywrightTimeoutError as e:
                browser.close()
                sys.exit(f"ERROR: Could not re-open the report after refresh.\nDetail: {e}")

        log("Triggering CSV download…")
        try:
            with page.expect_download(timeout=60_000) as dl_info:
                page.locator(
                    'button:has-text("Download"), a:has-text("Download"), '
                    '[aria-label*="download" i]'
                ).first.click()
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
    # Some dashboards prepend banner lines before the header; auto-skip until
    # we find a line that looks like a header row.
    text = csv_path.read_text(errors="replace")
    lines = text.splitlines()
    skip = 0
    for i, line in enumerate(lines):
        low = line.lstrip().lower()
        if low.startswith("date,") or low.startswith('"date"'):
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
        ("Site",        site_col),
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
        "Site":        df[site_col].astype(str).str.strip(),
        "Impressions": pd.to_numeric(df[impr_col], errors="coerce").fillna(0),
        "Revenue":     pd.to_numeric(df[rev_col],  errors="coerce").fillna(0),
        "eCPM":        pd.to_numeric(df[ecpm_col], errors="coerce").fillna(0),
    })
    out = out[~out["Site"].str.lower().isin(["", "nan", "none"])]
    if out.empty:
        sys.exit("ERROR: No valid rows after cleaning — aborting.")

    log(f"Valid: {len(out)} rows, dates {out['Date'].min()} → {out['Date'].max()}")
    return out


# ── Step 3 — Replace sheet contents ───────────────────────────────────────────

def write_sheet(df: pd.DataFrame, creds) -> None:
    log("Connecting to Google Sheets…")
    service = build("sheets", "v4", credentials=creds, cache_discovery=False)

    rows = [
        [r["Date"], r["Site"], r["Impressions"], r["Revenue"], r["eCPM"]]
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
    log("=== Connatix → Google Sheets daily sync ===")
    cfg   = load_config()
    creds = load_service_account()

    csv_path = download_connatix_csv(cfg["connatix_username"], cfg["connatix_password"])
    df       = process_csv(csv_path)
    write_sheet(df, creds)

    try: csv_path.unlink()
    except Exception: pass
    log("=== Done. ===")


if __name__ == "__main__":
    main()
