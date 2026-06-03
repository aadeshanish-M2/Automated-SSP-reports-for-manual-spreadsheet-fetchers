#!/usr/bin/env python3
"""
Adsolut / Playstream Media (manage.playstream.media) → Google Sheets daily sync.

Logs into the dashboard via Playwright, opens Reports → Report Builder, sets
date range to "Month to today" with Dimensions=Domains+Date and
Metrics=Impressions+Revenue+Gross CPM, generates the report, downloads the
CSV, filters to MTD and replaces the destination sheet.
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
SPREADSHEET_ID = os.environ.get("SHEETS_SPREADSHEET_ID") or "1rbAx28rkCrwmnV4WyWMI5hqVZm-iQKP-N3XeM2XdyXo"
SHEET_NAME     = "Sheet1"
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]

LOGIN_URL = "https://manage.playstream.media/"

# CSV column-name candidates — we'll fine-tune on first run once we see actual headers.
DATE_COL_CANDIDATES = ["Date", "Day", "date", "day"]
SITE_COL_CANDIDATES = ["Domain", "Domains", "Website", "Site", "domain"]
IMPR_COL_CANDIDATES = ["Impressions", "Impression", "Imps", "impressions", "impression"]
REV_COL_CANDIDATES  = ["Revenue", "Revenue, $", "Earnings", "revenue"]
ECPM_COL_CANDIDATES = ["Gross CPM", "Gross CPM ($)", "Gross eCPM",
                       "Publisher Cpm", "Publisher CPM",
                       "eCPM", "CPM", "ECPM"]

MAX_ALLOWED_AGE_DAYS = 5
HEADER = ["Domain", "Date", "Revenue", "Impression", "CPM"]


# ── Helpers ───────────────────────────────────────────────────────────────────

def log(msg: str) -> None:
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


_REQUIRED_KEYS = ("adsolut_username", "adsolut_password")


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

def download_adsolut_csv(username: str, password: str) -> Path:
    log("Launching browser…")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={"width": 1600, "height": 1000},
            accept_downloads=True,
        )
        page = context.new_page()

        log("Navigating to Adsolut login…")
        page.goto(LOGIN_URL, wait_until="domcontentloaded")
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
            page.wait_for_url(lambda u: "login" not in u.lower(), timeout=30_000)
            page.wait_for_timeout(4000)
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: Login failed (timeout). Check credentials.\nDetail: {e}")

        log("Logged in. Opening Reports → Report Builder…")
        try:
            page.locator('a:has-text("Reports"), button:has-text("Reports")').first.click()
            page.wait_for_timeout(1500)
            page.locator(':text("Report builder")').first.click()
            page.wait_for_timeout(6000)
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: Could not open Report Builder.\nDetail: {e}")

        log("Setting date range = Month To Today…")
        try:
            # The date-range area shows a date input like "28/05/26" — click it
            # to open the preset list.
            page.locator('text=/\\d{2}\\/\\d{2}\\/\\d{2}/').first.click()
            page.wait_for_timeout(1500)
            page.locator(':text("Month To Today")').first.click()
            page.wait_for_timeout(1000)
        except Exception as e:
            log(f"WARNING: date-range selector failed: {e}")

        def pick_via_plus(section_index: int, option_label: str, search_term: str = ""):
            """
            Open the "+" picker in the given section (0 = Dimensions, 1 = Metrics)
            and click the option with an EXACT-match label. Optionally filter the
            list with a search term first.
            """
            try:
                page.locator('button:has-text("add")').nth(section_index).click()
                page.wait_for_timeout(1200)
                if search_term:
                    try:
                        page.locator('.cdk-overlay-container input').first.fill(search_term)
                        page.wait_for_timeout(600)
                    except Exception:
                        pass
                page.locator('.cdk-overlay-container').get_by_text(
                    option_label, exact=True
                ).first.click(timeout=5000)
                page.wait_for_timeout(600)
                # Close picker reliably so the next picker open works cleanly.
                page.keyboard.press("Escape")
                try:
                    page.wait_for_function(
                        "() => document.querySelectorAll('.cdk-overlay-backdrop-showing').length === 0",
                        timeout=8_000,
                    )
                except Exception:
                    pass
                page.wait_for_timeout(400)
                log(f"  → added '{option_label}'")
            except Exception as e:
                log(f"WARNING: could not add '{option_label}': {e}")

        log("Choosing Dimensions: Date, Domain…")
        pick_via_plus(0, "Date")
        pick_via_plus(0, "Domain", search_term="domain")

        log("Choosing Metrics: Impression, Revenue, Gross CPM ($)…")
        pick_via_plus(1, "Impression",   search_term="impression")
        pick_via_plus(1, "Revenue",      search_term="revenue")
        pick_via_plus(1, "Gross CPM ($)", search_term="gross cpm")

        log("Clicking GENERATE REPORT…")
        try:
            page.locator('button:has-text("GENERATE REPORT")').first.click()
            # Download icon (file_download material icon) is initially disabled.
            # Wait until it becomes enabled.
            page.wait_for_function(
                """() => {
                    const btns = Array.from(document.querySelectorAll('button'));
                    const dl = btns.find(b => b.innerText.trim() === 'file_download');
                    return dl && !dl.disabled;
                }""",
                timeout=180_000,
            )
            page.wait_for_timeout(1500)
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: Report didn't generate (download stayed disabled).\nDetail: {e}")

        log("Triggering CSV download…")
        try:
            with page.expect_download(timeout=60_000) as dl_info:
                page.locator('button:has-text("file_download")').first.click()
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
    # Adsolut prepends 5-7 banner lines ("Report Name:", "Report Id:", …) before
    # the actual header row. Find the first line starting with "Date," and skip
    # everything above it.
    text = csv_path.read_text()
    lines = text.splitlines()
    skip = 0
    for i, line in enumerate(lines):
        if line.lstrip().lower().startswith("date,"):
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
        ("Gross CPM",   ecpm_col),
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
            f"ERROR: Newest date in CSV is {newest.date()} ({age_days} days ago). "
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
        "Gross CPM":   pd.to_numeric(df[ecpm_col], errors="coerce").fillna(0),
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
            "SPREADSHEET_ID_PLACEHOLDER with the Adsolut Google Sheet ID."
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
    log("=== Adsolut → Google Sheets daily sync ===")
    cfg   = load_config()
    creds = load_service_account()

    csv_path = download_adsolut_csv(cfg["adsolut_username"], cfg["adsolut_password"])
    df       = process_csv(csv_path)
    write_sheet(df, creds)

    try: csv_path.unlink()
    except Exception: pass
    log("=== Done. ===")


if __name__ == "__main__":
    main()
