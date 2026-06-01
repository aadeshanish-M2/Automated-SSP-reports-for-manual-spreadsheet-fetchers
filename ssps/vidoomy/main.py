#!/usr/bin/env python3
"""
Vidoomy (login.vidoomy.com) → Google Sheets daily sync.

Logs in via Playwright, opens Stats Pro Reports, switches the preset to
Daily Report, sets the date range to Current Month, ensures Site dimension
and Impressions/Revenue/CPM metrics, then triggers "Run to CSV".
"""

import json
import os
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
SPREADSHEET_ID = os.environ.get("SHEETS_SPREADSHEET_ID") or "1ijiHMiaf1l4ktHMQ-Z26tiWg99bpmn-X2_fv5xp2mjE"
SHEET_NAME     = "Sheet1"
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]

LOGIN_URL = "https://login.vidoomy.com/"

DATE_COL_CANDIDATES = ["Date", "Day", "date", "day"]
SITE_COL_CANDIDATES = ["Site Url", "Site URL", "Site", "Domain", "Website", "site"]
IMPR_COL_CANDIDATES = ["Impressions", "Impression", "impressions"]
REV_COL_CANDIDATES  = ["Revenue", "Net Revenue", "Earnings", "revenue"]
ECPM_COL_CANDIDATES = ["CPM", "eCPM", "ECPM", "Avg CPM"]

MAX_ALLOWED_AGE_DAYS = 5
HEADER = ["Date", "Site", "Impressions", "Revenue", "CPM"]


# ── Helpers ───────────────────────────────────────────────────────────────────

def log(msg: str) -> None:
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


_REQUIRED_KEYS = ("vidoomy_username", "vidoomy_password")


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


def clean_european_number(v) -> float:
    """
    Vidoomy uses European formatting: '.' as thousands sep, ',' as decimal,
    plus currency / percent prefixes. Convert to float.
        '$5,95'   -> 5.95
        '2.260'   -> 2260.0
        '20,90%'  -> 20.90
        '10.813'  -> 10813.0
    """
    if v is None:
        return 0.0
    s = str(v).strip().replace("$", "").replace("€", "").replace("%", "").replace(" ", "")
    if not s or s.lower() == "nan":
        return 0.0
    # If both . and , present, '.' must be thousands and ',' must be decimal.
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        # Only comma — treat as decimal separator.
        s = s.replace(",", ".")
    else:
        # Only periods — assume thousands grouping (Vidoomy never shows trailing
        # decimals without a comma). Drop the periods.
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


# ── Step 1 — Drive the dashboard via Playwright ───────────────────────────────

def download_vidoomy_csv(username: str, password: str) -> Path:
    log("Launching browser…")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={"width": 1600, "height": 1000},
            accept_downloads=True,
        )
        page = context.new_page()

        log("Navigating to Vidoomy login…")
        page.goto(LOGIN_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(4000)
        try:
            page.wait_for_selector(
                'input[type="email"], input[name="email"], input[name="username"], input[type="text"]',
                timeout=30_000,
            )
            page.fill('input[name="_username"], input[type="text"]', username)
            page.fill('input[name="_password"], input[type="password"]', password)
            page.locator(
                'button[type="submit"], button:has-text("Sign in")'
            ).first.click(timeout=10_000)
            # Wait until the password input is gone (login form replaced by dashboard).
            page.wait_for_selector('input[type="password"]', state="detached", timeout=30_000)
            page.wait_for_timeout(5000)
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: Login failed (timeout). Check credentials.\nDetail: {e}")

        log("Logged in. Opening Stats Pro Reports…")
        try:
            page.locator(
                'a:has-text("Stats Pro Reports"), button:has-text("Stats Pro Reports"), '
                ':text("Stats Pro Reports")'
            ).first.click()
            page.wait_for_timeout(6000)
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: Could not open Stats Pro Reports.\nDetail: {e}")

        log("Switching preset: Overall Report → Daily Report…")
        try:
            # Click the "Overall Report" trigger to open the preset list, then
            # pick the "Daily Report" anchor.
            page.locator('button:has-text("Overall Report")').first.click(timeout=8_000)
            page.wait_for_timeout(1200)
            page.locator('a:has-text("Daily Report")').first.click(timeout=5000)
            page.wait_for_timeout(1500)
        except Exception as e:
            log(f"WARNING: preset switch failed: {e}")

        log("Setting date range preset = Current month…")
        try:
            page.locator(':text("Date")').first.click(timeout=8_000)
            page.wait_for_timeout(1200)
            # Vidoomy's actual label is "Current month" (lowercase m).
            page.get_by_text("Current month", exact=True).first.click(timeout=5000)
            page.wait_for_timeout(1500)
        except Exception as e:
            log(f"WARNING: date-range selector failed: {e}")

        log("Adding CPM metric (Impressions and Revenue are pre-selected)…")
        try:
            # Open the Metric dropdown to make hidden chips visible, then click CPM.
            page.locator(':text("Metric")').first.click(timeout=8_000)
            page.wait_for_timeout(1200)
            page.locator('a.metric-item[rel="cpm"]').first.click(timeout=5000)
            page.wait_for_timeout(400)
            # Close the dropdown so it doesn't intercept the Run click.
            page.keyboard.press("Escape")
            page.wait_for_timeout(500)
        except Exception as e:
            log(f"WARNING: CPM metric click failed: {e}")

        log("Clicking 'Run to CSV' to trigger download…")
        try:
            with page.expect_download(timeout=120_000) as dl_info:
                page.locator(
                    'button:has-text("Run to CSV"), a:has-text("Run to CSV"), '
                    'button:has-text("Run"):has-text("CSV")'
                ).first.click(timeout=10_000)
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
        if low.startswith("date,") or low.startswith("day,") or low.startswith('"date"'):
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
        ("CPM",         ecpm_col),
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
        "Impressions": df[impr_col].apply(clean_european_number),
        "Revenue":     df[rev_col].apply(clean_european_number),
        "CPM":         df[ecpm_col].apply(clean_european_number),
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
        [r["Date"], r["Site"], r["Impressions"], r["Revenue"], r["CPM"]]
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
    log("=== Vidoomy → Google Sheets daily sync ===")
    cfg   = load_config()
    creds = load_service_account()

    csv_path = download_vidoomy_csv(cfg["vidoomy_username"], cfg["vidoomy_password"])
    df       = process_csv(csv_path)
    write_sheet(df, creds)

    try: csv_path.unlink()
    except Exception: pass
    log("=== Done. ===")


if __name__ == "__main__":
    main()
