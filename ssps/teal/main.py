#!/usr/bin/env python3
"""
Teal dashboard → Google Sheets daily revenue sync.

Logs into dashboard.teal.works, exports the last-28-days CSV from the Data tab,
computes CPM, and upserts rows into the target Google Sheet (keyed on Domain+Date).
Existing rows outside the export window are left untouched. Newest dates on top.
"""

import json
import os
import sys
import tempfile
import time
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# ── Configuration ─────────────────────────────────────────────────────────────

SCRIPT_DIR      = Path(__file__).parent
CONFIG_PATH     = SCRIPT_DIR / "config.json"
SA_PATH         = SCRIPT_DIR / "service_account.json"
SPREADSHEET_ID  = os.environ.get("SHEETS_SPREADSHEET_ID") or "1IdWwhdzOhzkqoeTV07xOZI6piRvkM_aJOh0ZoejsA2c"
SHEET_NAME      = "Sheet1"          # change if your tab has a different name
SCOPES          = ["https://www.googleapis.com/auth/spreadsheets"]

EXPECTED_COLUMNS = {"Domain", "Date", "Impressions", "Revenue"}
MAX_ALLOWED_AGE_DAYS = 5            # abort if newest row in CSV is older than this


# ── Helpers ───────────────────────────────────────────────────────────────────

def log(msg: str) -> None:
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


_REQUIRED_KEYS = ("teal_username", "teal_password")


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


# ── Step 1 — Download CSV from Teal ──────────────────────────────────────────

def download_teal_csv(username: str, password: str) -> Path:
    log("Launching browser…")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={"width": 1600, "height": 1000},
            accept_downloads=True,
        )
        page = context.new_page()

        # ── Login (Auth0-hosted page at auth.teal.works) ──
        log("Navigating to Teal dashboard login…")
        page.goto("https://dashboard.teal.works/", wait_until="domcontentloaded")
        try:
            page.wait_for_selector('input[name="username"]', timeout=30_000)
            page.fill('input[name="username"]', username)
            page.fill('input[name="password"]', password)
            page.click('button[type="submit"]')
            page.wait_for_url("https://dashboard.teal.works/**", timeout=30_000)
            page.wait_for_timeout(4000)
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: Login failed (timeout). Check credentials.\nDetail: {e}")

        log("Logged in. Navigating to Data tab…")
        try:
            page.locator('a:has-text("Data")').first.click()
            # Looker Studio embed keeps long-poll connections open, so networkidle
            # never settles — wait a fixed window for the iframe to render.
            page.wait_for_timeout(10_000)
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: Could not find or click the Data tab.\nDetail: {e}")

        # ── Switch into the Looker Studio iframe ──
        looker = next(
            (f for f in page.frames if "datastudio.google.com/embed" in f.url),
            None,
        )
        if looker is None:
            browser.close()
            sys.exit("ERROR: Looker Studio iframe not found on the Data tab.")

        log("Opening chart menu…")
        try:
            looker.locator('button[aria-label="Show chart menu"]').click(timeout=15_000)
            page.wait_for_timeout(1000)
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: Could not open the chart menu.\nDetail: {e}")

        try:
            looker.locator('[role="menuitem"]:has-text("Export chart")').click(timeout=5_000)
            page.wait_for_timeout(800)
            looker.locator('[role="menuitem"]:has-text("Export data")').click(timeout=5_000)
            page.wait_for_timeout(2500)
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: Could not find Export chart / Export data menu items.\nDetail: {e}")

        log("Export dialog open. Starting download…")
        try:
            with page.expect_download(timeout=60_000) as dl_info:
                looker.locator('button:has-text("Export")').click(timeout=10_000)
            download = dl_info.value
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: Download did not start within 60 s.\nDetail: {e}")

        tmp = Path(tempfile.mktemp(suffix=".csv"))
        download.save_as(str(tmp))
        browser.close()

    log(f"CSV downloaded → {tmp}")
    return tmp


# ── Step 2 — Validate & transform CSV ────────────────────────────────────────

def process_csv(csv_path: Path) -> pd.DataFrame:
    log("Reading CSV…")
    try:
        df = pd.read_csv(csv_path)
    except Exception as e:
        sys.exit(f"ERROR: Could not parse CSV.\nDetail: {e}")

    if df.empty:
        sys.exit("ERROR: Downloaded CSV is empty — aborting to protect the sheet.")

    # Normalise column names (strip whitespace, title-case)
    df.columns = [c.strip() for c in df.columns]

    missing = EXPECTED_COLUMNS - set(df.columns)
    if missing:
        sys.exit(
            f"ERROR: CSV is missing expected columns: {missing}\n"
            f"       Found columns: {list(df.columns)}\n"
            "       Aborting to protect the sheet."
        )

    # Parse dates — Looker Studio normally exports them as "14 May 2026", but
    # on the cloud the locale / output can shift. Try the pinned format first,
    # then fall back to pandas' auto-inference if too many rows fail.
    sample_dates = df["Date"].astype(str).head(5).tolist()
    log(f"Date column sample (first 5 raw values): {sample_dates}")

    parsed = pd.to_datetime(df["Date"], format="%d %b %Y", errors="coerce")
    if parsed.isna().sum() > len(df) * 0.5:
        log("Pinned format parsed <50% of rows; falling back to auto-inference.")
        parsed = pd.to_datetime(df["Date"], errors="coerce", dayfirst=True)

    df["Date"] = parsed
    bad_dates = df["Date"].isna().sum()
    if bad_dates:
        log(f"WARNING: {bad_dates} of {len(df)} row(s) had unparseable dates and will be dropped.")
        df = df.dropna(subset=["Date"])

    if df.empty:
        sys.exit(
            "ERROR: No valid rows remain after date parsing — aborting.\n"
            f"       Raw date samples were: {sample_dates}\n"
            "       The format may have shifted (locale / cloud rendering)."
        )

    newest = df["Date"].max()
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
    df = df[df["Date"] >= first_of_month]
    log(f"Filtered to MTD ({first_of_month.date()} onward): kept {len(df)}, dropped {before - len(df)}.")
    if df.empty:
        log("WARNING: 0 rows match current-month filter — falling back to full report (likely a month-boundary day, MTD data not available yet).")
        df = _pre_mtd_df

    # Coerce numeric columns
    df["Impressions"] = pd.to_numeric(df["Impressions"], errors="coerce").fillna(0)
    df["Revenue"]     = pd.to_numeric(df["Revenue"],     errors="coerce").fillna(0)

    # Compute CPM
    df["CPM"] = df.apply(
        lambda r: round(r["Revenue"] * 1000 / r["Impressions"], 4)
        if r["Impressions"] > 0 else 0.0,
        axis=1,
    )

    # Keep only the columns we care about, in order
    df = df[["Domain", "Date", "Impressions", "Revenue", "CPM"]]
    df["Date"] = df["Date"].dt.strftime("%Y-%m-%d")

    log(f"CSV valid: {len(df)} rows, dates {df['Date'].min()} → {df['Date'].max()}")
    return df


# ── Step 3 — Upsert into Google Sheets ───────────────────────────────────────

HEADER = ["Domain", "Date", "Revenue", "Impression", "CPM"]


def get_sheet_service(creds):
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def read_sheet(service, spreadsheet_id: str, sheet_name: str) -> list[list]:
    result = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=sheet_name)
        .execute()
    )
    return result.get("values", [])


def write_sheet(service, spreadsheet_id: str, sheet_name: str, rows: list[list]) -> None:
    body = {"values": rows}
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!A1",
        valueInputOption="USER_ENTERED",
        body=body,
    ).execute()


def clear_sheet(service, spreadsheet_id: str, sheet_name: str) -> None:
    service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=sheet_name,
        body={},
    ).execute()


def upsert_to_sheet(df: pd.DataFrame, creds) -> None:
    """MTD-only mode: replace the sheet contents with the freshly-pulled MTD rows."""
    log("Connecting to Google Sheets…")
    service = get_sheet_service(creds)

    # Build rows from the (already MTD-filtered) df, sorted newest first then by Domain.
    rows = [
        [
            str(r["Domain"]).strip(),
            str(r["Date"]).strip(),
            f"${float(pd.to_numeric(r['Revenue'], errors='coerce') or 0):.2f}",
            int(pd.to_numeric(r["Impressions"], errors="coerce") or 0),
            float(pd.to_numeric(r["CPM"], errors="coerce") or 0),
        ]
        for _, r in df.iterrows()
    ]
    rows.sort(key=lambda r: (r[1], r[0]))     # Date asc, Domain asc
    rows.sort(key=lambda r: r[1], reverse=True)  # then Date desc (stable)

    final_rows = [HEADER] + rows
    log(f"Writing {len(rows)} MTD rows to sheet (replacing existing)…")
    clear_sheet(service, SPREADSHEET_ID, SHEET_NAME)
    write_sheet(service, SPREADSHEET_ID, SHEET_NAME, final_rows)
    log("Sheet updated successfully.")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    log("=== Teal → Google Sheets daily sync ===")

    cfg   = load_config()
    creds = load_service_account()

    csv_path = download_teal_csv(cfg["teal_username"], cfg["teal_password"])
    df       = process_csv(csv_path)
    upsert_to_sheet(df, creds)

    # Clean up temp file
    try:
        csv_path.unlink()
    except Exception:
        pass

    log("=== Done. ===")


if __name__ == "__main__":
    main()
