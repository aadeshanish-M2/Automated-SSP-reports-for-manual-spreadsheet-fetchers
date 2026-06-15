#!/usr/bin/env python3
"""
Ogury (publishers.ogury.co) → Google Sheets daily sync.

Logs in via Playwright, opens Exclusive Demand → Report, sets the date
picker to the "This month" preset, exports the per-day-per-asset CSV, and
syncs to the destination sheet in the uniform format (Domain | Date |
Revenue | Impression | CPM). Previous-month history preserved.
"""

import json
import os
import sys
import tempfile
from datetime import date, datetime
from io import StringIO
from pathlib import Path

import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# ── Configuration ─────────────────────────────────────────────────────────────

SCRIPT_DIR     = Path(__file__).parent
CONFIG_PATH    = SCRIPT_DIR / "config.json"
SA_PATH        = SCRIPT_DIR / "service_account.json"
SPREADSHEET_ID = os.environ.get("SHEETS_SPREADSHEET_ID") or "1GxiIJ-kM8rKqic96F1cvxsfIMk7WzifoCC-nzL6JT34"
SHEET_NAME     = "Sheet1"
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]

LOGIN_URL  = "https://publishers.ogury.co/identity/login?redirect=%2F"
REPORT_URL = "https://publishers.ogury.co/exclusive-demand/reports"

# CSV headers (confirmed):
# Date,Asset Key,Asset,Asset Bundle,Asset Type Key,Asset Type,Revenues,Impressions,Requests,Ecpm
DATE_COL_CANDIDATES = ["Date", "date"]
SITE_COL_CANDIDATES = ["Asset", "Asset Bundle", "asset"]
IMPR_COL_CANDIDATES = ["Impressions", "impressions"]
REV_COL_CANDIDATES  = ["Revenues", "Revenue", "revenues"]
ECPM_COL_CANDIDATES = ["Ecpm", "eCPM", "ECPM"]

MAX_ALLOWED_AGE_DAYS = 5
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
        if s.startswith("www."):
            s = s[4:]
        return s

    bucket = {}
    for r in rows:
        domain = root(r[0])
        date_  = str(r[1]).strip()
        try:
            rev = float(str(r[2]).lstrip("$").replace(",", ""))
        except Exception:
            rev = 0.0
        try:
            imp = int(float(str(r[3])))
        except Exception:
            imp = 0
        k = (domain, date_)
        if k in bucket:
            bucket[k]["rev"] += rev
            bucket[k]["imp"] += imp
        else:
            bucket[k] = {"rev": rev, "imp": imp}

    new_rows = []
    for (domain, date_), v in bucket.items():
        cpm = round(v["rev"] / v["imp"] * 1000, 4) if v["imp"] > 0 else 0.0
        new_rows.append([domain, date_, f"${v['rev']:.2f}", v["imp"], cpm])
    return new_rows


# ── Helpers ───────────────────────────────────────────────────────────────────

def log(msg: str) -> None:
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


_REQUIRED_KEYS = ("ogury_username", "ogury_password")


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


def first_matching_column(df_cols, candidates):
    norm = {c.strip().lower(): c for c in df_cols}
    for cand in candidates:
        if cand.strip().lower() in norm:
            return norm[cand.strip().lower()]
    return None


# ── Step 1 — Drive the dashboard via Playwright ───────────────────────────────

def download_ogury_csv(username: str, password: str) -> Path:
    log("Launching browser…")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={"width": 1700, "height": 1100},
            accept_downloads=True,
        )
        page = context.new_page()

        log("Navigating to Ogury login…")
        page.goto(LOGIN_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(8000)
        try:
            page.fill('input[name="email"]', username)
            page.fill('input[name="password"]', password)
            page.click('button:has-text("SIGN IN")')
            page.wait_for_timeout(10_000)
            if "login" in page.url.lower():
                browser.close()
                sys.exit("ERROR: Still on login page after submit — check credentials.")
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: Login failed (timeout).\nDetail: {e}")

        log("Opening Exclusive Demand → Report…")
        page.goto(REPORT_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(10_000)

        log("Setting date range = This month…")
        try:
            # Wait for the report to actually render by watching for the
            # date-range button (shows e.g. "7 Jun — 14 Jun 2026"). The icon-font
            # <i> glyph inside it doesn't reliably register as "visible", so we
            # anchor on the button's date text instead.
            date_btn = page.locator(
                "button:has-text('Jun'), button:has-text('Jul'), "
                "button:has-text('2026'), button:has-text('2027')"
            ).first
            date_btn.wait_for(state="visible", timeout=60_000)
            page.wait_for_timeout(1500)
            date_btn.click(timeout=15_000)
            page.wait_for_timeout(2000)
            page.locator('a.btn-link:has-text("This month")').click(timeout=15_000)
            page.wait_for_timeout(1000)
            page.locator('span:has-text("Apply"), button:has-text("Apply")').first.click(timeout=10_000)
            page.wait_for_timeout(8000)
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: Could not set This month date range.\nDetail: {e}")

        log("Exporting CSV…")
        try:
            page.locator('button:has(i.oi-x24-export)').first.click(timeout=15_000)
            page.wait_for_timeout(1500)
            page.get_by_text("CSV (Comma Separated Values)").click(timeout=10_000)
            page.wait_for_timeout(500)
            with page.expect_download(timeout=60_000) as dl_info:
                page.get_by_text("DOWNLOAD EXPORT", exact=False).click(timeout=10_000)
            download = dl_info.value
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: CSV export did not start in time.\nDetail: {e}")

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
        if line.lstrip().lower().startswith(("date,", '"date"')):
            skip = i
            break
    try:
        df = pd.read_csv(StringIO("\n".join(lines[skip:])))
    except Exception as e:
        sys.exit(f"ERROR: Could not parse CSV.\nDetail: {e}")

    if df.empty:
        sys.exit("ERROR: Downloaded CSV is empty — aborting.")

    df.columns = [c.strip() for c in df.columns]
    log(f"CSV columns: {list(df.columns)}")

    date_col = first_matching_column(df.columns, DATE_COL_CANDIDATES)
    site_col = first_matching_column(df.columns, SITE_COL_CANDIDATES)
    impr_col = first_matching_column(df.columns, IMPR_COL_CANDIDATES)
    rev_col  = first_matching_column(df.columns, REV_COL_CANDIDATES)
    ecpm_col = first_matching_column(df.columns, ECPM_COL_CANDIDATES)

    missing = [name for name, col in [
        ("date",        date_col),
        ("site",        site_col),
        ("impressions", impr_col),
        ("revenue",     rev_col),
        ("cpm",         ecpm_col),
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

    # MTD filter with month-boundary fallback
    first_of_month = pd.Timestamp(datetime.now().replace(day=1).date())
    _pre_mtd_df = df.copy()
    df = df[df["__date"] >= first_of_month]
    if df.empty:
        log("WARNING: 0 rows match current-month filter — falling back to full report "
            "(likely a month-boundary day, MTD data not available yet).")
        df = _pre_mtd_df

    out = pd.DataFrame({
        "Date":        df["__date"].dt.strftime("%Y-%m-%d"),
        "Site":        df[site_col].astype(str).str.strip(),
        "Impressions": pd.to_numeric(df[impr_col], errors="coerce").fillna(0),
        "Revenue":     pd.to_numeric(df[rev_col],  errors="coerce").fillna(0),
        "CPM":         pd.to_numeric(df[ecpm_col], errors="coerce").fillna(0),
    })
    out = out[~out["Site"].str.lower().isin(["", "nan", "none"])]
    if out.empty:
        sys.exit("ERROR: No valid rows after cleaning — aborting.")

    log(f"Valid: {len(out)} rows, dates {out['Date'].min()} → {out['Date'].max()}")
    return out


# ── Step 3 — Write to sheet (preserve previous-month history) ─────────────────

def write_sheet(df: pd.DataFrame, creds) -> None:
    log("Connecting to Google Sheets…")
    service = build("sheets", "v4", credentials=creds, cache_discovery=False)

    # Auto-detect the first tab name (sheet might not be called "Sheet1").
    try:
        _meta = service.spreadsheets().get(
            spreadsheetId=SPREADSHEET_ID, fields="sheets.properties.title"
        ).execute()
        _tab = _meta["sheets"][0]["properties"]["title"]
    except Exception:
        _tab = SHEET_NAME
    log(f"Using sheet tab: {_tab!r}")

    rows = [
        [
            str(r["Site"]).strip(),
            str(r["Date"]).strip(),
            f"${float(pd.to_numeric(r['Revenue'], errors='coerce') or 0):.2f}",
            int(pd.to_numeric(r["Impressions"], errors="coerce") or 0),
            float(pd.to_numeric(r["CPM"], errors="coerce") or 0),
        ]
        for _, r in df.iterrows()
    ]
    rows = _normalize_and_aggregate(rows)
    rows.sort(key=lambda r: (r[0], r[1]))
    rows.sort(key=lambda r: r[1], reverse=True)

    # Preserve previous-month rows; only the current month gets refreshed.
    existing = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=_tab,
    ).execute().get("values", [])
    current_month = datetime.now().strftime("%Y-%m")
    preserved = []
    for _r in existing:
        if not _r or _r == HEADER:
            continue
        if len(_r) >= 2 and str(_r[1]).strip().startswith(current_month):
            continue  # current-month row → replaced by fresh data
        preserved.append(_r)

    merged = rows + preserved
    merged.sort(key=lambda r: (str(r[1]) if len(r) > 1 else "", str(r[0]) if len(r) > 0 else ""))
    merged.sort(key=lambda r: str(r[1]) if len(r) > 1 else "", reverse=True)

    def _stringify(cell):
        if cell is None:
            return ""
        if isinstance(cell, (int, float)):
            if isinstance(cell, float) and (cell != cell):  # NaN
                return ""
            return cell
        return str(cell)
    merged = [[_stringify(c) for c in r] for r in merged]

    final_rows = [HEADER] + merged
    log(f"Writing {len(merged)} rows = {len(rows)} new MTD + {len(preserved)} preserved (history)…")
    service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID, range=_tab, body={},
    ).execute()
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{_tab}!A1",
        valueInputOption="USER_ENTERED",
        body={"values": final_rows},
    ).execute()
    log("Sheet updated successfully.")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    log("=== Ogury → Google Sheets daily sync ===")
    cfg   = load_config()
    creds = load_service_account()

    csv_path = download_ogury_csv(cfg["ogury_username"], cfg["ogury_password"])
    df       = process_csv(csv_path)
    write_sheet(df, creds)

    try: csv_path.unlink()
    except Exception: pass
    log("=== Done. ===")


if __name__ == "__main__":
    main()
