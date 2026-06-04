#!/usr/bin/env python3
"""
Kuantyx (panel.kuantyx.com) → Google Sheets daily revenue sync.

Logs into the dashboard via Playwright, opens Statistics, sets Group by = Day
and Group by (extra) = Website with date range = This Month, generates the
report, downloads the CSV, filters to MTD and replaces the destination sheet.
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
SPREADSHEET_ID = os.environ.get("SHEETS_SPREADSHEET_ID") or "12fxBEEDMYqeQQfhxAtK9tvFufwexUQ8J986xrMzWdco"
SHEET_NAME     = "Sheet1"
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]

LOGIN_URL = "https://panel.kuantyx.com/publish"

# CSV column-name candidates — adjust on first run once we see the actual headers.
DATE_COL_CANDIDATES = ["Date", "Day", "date", "day"]
SITE_COL_CANDIDATES = ["Website", "Site", "Domain", "website", "site"]
IMPR_COL_CANDIDATES = ["Impressions", "Imps", "impressions"]
REV_COL_CANDIDATES  = ["Revenue, $", "Revenue", "Earnings", "revenue"]
ECPM_COL_CANDIDATES = ["eCPM, $", "eCPM", "ECPM", "Effective CPM", "ecpm"]

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


_REQUIRED_KEYS = ("kuantyx_username", "kuantyx_password")


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

def download_kuantyx_csv(username: str, password: str) -> Path:
    log("Launching browser…")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={"width": 1600, "height": 1000},
            accept_downloads=True,
        )
        page = context.new_page()

        log("Navigating to Kuantyx login…")
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
                'button[type="submit"], button:has-text("Login"), '
                'button:has-text("Log in"), button:has-text("Sign in")'
            )
            page.wait_for_url(lambda u: "login" not in u.lower(), timeout=30_000)
            page.wait_for_timeout(4000)
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: Login failed (timeout). Check credentials.\nDetail: {e}")

        log("Logged in. Opening Statistics…")
        try:
            page.locator(
                'a:has-text("Statistics"), button:has-text("Statistics"), '
                '[role="menuitem"]:has-text("Statistics")'
            ).first.click()
            page.wait_for_timeout(6000)
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: Could not open Statistics.\nDetail: {e}")

        log("Setting Group by = Day, Group by (extra) = Website, range = This Month…")
        try:
            # Primary group-by → "Day"
            page.locator(
                'label:has-text("Group by") + * select, '
                'label:has-text("Group by") ~ select, '
                'select:near(:text("Group by"))'
            ).first.select_option(label="Day")
            page.wait_for_timeout(500)
        except Exception as e:
            log(f"WARNING: primary group-by select failed: {e}")

        try:
            # Secondary "Group by (extra)" → "Website"
            page.locator(
                'label:has-text("Group by (extra)") + * select, '
                'label:has-text("Group by (extra)") ~ select'
            ).first.select_option(label="Website")
            page.wait_for_timeout(500)
        except Exception as e:
            log(f"WARNING: extra group-by select failed: {e}")

        try:
            # Date range → "This Month"
            page.locator(
                'button:has-text("This Month"), '
                'label:has-text("Date") ~ * :text("This Month"), '
                'select:near(:text("Date range"))'
            ).first.click()
            page.wait_for_timeout(500)
        except Exception as e:
            log(f"WARNING: date-range selector failed: {e}")

        log("Clicking Generate…")
        try:
            page.locator('button:has-text("Generate")').first.click()
            # Wait for the Export button to become available (or visible).
            page.wait_for_selector(
                'button:has-text("Export"), a:has-text("Export")',
                timeout=120_000,
            )
            page.wait_for_timeout(2000)
        except PlaywrightTimeoutError as e:
            browser.close()
            sys.exit(f"ERROR: Report didn't generate in time.\nDetail: {e}")

        log("Triggering Export → CSV…")
        try:
            page.locator('button:has-text("Export"), a:has-text("Export")').first.click()
            page.wait_for_timeout(800)
            with page.expect_download(timeout=60_000) as dl_info:
                page.locator(':text("CSV")').first.click()
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
    try:
        df = pd.read_csv(csv_path)
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
        ("Website",     site_col),
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
    # Kuantyx exports dates as DD.MM.YYYY. Pin the format so pandas doesn't
    # auto-detect them as MM.DD and silently corrupt half the rows.
    parsed = pd.to_datetime(df[date_col], format="%d.%m.%Y", errors="coerce")
    bad = parsed.isna().sum()
    if bad:
        log(f"WARNING: {bad} row(s) had unparseable dates and will be dropped.")
    df = df[parsed.notna()].copy()
    df["__date"] = pd.to_datetime(df[date_col], format="%d.%m.%Y", errors="coerce")
    if df.empty:
        sys.exit("ERROR: No valid rows after date parsing — aborting.")

    newest = df["__date"].max()
    age_days = (datetime.now() - newest).days
    if age_days > MAX_ALLOWED_AGE_DAYS:
        sys.exit(
            f"ERROR: Newest date in CSV is {newest.date()} ({age_days} days ago). "
            f"Expected within {MAX_ALLOWED_AGE_DAYS} days — aborting."
        )

    # Defensive MTD filter (the dashboard preset is already "this month").
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
        "Website":     df[site_col].astype(str).str.strip(),
        "Impressions": pd.to_numeric(df[impr_col], errors="coerce").fillna(0),
        "Revenue":     pd.to_numeric(df[rev_col],  errors="coerce").fillna(0),
        "eCPM":        pd.to_numeric(df[ecpm_col], errors="coerce").fillna(0),
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
        [
            str(r["Website"]).strip(),
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

    # Preserve previous-month rows; only the current month gets refreshed.


    existing = service.spreadsheets().values().get(


        spreadsheetId=SPREADSHEET_ID, range=SHEET_NAME,


    ).execute().get("values", [])


    current_month = datetime.now().strftime("%Y-%m")


    preserved = []


    for _r in existing:


        if not _r or _r == HEADER:


            continue


        if len(_r) >= 2 and str(_r[1]).strip().startswith(current_month):


            continue  # current-month row → will be replaced by fresh data


        preserved.append(_r)


    merged = rows + preserved


    merged.sort(key=lambda r: (str(r[1]) if len(r) > 1 else "", str(r[0]) if len(r) > 0 else ""))


    merged.sort(key=lambda r: str(r[1]) if len(r) > 1 else "", reverse=True)



    final_rows = [HEADER] + merged
    log(f"Writing {len(merged)} rows = {len(rows)} new MTD + {len(preserved)} preserved (history)…")
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
    log("=== Kuantyx → Google Sheets daily sync ===")
    cfg   = load_config()
    creds = load_service_account()

    csv_path = download_kuantyx_csv(cfg["kuantyx_username"], cfg["kuantyx_password"])
    df       = process_csv(csv_path)
    write_sheet(df, creds)

    try: csv_path.unlink()
    except Exception: pass
    log("=== Done. ===")


if __name__ == "__main__":
    main()
