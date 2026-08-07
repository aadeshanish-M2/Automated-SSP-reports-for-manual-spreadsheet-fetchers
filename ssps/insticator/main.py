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
            # The post-login redirect can fire before wait_for_url starts
            # listening (or navigate via SPA without a load event), so don't hard
            # fail on the navigation wait — verify by URL afterwards instead.
            try:
                page.wait_for_url(
                    lambda u: "signin" not in u.lower() and "login" not in u.lower(),
                    timeout=30_000,
                )
            except PlaywrightTimeoutError:
                pass
            page.wait_for_timeout(5000)
            if "signin" in page.url.lower() or "login" in page.url.lower():
                browser.close()
                sys.exit("ERROR: Still on login page after submit — check credentials.")
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
            # The report list grows over time and the page can render slowly, so
            # the button occasionally isn't immediately actionable. Wait for it to
            # be visible, scroll it into view, then click — retry once if the first
            # attempt races the page render.
            btn = page.get_by_role("button", name="Create New Report")
            if btn.count() == 0:
                btn = page.locator('button:has-text("Create New Report")')
            clicked = False
            for attempt in range(2):
                try:
                    btn.first.wait_for(state="visible", timeout=30_000)
                    btn.first.scroll_into_view_if_needed(timeout=10_000)
                    btn.first.click(timeout=20_000)
                    clicked = True
                    break
                except PlaywrightTimeoutError:
                    if attempt == 0:
                        page.wait_for_timeout(4000)
            if not clicked:
                raise PlaywrightTimeoutError("Create New Report not clickable after retry")
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
        # Insticator generates the report ASYNCHRONOUSLY and it is genuinely slow
        # — the row shows a "processing" clock for ~10-15 minutes before the icon
        # becomes a live download. Clicking before then does nothing (no download
        # event fires). So poll for up to ~20 minutes: each attempt reloads the
        # list, re-finds the row, and tries the download; it succeeds the moment
        # the report finishes generating.
        MAX_ATTEMPTS = 22          # ~22 × ~55s ≈ 20 min
        download = None
        last_err = None
        for attempt in range(MAX_ATTEMPTS):
            try:
                if attempt > 0:
                    page.reload(wait_until="domcontentloaded")
                    page.wait_for_timeout(6000)
                name_cell = page.get_by_text(report_name, exact=True).first
                name_cell.wait_for(state="visible", timeout=20_000)
                download_icon = name_cell.locator('xpath=..').locator('span.name-download').first
                download_icon.wait_for(state="visible", timeout=15_000)
                download_icon.scroll_into_view_if_needed(timeout=8000)
                with page.expect_download(timeout=20_000) as dl_info:
                    download_icon.click(timeout=15_000)
                download = dl_info.value
                break
            except PlaywrightTimeoutError as e:
                last_err = e
                log(f"  not ready yet (attempt {attempt + 1}/{MAX_ATTEMPTS}) — report "
                    "still generating (~10-15 min is normal); waiting before retry…")
                page.wait_for_timeout(25_000)
        if download is None:
            browser.close()
            sys.exit(f"ERROR: CSV download did not start in time.\nDetail: {last_err}")

        tmp = Path(tempfile.mktemp(suffix=".csv"))
        download.save_as(str(tmp))

        # Delete the report we just created. Insticator creates a NEW report every
        # run, and left uncleaned they pile up (hit 111 once) — a large backlog
        # congests report GENERATION, so new reports stay stuck "processing" and
        # downloads stop working entirely. Cleaning up after ourselves keeps the
        # list tiny. Best-effort: the data is already downloaded, so never fail
        # the sync if cleanup doesn't take.
        try:
            page.on("dialog", lambda d: d.accept())
            trash = (page.get_by_text(report_name, exact=True).first
                     .locator('xpath=..').locator('span.name-trash').first)
            trash.click(timeout=8000)
            page.wait_for_timeout(800)
            for _lbl in ("Delete", "Yes", "Confirm", "OK", "Remove"):
                _btn = page.locator(f'button:has-text("{_lbl}"):visible')
                if _btn.count() > 0:
                    _btn.first.click(timeout=4000)
                    break
            page.wait_for_timeout(1500)
            log("Cleaned up the generated report (prevents backlog buildup).")
        except Exception as _e:
            log(f"NOTE: could not delete the generated report (harmless): {str(_e)[:60]}")

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


    # Auto-detect the first tab name (sheet might not be called "Sheet1").
    try:
        _meta = service.spreadsheets().get(
            spreadsheetId=SPREADSHEET_ID, fields="sheets.properties.title"
        ).execute()
        _tab = _meta["sheets"][0]["properties"]["title"]
    except Exception as _e:
        _tab = SHEET_NAME
    log(f"Using sheet tab: {_tab!r}")
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


            continue  # current-month row → will be replaced by fresh data


        preserved.append(_r)


    merged = rows + preserved

    # Guarantee ONE row per (Domain, Date) across the WHOLE sheet — not just the
    # fresh rows. `rows` (this run's data) are listed first, so on any collision
    # the fresh value wins and stale/duplicate preserved-history rows are dropped.
    # This makes every write self-healing: it repairs pre-existing duplicate
    # history (which had inflated month totals) instead of freezing it forever.
    _seen = set()
    _deduped = []
    for _r in merged:
        _key = (str(_r[0]).strip().lower(), str(_r[1]).strip()) if len(_r) >= 2 else None
        if _key is not None and _key in _seen:
            continue
        if _key is not None:
            _seen.add(_key)
        _deduped.append(_r)
    merged = _deduped


    merged.sort(key=lambda r: (str(r[1]) if len(r) > 1 else "", str(r[0]) if len(r) > 0 else ""))


    merged.sort(key=lambda r: str(r[1]) if len(r) > 1 else "", reverse=True)



    def _stringify(cell):
        if cell is None:
            return ""
        if isinstance(cell, (int, float)):
            if isinstance(cell, float) and (cell != cell):  # NaN check
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
