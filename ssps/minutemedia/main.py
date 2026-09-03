#!/usr/bin/env python3
"""
Minute Media / Voltax (monetizemore.voltax.io) → Google Sheets daily sync.

Pure API integration against Minute Media's documented Reports API — no browser,
no Auth0 login (the dashboard login is CAPTCHA-gated). Authenticates with a
per-user API key via the `MM-API-Key` scheme, POSTs a month-to-date report query
to the `pubs_external` datasource, and syncs to the destination sheet in the
uniform format (Domain | Date | Revenue | Impression | CPM). Previous-month
history in the sheet is preserved; only current-month rows are refreshed.

API reference (Minute Media "API Data Retrieval Guide", 09/2025):
  POST {BASE}/organization/{org}/datasources/{datasource}/report-query
  Headers: Tenant: mmplus | Authorization: MM-API-Key <key> | Content-Type: application/json
  Body: dateRangePreset "month" = current month-to-date.
Field slugs (pubs_external): domain, date, total=Net Revenue,
  impressions=Impressions, net_cpm=CPM.
"""

import json
import os
import sys
from datetime import datetime, date, timedelta

import pandas as pd
import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build

from pathlib import Path

# ── Configuration ─────────────────────────────────────────────────────────────

SCRIPT_DIR     = Path(__file__).parent
CONFIG_PATH    = SCRIPT_DIR / "config.json"
SA_PATH        = SCRIPT_DIR / "service_account.json"
SPREADSHEET_ID = os.environ.get("SHEETS_SPREADSHEET_ID") or "1AcWUPmjQTzE36fuHx0U2OmzvMQJhpMFm_vcwumAVAS0"
SHEET_NAME     = "Sheet1"
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]

API_BASE     = "https://us-central1-bqservingevents.cloudfunctions.net/mm-reports-prod/api"
ORGANIZATION = "monetize-more"
DATASOURCE   = "pubs_external"
TENANT       = "mmplus"
REPORT_URL   = f"{API_BASE}/organization/{ORGANIZATION}/datasources/{DATASOURCE}/report-query"

MAX_ALLOWED_AGE_DAYS = 5
HEADER = ["Domain", "Date", "Revenue", "Impression", "CPM"]

# At the start of a month, the current month's data may not be populated yet.
# Within this window an empty pull is treated as normal (not a failure): we exit
# cleanly WITHOUT touching the sheet, so preserved previous-month history keeps
# showing until new data flows. Outside the window, an empty pull is a hard error.
MONTH_START_GRACE_DAYS = 5


def _empty_data_exit(reason: str) -> None:
    day = datetime.now().day
    if day <= MONTH_START_GRACE_DAYS:
        log(f"No current-month data yet ({reason}); within first "
            f"{MONTH_START_GRACE_DAYS} days of the month → treating as normal "
            "reporting lag, not a failure. Leaving existing sheet data in place. "
            "Skipping this run.")
        sys.exit(0)
    sys.exit(f"ERROR: {reason} — aborting.")


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


_REQUIRED_KEYS = ("minutemedia_api_key",)


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


# ── Step 1 — Fetch the report from the Minute Media API ───────────────────────

def fetch_report(api_key: str) -> list:
    """POST a month-to-date report query and return the list of report rows."""
    log(f"Requesting MTD report (org={ORGANIZATION}, datasource={DATASOURCE})…")
    headers = {
        "Tenant": TENANT,
        "Authorization": f"MM-API-Key {api_key}",
        "Content-Type": "application/json",
    }
    _today = date.today(); _from = _today - timedelta(days=30)
    payload = {
        # Rolling ~30-day window (not the "month" preset) so the previous month's
        # final day is refreshed once it is no longer "today" — fixes month-boundary freeze.
        "dateRange": [_from.isoformat(), _today.isoformat()],
        "dateRangePreset": None,
        "dimensions": ["date", "domain"],
        "filters": [],
        "metrics": ["total", "impressions", "net_cpm"],
        "timezone": "UTC+0",
        "limit": 15000,
        "offset": 0,
    }
    try:
        resp = requests.post(REPORT_URL, headers=headers, data=json.dumps(payload), timeout=90)
    except Exception as e:
        sys.exit(f"ERROR: report request failed to send.\nDetail: {e}")

    if resp.status_code == 401:
        sys.exit("ERROR: 401 unauthorized — the Minute Media API key is invalid or expired. "
                 "Retrieve a fresh key from Explore in Voltax and update the secret.")
    if resp.status_code != 200:
        sys.exit(f"ERROR: report request failed ({resp.status_code}): {resp.text[:300]}")

    try:
        data = resp.json().get("data") or {}
    except Exception as e:
        sys.exit(f"ERROR: could not parse API JSON.\nDetail: {e}")
    reports = data.get("reports") or []
    log(f"API returned {len(reports)} rows (count={data.get('count')}).")
    return reports


# ── Step 2 — Validate & shape ─────────────────────────────────────────────────

def process_reports(reports: list) -> pd.DataFrame:
    log("Shaping report rows…")
    if not reports:
        _empty_data_exit("API returned no rows for the current month")

    df = pd.DataFrame(reports)
    for col in ("date", "domain", "total", "impressions", "net_cpm"):
        if col not in df.columns:
            sys.exit(f"ERROR: expected field {col!r} missing from API response. "
                     f"Got: {list(df.columns)} — aborting to protect the sheet.")

    df = df.dropna(subset=["date", "domain"])
    df["__date"] = pd.to_datetime(df["date"], errors="coerce")
    bad = df["__date"].isna().sum()
    if bad:
        log(f"WARNING: {bad} row(s) had unparseable dates and will be dropped.")
    df = df[df["__date"].notna()].copy()
    if df.empty:
        _empty_data_exit("no valid rows after date parsing")

    newest = df["__date"].max()
    age_days = (datetime.now() - newest).days
    if age_days > MAX_ALLOWED_AGE_DAYS:
        sys.exit(
            f"ERROR: Newest date is {newest.date()} ({age_days} days ago). "
            f"Expected within {MAX_ALLOWED_AGE_DAYS} days — aborting."
        )

    # Keep the FULL pulled window (rolling ~30 days) — do NOT MTD-filter, so the
    # previous month's final day is refreshed once it is no longer "today". The
    # write step's per-(Domain, Date) dedup overwrites any frozen partial and
    # older history is preserved. Fixes the month-boundary freeze.
    log(f"Keeping full pulled window: {len(df)} rows.")

    out = pd.DataFrame({
        "Date":        df["__date"].dt.strftime("%Y-%m-%d"),
        "Domain":      df["domain"].astype(str).str.strip(),
        "Impressions": pd.to_numeric(df["impressions"], errors="coerce").fillna(0),
        "Revenue":     pd.to_numeric(df["total"],       errors="coerce").fillna(0),
        "CPM":         pd.to_numeric(df["net_cpm"],      errors="coerce").fillna(0),
    })
    out = out[~out["Domain"].str.lower().isin(["", "nan", "none"])]
    if out.empty:
        _empty_data_exit("no valid rows after cleaning")

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
            str(r["Domain"]).strip(),
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
    log("=== Minute Media → Google Sheets daily sync ===")
    cfg   = load_config()
    creds = load_service_account()

    reports = fetch_report(cfg["minutemedia_api_key"])
    df      = process_reports(reports)
    write_sheet(df, creds)
    log("=== Done. ===")


if __name__ == "__main__":
    main()
