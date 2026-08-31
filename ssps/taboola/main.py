#!/usr/bin/env python3
"""
Taboola (Backstage API) → Google Sheets daily sync.

Pure API integration — no browser. Authenticates via OAuth2 client-credentials,
pulls the month-to-date revenue-summary report broken down by day + site +
placement, aggregates to one row per (site, date), and syncs to the destination
sheet in the uniform format (Domain | Date | Revenue | Impression | CPM).
Previous-month history is preserved; only current-month rows are refreshed.

Notes on this account (monetizemorereseller-network):
  • It exposes a single publisher ("Monetize More Reseller"), so the sheet is
    effectively one row per date. The code still aggregates by (publisher, date),
    so if Taboola ever exposes multiple sites they each get their own rows.
  • Taboola reports in the ACCOUNT's native currency (CAD for this account), so
    the Revenue column is CAD, not USD. Impression = page_views; CPM is recomputed
    from revenue/page_views*1000 (equals Taboola's ad_rpm).

API ref: Backstage API - Publisher Reports (revenue-summary).
"""

import json
import os
import sys
from datetime import date, datetime

import pandas as pd
import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build

from pathlib import Path

# ── Configuration ─────────────────────────────────────────────────────────────

SCRIPT_DIR     = Path(__file__).parent
CONFIG_PATH    = SCRIPT_DIR / "config.json"
SA_PATH        = SCRIPT_DIR / "service_account.json"
SPREADSHEET_ID = os.environ.get("SHEETS_SPREADSHEET_ID") or "1P3kHo5OUjRFDTdJe54DZEyo4stqNJdIsTQvF4biCpL4"
SHEET_NAME     = "Sheet1"
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]

BACKSTAGE  = "https://backstage.taboola.com/backstage"
TOKEN_URL  = f"{BACKSTAGE}/oauth/token"
DIMENSION  = "day_site_placement_breakdown"   # date + publisher + page_type + placement

MAX_ALLOWED_AGE_DAYS = 5
HEADER = ["Domain", "Date", "Revenue", "Impression", "CPM"]

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
    root domain, then sum Impression + Revenue across same (Domain, Date)
    groups. CPM recomputed from totals (Revenue/Impression * 1000). Non-domain
    labels (e.g. "Monetize More Reseller") are passed through unchanged.
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
        low = s.lower().split("/")[0].split(":")[0]
        if not low:
            return ""
        if _extract is not None:
            ext = _extract(low)
            if ext.domain and ext.suffix:
                return f"{ext.domain}.{ext.suffix}"
        if low.startswith("www."):
            return low[4:]
        # Not a domain — keep the original (readable) label.
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


_REQUIRED_KEYS = ("taboola_client_id", "taboola_client_secret", "taboola_account_id")


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


# ── Step 1 — Fetch the report from the Taboola Backstage API ──────────────────

def fetch_report(cfg: dict) -> list:
    """OAuth2 client-credentials → GET the MTD revenue-summary report. Returns rows."""
    log("Requesting OAuth2 token…")
    try:
        r = requests.post(
            TOKEN_URL,
            data={"client_id": cfg["taboola_client_id"],
                  "client_secret": cfg["taboola_client_secret"],
                  "grant_type": "client_credentials"},
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            timeout=30,
        )
    except Exception as e:
        sys.exit(f"ERROR: token request failed to send.\nDetail: {e}")
    if r.status_code != 200:
        sys.exit(f"ERROR: Taboola auth failed ({r.status_code}): {r.text[:300]}")
    token = r.json().get("access_token")
    if not token:
        sys.exit(f"ERROR: no access_token in auth response: {r.text[:300]}")

    today = date.today()
    first = date(today.year, today.month, 1)
    url = (f"{BACKSTAGE}/api/1.0/{cfg['taboola_account_id']}"
           f"/reports/revenue-summary/dimensions/{DIMENSION}")
    log(f"Requesting MTD report {first.isoformat()} → {today.isoformat()} ({DIMENSION})…")
    try:
        rr = requests.get(url,
                          params={"start_date": first.isoformat(), "end_date": today.isoformat()},
                          headers={"Authorization": f"Bearer {token}"}, timeout=90)
    except Exception as e:
        sys.exit(f"ERROR: report request failed to send.\nDetail: {e}")
    if rr.status_code == 401:
        sys.exit("ERROR: 401 unauthorized — Taboola token rejected. Check the "
                 "client id/secret and that the account id is correct.")
    if rr.status_code != 200:
        sys.exit(f"ERROR: report request failed ({rr.status_code}): {rr.text[:300]}")
    try:
        data = rr.json()
    except Exception as e:
        sys.exit(f"ERROR: could not parse API JSON.\nDetail: {e}")
    results = data.get("results") or []
    log(f"API returned {len(results)} rows (currency: "
        f"{results[0].get('currency') if results else '?'}, timezone: {data.get('timezone')}).")
    return results


# ── Step 2 — Validate & shape ─────────────────────────────────────────────────

def process_results(results: list) -> pd.DataFrame:
    log("Shaping report rows…")
    if not results:
        _empty_data_exit("API returned no rows for the current month")

    df = pd.DataFrame(results)
    for col in ("date", "publisher_name", "ad_revenue", "page_views"):
        if col not in df.columns:
            sys.exit(f"ERROR: expected field {col!r} missing from API response. "
                     f"Got: {list(df.columns)} — aborting to protect the sheet.")

    df = df.dropna(subset=["date", "publisher_name"])
    # Taboola dates look like "2026-08-30 00:00:00.0" — take the date part.
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

    # MTD filter with month-boundary fallback.
    first_of_month = pd.Timestamp(datetime.now().replace(day=1).date())
    _pre_mtd_df = df.copy()
    df = df[df["__date"] >= first_of_month]
    if df.empty:
        log("WARNING: 0 rows match current-month filter — falling back to full report "
            "(likely a month-boundary day).")
        df = _pre_mtd_df

    out = pd.DataFrame({
        "Date":        df["__date"].dt.strftime("%Y-%m-%d"),
        "Domain":      df["publisher_name"].astype(str).str.strip(),
        "Impressions": pd.to_numeric(df["page_views"], errors="coerce").fillna(0),
        "Revenue":     pd.to_numeric(df["ad_revenue"], errors="coerce").fillna(0),
        "CPM":         0.0,   # recomputed from totals in _normalize_and_aggregate
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

    existing = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=_tab,
    ).execute().get("values", [])
    current_month = datetime.now().strftime("%Y-%m")
    preserved = []
    for _r in existing:
        if not _r or _r == HEADER:
            continue
        if len(_r) >= 2 and str(_r[1]).strip().startswith(current_month):
            continue
        preserved.append(_r)

    merged = rows + preserved

    # Guarantee ONE row per (Domain, Date) across the WHOLE sheet; fresh rows win.
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
            if isinstance(cell, float) and (cell != cell):
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
    log("=== Taboola → Google Sheets daily sync ===")
    cfg   = load_config()
    creds = load_service_account()

    results = fetch_report(cfg)
    df      = process_results(results)
    write_sheet(df, creds)
    log("=== Done. ===")


if __name__ == "__main__":
    main()
