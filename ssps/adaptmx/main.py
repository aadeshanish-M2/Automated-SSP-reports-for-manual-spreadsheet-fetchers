#!/usr/bin/env python3
"""
AdaptMX (Reporting API v2) → Google Sheets daily sync.  [SERVER-SIDE / S2S]

This AdaptMX account is a server-side (S2S) integration, so revenue is written net of
a 7% service-provider fee (see SERVER_SIDE_FEE); CPM is recomputed from the reduced
revenue and drops the same 7%.

DEMAND FILTER: AdaptMX reports several demand types. "Magnite PubDirect" (API enum
PUB_DIRECT, Magnite direct deals) is NOT paid to us, so it is EXCLUDED. Everything
else is kept: "AMX RTB" (WEB_AUDIENCE_DEMAND, the demand AdaptMX pays us for) plus
any Yieldmo-direct / managed demand if it ever appears. Labels seen are logged on
every run.

Pure API integration, no browser. Paginated GET grouped by day + domain + demand_type,
synced to the destination sheet in the uniform format
(Domain | Date | Revenue | Impression | CPM). Older history is preserved.

Mapping (AdaptMX → uniform):
  Domain     = domain (collapsed to its root domain, subdomains summed)
  Date       = day ("2026-11-01T00:00:00+00:00" → 2026-11-01, tz=UTC)
  Revenue    = revenue (written net of the 7% fee)
  Impression = impressions
  CPM        = recomputed revenue/impressions*1000 (net of the fee)

API ref: https://docs.adaptmx.com/reference/getapiv2reportingaccount
  GET https://reporting.adaptmx.com/api/v2/reporting/account
  Auth : X-API-Key: <key>   (dashboard → click your email → account details)
  Query: dimensions[]=daily&dimensions[]=domain&dimensions[]=demand_type
         &metrics[]=impressions&metrics[]=revenue&tz=UTC
         &start_date=YYYY-MM-DDT00:00:00Z&end_date=YYYY-MM-DDT23:59:59Z&page=&page_size=
  → {data:[{day,domain,demand_type,impressions,revenue}], pagination:{total_pages,...}, errors}
"""

import json
import os
import sys
import time
from datetime import date, datetime, timedelta

import pandas as pd
import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build

from pathlib import Path

# ── Configuration ─────────────────────────────────────────────────────────────

SCRIPT_DIR     = Path(__file__).parent
CONFIG_PATH    = SCRIPT_DIR / "config.json"
SA_PATH        = SCRIPT_DIR / "service_account.json"
SPREADSHEET_ID = os.environ.get("SHEETS_SPREADSHEET_ID") or "1ZuAqPO3mO_vEETFP_ahp0oCH8P1ruQjCSnLnDe_7tl0"
SHEET_NAME     = "Sheet1"
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]

REPORT_URL = os.environ.get("ADAPTMX_REPORT_URL") or "https://reporting.adaptmx.com/api/v2/reporting/account"
PAGE_SIZE  = 10000        # API max
MAX_PAGES  = 200          # safety cap on pagination

# Demand types NOT paid to us → excluded. Matched case-insensitively against the
# returned demand_type label (e.g. "Magnite PubDirect") and the API enum (PUB_DIRECT).
EXCLUDED_DEMAND = ("magnite", "pubdirect", "pub_direct")
# …but never exclude Yieldmo direct (enum PUB_DIRECT_YIELDMO) via the pub_direct match.
_KEEP_OVERRIDE  = ("yieldmo",)

MAX_ALLOWED_AGE_DAYS = 5
HEADER = ["Domain", "Date", "Revenue", "Impression", "CPM"]

# Pull window reaches back this far so the LAST day of a month keeps refreshing over
# the next month's first days (data settles a few days late) instead of freezing partial.
LOOKBACK_DAYS = 5

# Server-side (S2S) integration: 7% service-provider fee deducted from publisher-net
# revenue. Revenue is written net; CPM (= revenue/impressions*1000) drops the same 7%
# since it is recomputed from the reduced revenue. Override with ADAPTMX_SERVER_SIDE_FEE.
SERVER_SIDE_FEE = float(os.environ.get("ADAPTMX_SERVER_SIDE_FEE") or 0.07)

# Revenue column display format: min 2 decimals, up to 6 shown ($17.91, $0.000186).
# Stamped after every write because values().update() with a "$"-prefixed string makes
# Sheets auto-apply a sticky currency format that would truncate full-precision values.
REVENUE_COL_INDEX     = 2   # 0-based: Domain=0, Date=1, Revenue=2
REVENUE_NUMBER_FORMAT = '"$"#,##0.00####'


def _empty_data_exit(reason: str) -> None:
    # Newly connected / low-volume: "no revenue" is a valid state, not a failure.
    # Exit cleanly WITHOUT touching the sheet so preserved history stays in place.
    log(f"No monetized activity yet ({reason}); AdaptMX may still be ramping. "
        "Leaving existing sheet data in place. Skipping this run.")
    sys.exit(0)


def _fmt_money(v) -> str:
    """Format a dollar value with up to 6 decimals (trailing zeros trimmed, min 2 dp).
    At sub-cent daily revenues — and after the ×0.93 fee — flat 2 dp rounding would shift
    each row by up to half a cent and those errors accumulate when summed; 6 dp keeps every
    stored value effectively lossless while staying readable."""
    try:
        v = float(v)
    except Exception:
        v = 0.0
    s = f"{v:.6f}"
    if "." in s:
        intp, frac = s.split(".")
        frac = frac.rstrip("0")
        if len(frac) < 2:
            frac = (frac + "00")[:2]
        s = f"{intp}.{frac}"
    return f"${s}"


def _normalize_and_aggregate(rows):
    """
    Collapse URL paths AND subdomains in Domain values to their registered root
    domain, then sum Impression + Revenue across same (Domain, Date) groups. CPM
    recomputed from totals (Revenue/Impression * 1000). Revenue is taken net of the
    server-side fee here, so both the stored revenue and the CPM carry it. App bundle
    IDs (no dot / not a domain) pass through unchanged.
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
        return s

    bucket = {}
    for r in rows:
        domain = root(r[0])
        date_  = str(r[1]).strip()
        try:
            rev = float(str(r[2]).lstrip("$").replace(",", ""))
        except Exception:
            rev = 0.0
        # Server-side fee: report revenue net of the 7% service-provider fee.
        rev *= (1 - SERVER_SIDE_FEE)
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
        cpm = round(v["rev"] / v["imp"] * 1000, 6) if v["imp"] > 0 else 0.0
        new_rows.append([domain, date_, _fmt_money(v["rev"]), v["imp"], cpm])
    return new_rows


# ── Helpers ───────────────────────────────────────────────────────────────────

def log(msg: str) -> None:
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


_REQUIRED_KEYS = ("adaptmx_api_key",)


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
            f"Set env vars [{env_names}] or fill config.json. "
            "(AdaptMX dashboard → click your email in the top menu → API key.)"
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


# ── Step 1 — Fetch from the AdaptMX Reporting API ─────────────────────────────

def _get_with_retry(sess, params, headers):
    """GET with exponential backoff on 429/5xx and connection errors."""
    attempts = 5
    r = None
    for i in range(1, attempts + 1):
        try:
            r = sess.get(REPORT_URL, params=params, headers=headers, timeout=120)
        except requests.RequestException as e:
            if i == attempts:
                sys.exit(f"ERROR: AdaptMX request failed (connection error: {e}). Sheet left unchanged.")
            log(f"  attempt {i}/{attempts}: connection error ({e}); retrying…")
            time.sleep(min(2 ** i, 30))
            continue
        if r.status_code < 500 and r.status_code != 429:
            return r
        if i == attempts:
            sys.exit(f"ERROR: AdaptMX returned HTTP {r.status_code} after {attempts} attempts "
                     f"({r.text[:160]}). Sheet left unchanged.")
        wait = int(r.headers.get("Retry-After") or 0) or min(2 ** i, 30)
        log(f"  attempt {i}/{attempts}: HTTP {r.status_code}; retrying in {wait}s…")
        time.sleep(wait)
    return r


def fetch_report(cfg: dict) -> list:
    """Pull the report (day × domain × demand_type), paging through all results."""
    sess = requests.Session()
    headers = {"X-API-Key": cfg["adaptmx_api_key"], "Accept": "application/json"}
    today = date.today()
    first = date(today.year, today.month, 1)
    start = min(first, today - timedelta(days=LOOKBACK_DAYS))
    base = [
        ("dimensions[]", "daily"), ("dimensions[]", "domain"), ("dimensions[]", "demand_type"),
        ("metrics[]", "impressions"), ("metrics[]", "revenue"),
        ("tz", "UTC"),
        ("start_date", f"{start.isoformat()}T00:00:00Z"),
        ("end_date",   f"{today.isoformat()}T23:59:59Z"),
        ("page_size",  str(PAGE_SIZE)),
    ]
    log(f"Requesting report {start.isoformat()} → {today.isoformat()} "
        "(daily + domain + demand_type, UTC)…")
    all_rows = []
    page = 1
    while page <= MAX_PAGES:
        r = _get_with_retry(sess, base + [("page", str(page))], headers)
        if r.status_code in (401, 403):
            sys.exit(f"ERROR: {r.status_code} — the AdaptMX API key was rejected. Check "
                     "ADAPTMX_API_KEY (dashboard → click your email → API key). Sheet left unchanged.")
        if r.status_code != 200:
            sys.exit(f"ERROR: AdaptMX report failed ({r.status_code}): {r.text[:300]}")
        try:
            j = r.json()
        except Exception as e:
            sys.exit(f"ERROR: could not parse AdaptMX JSON: {e}")
        if j.get("errors"):
            sys.exit(f"ERROR: AdaptMX reported errors: {j['errors']}. Sheet left unchanged.")
        rows = j.get("data") or []
        all_rows.extend(rows)
        total_pages = (j.get("pagination") or {}).get("total_pages") or page
        log(f"  page {page}/{total_pages}: +{len(rows)} rows (total {len(all_rows)})")
        if page >= total_pages or not rows:
            break
        page += 1

    log(f"API returned {len(all_rows)} rows across {page} page(s).")
    return all_rows


def _is_excluded(label) -> bool:
    low = str(label or "").strip().lower()
    if any(k in low for k in _KEEP_OVERRIDE):
        return False
    return any(k in low for k in EXCLUDED_DEMAND)


# ── Step 2 — Validate & shape ─────────────────────────────────────────────────

def process_results(rows: list) -> pd.DataFrame:
    log("Shaping report rows…")
    if not rows:
        _empty_data_exit("API returned no rows for the pulled window")

    df = pd.DataFrame(rows)
    for col in ("day", "domain", "demand_type", "impressions", "revenue"):
        if col not in df.columns:
            sys.exit(f"ERROR: expected field {col!r} missing from API response. "
                     f"Got: {list(df.columns)} — aborting to protect the sheet.")

    # Demand filter — log every label with its totals so a new/renamed type is visible.
    df["revenue"] = pd.to_numeric(df["revenue"], errors="coerce").fillna(0)
    df["impressions"] = pd.to_numeric(df["impressions"], errors="coerce").fillna(0)
    for label, g in df.groupby("demand_type", dropna=False):
        tag = "EXCLUDED" if _is_excluded(label) else "kept"
        log(f"  demand_type {label!r}: {len(g)} rows, {int(g['impressions'].sum())} imps, "
            f"${g['revenue'].sum():.4f} gross → {tag}")
    df = df[~df["demand_type"].map(_is_excluded)].copy()
    if df.empty:
        _empty_data_exit("no paid (non-Magnite) demand rows in the pulled window")

    df["Domain"] = df["domain"].astype(str).str.strip()
    df["__date"] = pd.to_datetime(df["day"].astype(str).str[:10], errors="coerce")
    bad = df["__date"].isna().sum()
    if bad:
        log(f"WARNING: {bad} row(s) had unparseable dates and will be dropped.")
    df = df[df["__date"].notna()].copy()
    if df.empty:
        _empty_data_exit("no valid rows after date parsing")

    newest = df["__date"].max()
    age_days = (datetime.now() - newest).days
    newest_in_current_month = newest.strftime("%Y-%m") == datetime.now().strftime("%Y-%m")
    if age_days > MAX_ALLOWED_AGE_DAYS and not newest_in_current_month:
        sys.exit(
            f"ERROR: Newest date is {newest.date()} ({age_days} days ago) and is "
            f"outside the current month — looks stuck/broken, aborting."
        )

    log(f"Keeping full pulled window: {len(df)} rows, "
        f"{df['__date'].min().date()} → {df['__date'].max().date()}.")

    out = pd.DataFrame({
        "Date":        df["__date"].dt.strftime("%Y-%m-%d"),
        "Domain":      df["Domain"],
        "Impressions": df["impressions"],
        "Revenue":     df["revenue"],
        "CPM":         0.0,   # recomputed (net of fee) in _normalize_and_aggregate
    })
    out = out[~out["Domain"].str.lower().isin(["", "nan", "none"])]
    # Drop no-activity rows (0 impressions AND $0) — AdaptMX lists many zero rows.
    out = out[(out["Impressions"] > 0) | (out["Revenue"] > 0)]
    if out.empty:
        _empty_data_exit("no rows with monetized activity (impressions/revenue) yet")

    log(f"Valid: {len(out)} rows, dates {out['Date'].min()} → {out['Date'].max()}")
    return out


# ── Step 3 — Write to sheet (preserve previous-month history) ─────────────────

def write_sheet(df: pd.DataFrame, creds) -> None:
    log("Connecting to Google Sheets…")
    service = build("sheets", "v4", credentials=creds, cache_discovery=False)

    _sheet_id = 0
    try:
        _meta = service.spreadsheets().get(
            spreadsheetId=SPREADSHEET_ID, fields="sheets.properties(title,sheetId)"
        ).execute()
        _props = _meta["sheets"][0]["properties"]
        _tab = _props["title"]
        _sheet_id = _props["sheetId"]
    except Exception:
        _tab = SHEET_NAME
    log(f"Using sheet tab: {_tab!r}")

    rows = [
        [
            str(r["Domain"]).strip(),
            str(r["Date"]).strip(),
            _fmt_money(pd.to_numeric(r["Revenue"], errors="coerce") or 0),
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

    # Stamp the adaptive revenue number format on every data row so full precision
    # is visible (overrides the sticky currency format values().update() applies).
    try:
        fmt_rows = [
            {"values": [{"userEnteredFormat": {"numberFormat": {
                "type": "NUMBER", "pattern": REVENUE_NUMBER_FORMAT}}}]}
            for _ in range(len(merged))
        ]
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": [{"updateCells": {
                "start": {"sheetId": _sheet_id, "rowIndex": 1,
                          "columnIndex": REVENUE_COL_INDEX},
                "rows": fmt_rows,
                "fields": "userEnteredFormat.numberFormat",
            }}]},
        ).execute()
    except Exception as e:
        log(f"WARNING: could not stamp revenue number format (display only): {e}")

    log("Sheet updated successfully.")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    log("=== AdaptMX → Google Sheets daily sync (S2S, excl. Magnite PubDirect) ===")
    cfg   = load_config()
    creds = load_service_account()

    rows = fetch_report(cfg)
    df = process_results(rows)
    write_sheet(df, creds)
    log("=== Done. ===")


if __name__ == "__main__":
    main()
