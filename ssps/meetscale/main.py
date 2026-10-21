#!/usr/bin/env python3
"""
Meetscale / Sparteo (Data API) → Google Sheets daily sync.   [SERVER-SIDE / S2S]

This Meetscale account is a server-side (S2S) integration, so revenue is written
net of a 7% service-provider fee (see SERVER_SIDE_FEE); CPM is recomputed from the
reduced revenue and drops the same 7%. (partner_revenue is already the publisher's
estimated NET revenue; the 7% is MonetizeMore's S2S service fee on top of that.)

Pure API integration — no browser. Logs in with the platform credentials to get a
short-lived access token, then POSTs a month-to-date data query (grouped by domain;
the API returns one row per domain+date, with event_date on every row), and syncs
to the destination sheet in the uniform format
(Domain | Date | Revenue | Impression | CPM). Previous-month history is preserved;
only current-month rows are refreshed.

Mapping (Meetscale → uniform):
  Domain     = page_domain
  Date       = event_date         (returned on every row automatically)
  Revenue    = partner_revenue    (Estimated Partner Net Revenue, USD, full precision)
  Impression = imps
  CPM        = recomputed revenue/imps*1000 (net of the fee; ~= rpm_impression×0.93)

API ref: https://docs.meetscale.com/platform/meetscale-data-api-documentation
  Base : https://api.meetscale.com/v1/
  Login: POST /login  {username,password} → {access_token}  (short-lived; log in per run)
  Data : POST /data   Authorization: Bearer <token>
         body {columns:[page_domain], metrics:[imps,partner_revenue,rpm_impression],
               filters:[{col:event_date, op:>=, val:<first-of-month>}], currency:"USD"}
         → {results:[{page_domain, event_date, imps, partner_revenue, rpm_impression}]}
  NOTE: event_date is a FILTER column only (not groupable), but is returned on every
        row, so a single call yields the full daily breakdown.

NOTE: This integration is newly connected and barely active (ramping). A run that
finds no monetized activity is a normal "nothing to report yet" state, so it exits
cleanly and leaves the sheet untouched rather than erroring.
"""

import json
import os
import sys
import time
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
SPREADSHEET_ID = os.environ.get("SHEETS_SPREADSHEET_ID") or "18rtXqxd86m2AnETKbt6lhVWxL1PgmDni5FZfNzmX20w"
SHEET_NAME     = "Sheet1"
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]

API_BASE  = "https://api.meetscale.com/v1"
LOGIN_URL = f"{API_BASE}/login"
DATA_URL  = f"{API_BASE}/data"

MAX_ALLOWED_AGE_DAYS = 5
HEADER = ["Domain", "Date", "Revenue", "Impression", "CPM"]

# Server-side (S2S) integration: 7% service-provider fee deducted from publisher-net
# revenue. Revenue is written net; CPM (= revenue/impressions*1000) drops the same 7%
# since it is recomputed from the reduced revenue. Override with MEETSCALE_SERVER_SIDE_FEE.
SERVER_SIDE_FEE = float(os.environ.get("MEETSCALE_SERVER_SIDE_FEE") or 0.07)

# Revenue column display format: min 2 decimals, up to 6 shown ($17.91, $0.28086,
# $0.000186). Stamped on the revenue column after every write because
# values().update() with a "$"-prefixed string makes Sheets auto-apply a sticky
# currency format matching that string's decimal count — which would truncate the
# display of full-precision values on the next run. Stored values stay full precision.
REVENUE_COL_INDEX     = 2   # 0-based: Domain=0, Date=1, Revenue=2
REVENUE_NUMBER_FORMAT = '"$"#,##0.00####'


def _empty_data_exit(reason: str) -> None:
    # This integration is newly connected and not yet monetizing, so "no revenue"
    # is a valid state, not a failure. Exit cleanly WITHOUT touching the sheet so any
    # preserved history stays in place. (Tighten to a hard error once it is reliably
    # earning, to catch a genuinely broken pull.)
    log(f"No monetized activity yet ({reason}); Meetscale is still ramping. "
        "Leaving existing sheet data in place. Skipping this run.")
    sys.exit(0)


def _fmt_money(v) -> str:
    """
    Format a dollar value with up to 6 decimals (trailing zeros trimmed, min 2 dp)
    instead of a flat 2 dp. At sub-cent daily revenues — and after the ×0.93 fee,
    which produces long decimals — 2 dp rounding shifts each row by up to half a cent,
    and those errors accumulate when many days/domains are summed. 6 dp keeps every
    stored value effectively lossless while staying readable (e.g. $17.91, $0.000186).
    """
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
    server-side fee here so both the stored revenue and the recomputed CPM carry it.
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


_REQUIRED_KEYS = ("meetscale_username", "meetscale_password")


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


# ── Step 1 — Fetch from the Meetscale Data API ────────────────────────────────

def _post_with_retry(sess, url, json_body, headers):
    """POST with exponential backoff on 5xx and connection errors."""
    attempts = 4
    r = None
    for i in range(1, attempts + 1):
        try:
            r = sess.post(url, json=json_body, headers=headers, timeout=90)
        except requests.RequestException as e:
            if i == attempts:
                sys.exit(f"ERROR: Meetscale request failed (connection error: {e}). "
                         "Likely a temporary network/platform issue; sheet left unchanged.")
            log(f"  attempt {i}/{attempts}: connection error ({e}); retrying…")
            time.sleep(min(2 ** i, 20))
            continue
        if r.status_code < 500:
            return r
        if i == attempts:
            sys.exit(f"ERROR: Meetscale returned HTTP {r.status_code} after {attempts} "
                     "attempts — a platform-side server error, not the sync. Sheet left unchanged.")
        log(f"  attempt {i}/{attempts}: HTTP {r.status_code}; retrying…")
        time.sleep(min(2 ** i, 20))
    return r


def fetch_report(cfg: dict) -> list:
    """Log in for a fresh token, then POST the MTD data query. Returns row dicts."""
    sess = requests.Session()

    log("Logging in to Meetscale…")
    r = _post_with_retry(sess, LOGIN_URL,
                        {"username": cfg["meetscale_username"],
                         "password": cfg["meetscale_password"]},
                        {"Content-Type": "application/json"})
    if r.status_code != 200:
        sys.exit(f"ERROR: Meetscale login failed ({r.status_code}): {r.text[:300]} — "
                 "check the username/password. Sheet left unchanged.")
    try:
        token = r.json().get("access_token")
    except Exception as e:
        sys.exit(f"ERROR: could not parse login response: {e}")
    if not token:
        sys.exit(f"ERROR: no access_token in login response: {str(r.json())[:200]}")

    today = date.today()
    first = date(today.year, today.month, 1)
    body = {
        "columns": ["page_domain"],
        "metrics": ["imps", "partner_revenue", "rpm_impression"],
        # event_date is a filter-only column but is returned on every row, so one
        # call yields the full daily breakdown for the MTD window.
        "filters": [{"col": "event_date", "op": ">=", "val": first.isoformat()}],
        "currency": "USD",
    }
    log(f"Requesting MTD data {first.isoformat()} → {today.isoformat()} "
        "(columns=page_domain, metrics=imps,partner_revenue)…")
    rr = _post_with_retry(sess, DATA_URL, body,
                         {"Authorization": f"Bearer {token}",
                          "Content-Type": "application/json"})
    if rr.status_code == 401:
        sys.exit("ERROR: 401 unauthorized — Meetscale rejected the token. "
                 "Check the credentials. Sheet left unchanged.")
    if rr.status_code != 200:
        sys.exit(f"ERROR: data request failed ({rr.status_code}): {rr.text[:300]}")
    try:
        j = rr.json()
    except Exception as e:
        sys.exit(f"ERROR: could not parse data JSON: {e}")
    rows = j.get("results") if isinstance(j, dict) else j
    rows = rows or []
    log(f"API returned {len(rows)} rows.")
    return rows


# ── Step 2 — Validate & shape ─────────────────────────────────────────────────

def process_results(rows: list) -> pd.DataFrame:
    log("Shaping report rows…")
    if not rows:
        _empty_data_exit("API returned no rows for the current month")

    df = pd.DataFrame(rows)
    for col in ("event_date", "page_domain", "partner_revenue", "imps"):
        if col not in df.columns:
            sys.exit(f"ERROR: expected field {col!r} missing from API response. "
                     f"Got: {list(df.columns)} — aborting to protect the sheet.")

    df["Domain"] = df["page_domain"].astype(str).str.strip()

    df = df.dropna(subset=["event_date"])
    df["__date"] = pd.to_datetime(df["event_date"], errors="coerce")
    bad = df["__date"].isna().sum()
    if bad:
        log(f"WARNING: {bad} row(s) had unparseable dates and will be dropped.")
    df = df[df["__date"].notna()].copy()
    if df.empty:
        _empty_data_exit("no valid rows after date parsing")

    # Freshness guard, current-month-tolerant: only hard-fail when the newest date
    # is OUTSIDE the current month (a genuinely stuck/wrong-month pull). A low or
    # lagging newest date within the current month is acceptable for this account.
    newest = df["__date"].max()
    age_days = (datetime.now() - newest).days
    newest_in_current_month = newest.strftime("%Y-%m") == datetime.now().strftime("%Y-%m")
    if age_days > MAX_ALLOWED_AGE_DAYS and not newest_in_current_month:
        sys.exit(
            f"ERROR: Newest date is {newest.date()} ({age_days} days ago) and is "
            f"outside the current month — looks stuck/broken, aborting."
        )

    # MTD filter with month-boundary fallback.
    first_of_month = pd.Timestamp(datetime.now().replace(day=1).date())
    _pre_mtd_df = df.copy()
    df = df[df["__date"] >= first_of_month]
    if df.empty:
        log("WARNING: 0 rows match current-month filter — falling back to full report.")
        df = _pre_mtd_df

    out = pd.DataFrame({
        "Date":        df["__date"].dt.strftime("%Y-%m-%d"),
        "Domain":      df["Domain"].astype(str).str.strip(),
        "Impressions": pd.to_numeric(df["imps"], errors="coerce").fillna(0),
        "Revenue":     pd.to_numeric(df["partner_revenue"], errors="coerce").fillna(0),
        "CPM":         0.0,   # recomputed (net of fee) in _normalize_and_aggregate
    })
    out = out[~out["Domain"].str.lower().isin(["", "nan", "none"])]
    # Drop no-activity rows (0 impressions AND $0) — keep any day with real activity.
    out = out[(pd.to_numeric(out["Impressions"], errors="coerce").fillna(0) > 0)
              | (pd.to_numeric(out["Revenue"], errors="coerce").fillna(0) > 0)]
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
    log("=== Meetscale → Google Sheets daily sync ===")
    cfg   = load_config()
    creds = load_service_account()

    rows = fetch_report(cfg)
    df = process_results(rows)
    write_sheet(df, creds)
    log("=== Done. ===")


if __name__ == "__main__":
    main()
