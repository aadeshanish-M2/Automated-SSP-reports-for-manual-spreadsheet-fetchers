#!/usr/bin/env python3
"""
MGID Publisher Reports API → Google Sheets daily revenue sync.

Pulls month-to-date data grouped by Date + Website via MGID's
v2 website-custom-report endpoint, computes CPM = Revenue/Clicks*1000,
and replaces the destination sheet contents (MTD-only mode).

If the API response is empty, missing expected fields, or stale, the sheet
is left untouched.
"""

import json
import os
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ── Configuration ─────────────────────────────────────────────────────────────

SCRIPT_DIR     = Path(__file__).parent
CONFIG_PATH    = SCRIPT_DIR / "config.json"
SA_PATH        = SCRIPT_DIR / "service_account.json"
SPREADSHEET_ID = os.environ.get("SHEETS_SPREADSHEET_ID") or "1cUNyk4Zrx2ZTZCjVl0O97zcNqwefr3yofJCz88HJpEc"
SHEET_NAME     = "Sheet1"
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]

API_BASE       = "https://api.mgid.com/v2/pub/account"
DATE_INTERVAL  = "thisMonth"      # MGID preset = month-to-date
DIMENSIONS     = "date,website"
METRICS        = "revenue,pageViews,adCPM"

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


_REQUIRED_KEYS = ("mgid_client_id", "mgid_api_token")


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


# ── Step 1 — Fetch report from MGID API ───────────────────────────────────────

def fetch_report(client_id: str, token: str) -> list[dict]:
    url = f"{API_BASE}/{client_id}/website-custom-report"
    params = {
        "dateInterval": DATE_INTERVAL,
        "dimensions":   DIMENSIONS,
        "metrics":      METRICS,
    }
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept":        "application/json",
    }
    log(f"Requesting MGID report ({DATE_INTERVAL}, dims={DIMENSIONS}, metrics={METRICS})…")
    r = requests.get(url, params=params, headers=headers, timeout=60)
    if r.status_code != 200:
        sys.exit(f"ERROR: MGID API failed ({r.status_code}): {r.text[:400]}")
    try:
        data = r.json()
    except ValueError:
        sys.exit(f"ERROR: MGID response was not valid JSON: {r.text[:300]}")

    # The API can return either a list of rows directly, or {"data": [...]},
    # or {"errors": [...]} — handle each.
    if isinstance(data, dict):
        if data.get("errors"):
            sys.exit(f"ERROR: MGID returned errors: {data['errors']}")
        if isinstance(data.get("data"), list):
            data = data["data"]
        else:
            # Sometimes the payload is wrapped in {"result": [...]} or similar.
            for k in ("result", "items", "rows"):
                if isinstance(data.get(k), list):
                    data = data[k]
                    break
            else:
                sys.exit(f"ERROR: Unexpected MGID response shape: {str(data)[:300]}")

    if not isinstance(data, list):
        sys.exit(f"ERROR: Expected list of rows from MGID, got {type(data).__name__}.")

    log(f"MGID returned {len(data)} rows.")
    return data


# ── Step 2 — Validate & shape ─────────────────────────────────────────────────

def process_rows(rows: list[dict]) -> pd.DataFrame:
    if not rows:
        sys.exit("ERROR: MGID returned 0 rows — aborting to protect the sheet.")

    df = pd.DataFrame(rows)
    df.columns = [c.strip() for c in df.columns]

    # Field-name candidates — MGID's exact casing isn't documented.
    def find_col(*cands):
        norm = {c.lower(): c for c in df.columns}
        for c in cands:
            if c.lower() in norm:
                return norm[c.lower()]
        return None

    date_col    = find_col("date", "day")
    website_col = find_col("website", "domain", "site")
    impr_col    = find_col("pageViews", "impressions", "imps", "clicks")
    rev_col     = find_col("revenue", "earnings")
    cpm_col     = find_col("adCPM", "cpm", "eCPM")

    missing = [name for name, col in [
        ("date", date_col), ("website", website_col),
        ("impressions", impr_col), ("revenue", rev_col),
    ] if col is None]
    if missing:
        sys.exit(
            f"ERROR: Missing expected fields {missing} in MGID response.\n"
            f"       Got columns: {list(df.columns)}"
        )

    df = df.dropna(subset=[website_col, date_col])
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

    # MGID's thisMonth preset already returns MTD, but defensively re-filter
    # in case the preset behaviour ever shifts.
    first_of_month = pd.Timestamp(datetime.now().replace(day=1).date())
    _pre_mtd_df = df.copy()
    before = len(df)
    df = df[df["__date"] >= first_of_month]
    if before != len(df):
        log(f"Filtered to MTD ({first_of_month.date()} onward): kept {len(df)}, dropped {before - len(df)}.")

    impressions = pd.to_numeric(df[impr_col], errors="coerce").fillna(0)
    revenue     = pd.to_numeric(df[rev_col],   errors="coerce").fillna(0)
    if cpm_col is not None:
        cpm = pd.to_numeric(df[cpm_col], errors="coerce").fillna(0)
    else:
        cpm = (revenue.where(impressions > 0, 0)
               / impressions.where(impressions > 0, 1) * 1000)

    out = pd.DataFrame({
        "Date":        df["__date"].dt.strftime("%Y-%m-%d"),
        "Website":     df[website_col].astype(str).str.strip(),
        "Impressions": impressions.astype(int),
        "Revenue":     revenue.round(4),
        "CPM":         cpm.round(4),
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
            str(r["Website"]).strip(),
            str(r["Date"]).strip(),
            f"${float(pd.to_numeric(r['Revenue'], errors='coerce') or 0):.2f}",
            int(pd.to_numeric(r["Impressions"], errors="coerce") or 0),
            float(pd.to_numeric(r["CPM"], errors="coerce") or 0),
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
    log("=== MGID API → Google Sheets daily sync ===")
    cfg   = load_config()
    creds = load_service_account()

    rows = fetch_report(cfg["mgid_client_id"], cfg["mgid_api_token"])
    df   = process_rows(rows)
    write_sheet(df, creds)
    log("=== Done. ===")


if __name__ == "__main__":
    main()
