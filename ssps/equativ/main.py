#!/usr/bin/env python3
"""
Equativ (Reporting API) → Google Sheets daily sync.   [SPLIT: CLIENT + SERVER-SIDE]

Equativ carries BOTH client-side and server-side demand on one account. We pull a
single month-to-date report grouped by Day + domain + ServerSideBiddingCallerName,
then SPLIT the rows by that caller into two separate sheets:

  • PrebidServer → the Equativ-S2S sheet, revenue net of a 7% S2S fee
  • N/A         → the Equativ (client) sheet, no fee
  • Google (and any other caller) → IGNORED — this is open-bidding revenue and must
    NOT appear in either sheet. (The earlier HeaderBiddingType split wrongly lumped
    Google open bidding in with server-side; the caller dimension separates it.)

CPM is recomputed from the (fee-adjusted) revenue on each sheet. Sub-cent values are
stored at full precision. Previous-month history is preserved on each sheet.

Mapping (Equativ → uniform):
  Domain     = AppOrSiteDomain        (web domain or app bundle id)
  Date       = Day
  Revenue    = EstimatedPublisherNetRevenuePublisherCurrency  ("Estimated Net Revenue")
               — the only populated NET revenue metric for this account (PublisherNet-
               Revenue / AdServing / AdNetwork all come back 0; Gross is ~15% higher).
  Impression = Impressions
  CPM        = recomputed revenue/impressions*1000 (net of the fee on the S2S sheet)
  split on   = ServerSideBiddingCallerName  (PrebidServer=S2S · N/A=client · Google=ignore)

API ref: https://help.equativ.com/reporting-api-get-started
  Token : POST https://login.eqtv.io/oauth2/token  (OAuth2 client_credentials → Bearer)
  Report: POST https://supply-api.eqtv.io/report-async        → report id (async)
          GET  https://supply-api.eqtv.io/report-async/{id}   → status + instanceId (CSV URL)
          status 1 = SUCCESS; download the signed .csv.gz (semicolon-separated).
"""

import csv
import gzip
import io
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
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]

# Two destination sheets (one Equativ account, split by integration type).
SERVER_SPREADSHEET_ID = os.environ.get("EQUATIV_SERVER_SPREADSHEET_ID") or "19ft_HDZgoV6kfG0PL93JbAHlcowymAj4BE8tgL99URg"
CLIENT_SPREADSHEET_ID = os.environ.get("EQUATIV_CLIENT_SPREADSHEET_ID") or "1-GqZamGxysV8xyF7yS_qeoK-ava1v-uhDczqi-a3okw"
SHEET_NAME            = "Sheet1"

TOKEN_URL   = "https://login.eqtv.io/oauth2/token"
REPORT_URL  = "https://supply-api.eqtv.io/report-async"
POLL_EVERY  = 6
POLL_MAX_S  = 420          # up to 7 min for the async report to finish
STATUS_SUCCESS, STATUS_FAILED, STATUS_DELETED = 1, 9, 17

MAX_ALLOWED_AGE_DAYS = 5
HEADER = ["Domain", "Date", "Revenue", "Impression", "CPM"]

# Pull window reaches back this far so the LAST day of a month keeps refreshing over
# the next month's first days (data settles a few days late) instead of freezing partial.
LOOKBACK_DAYS = 5

# Server-side (S2S) fee: 7% deducted from publisher-net revenue on the S2S sheet only.
SERVER_SIDE_FEE = float(os.environ.get("EQUATIV_SERVER_SIDE_FEE") or 0.07)

REVENUE_COL_INDEX     = 2   # 0-based: Domain=0, Date=1, Revenue=2
REVENUE_NUMBER_FORMAT = '"$"#,##0.00####'

MONTH_START_GRACE_DAYS = 5


def _empty_data_exit(reason: str) -> None:
    day = datetime.now().day
    if day <= MONTH_START_GRACE_DAYS:
        log(f"No current-month data yet ({reason}); within first "
            f"{MONTH_START_GRACE_DAYS} days of the month → treating as normal reporting "
            "lag, not a failure. Leaving existing sheet data in place. Skipping this run.")
        sys.exit(0)
    sys.exit(f"ERROR: {reason} — aborting.")


def _fmt_money(v) -> str:
    """Up to 6 decimals, trailing zeros trimmed, min 2 dp — keeps sub-cent (and
    ×0.93-fee'd) values lossless when summed, while staying readable."""
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


def _normalize_and_aggregate(rows, fee=0.0):
    """
    Collapse URL paths/subdomains to the registered root domain, then sum
    Impression + Revenue per (Domain, Date). Revenue is taken net of `fee` (0 for the
    client sheet, SERVER_SIDE_FEE for the S2S sheet), and CPM is recomputed from the
    net revenue so it carries the fee too. App bundle ids (no dot) pass through.
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
        rev *= (1 - fee)
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


_REQUIRED_KEYS = ("equativ_client_id", "equativ_client_secret")


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
        sys.exit(f"ERROR: Missing credentials {missing}. "
                 f"Set env vars [{env_names}] or fill config.json.")
    return cfg


def load_service_account():
    """Service account: env var GOOGLE_SERVICE_ACCOUNT_JSON first, then file."""
    sa_json_str = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if sa_json_str:
        try:
            sa_info = json.loads(sa_json_str)
            return service_account.Credentials.from_service_account_info(sa_info, scopes=SCOPES)
        except Exception as e:
            sys.exit(f"ERROR: GOOGLE_SERVICE_ACCOUNT_JSON env var invalid: {e}")
    if not SA_PATH.exists():
        sys.exit(f"ERROR: {SA_PATH} not found and GOOGLE_SERVICE_ACCOUNT_JSON not set.")
    sa = json.loads(SA_PATH.read_text())
    if not sa.get("client_email") or not sa.get("private_key"):
        sys.exit("ERROR: service_account.json is not populated.")
    return service_account.Credentials.from_service_account_file(str(SA_PATH), scopes=SCOPES)


# ── Step 1 — Fetch the report from the Equativ Reporting API ──────────────────

def _get_token(cfg: dict) -> str:
    log("Authenticating with Equativ (OAuth2 client_credentials)…")
    r = requests.post(TOKEN_URL,
                      data={"grant_type": "client_credentials",
                            "client_id": cfg["equativ_client_id"],
                            "client_secret": cfg["equativ_client_secret"]},
                      headers={"Content-Type": "application/x-www-form-urlencoded"},
                      timeout=30)
    if r.status_code != 200:
        sys.exit(f"ERROR: Equativ token request failed ({r.status_code}): {r.text[:300]} — "
                 "check the client id/secret. Sheet left unchanged.")
    tok = r.json().get("access_token")
    if not tok:
        sys.exit(f"ERROR: no access_token in token response: {str(r.json())[:200]}")
    return tok


def fetch_report(cfg: dict) -> list:
    """Submit the MTD async report, poll to completion, download + parse the CSV."""
    token = _get_token(cfg)
    H = {"Authorization": f"Bearer {token}"}

    today = date.today()
    first = today.replace(day=1)
    start = min(first, today - timedelta(days=LOOKBACK_DAYS))
    end   = today + timedelta(days=1)   # endDate = next day 00:00 → full current day
    body = {
        "reportName": "mm_daily_split",
        "startDate": f"{start.isoformat()}T00:00:00",
        "endDate":   f"{end.isoformat()}T00:00:00",
        "dimensions": ["Day", "AppOrSiteDomain", "ServerSideBiddingCallerName"],
        "metrics":    ["Impressions", "EstimatedPublisherNetRevenuePublisherCurrency"],
        "useCaseId":  "Holistic",
        "timezone":   "UTC",
        "dateFormat": "yyyy-MM-dd",
    }
    log(f"Submitting MTD report {start.isoformat()} → {today.isoformat()} "
        "(Day + domain + ServerSideBiddingCaller)…")
    r = requests.post(REPORT_URL, json=body, headers=H, timeout=60)
    if r.status_code != 200:
        sys.exit(f"ERROR: report submission failed ({r.status_code}): {r.text[:300]}")
    report_id = r.text.strip().strip('"')
    if not report_id:
        sys.exit(f"ERROR: no report id returned: {r.text[:200]}")
    log(f"Report queued: {report_id}. Polling…")

    deadline = time.time() + POLL_MAX_S
    download_url = None
    while time.time() < deadline:
        g = requests.get(f"{REPORT_URL}/{report_id}", headers=H, timeout=30)
        if g.status_code != 200:
            sys.exit(f"ERROR: report status check failed ({g.status_code}): {g.text[:200]}")
        j = g.json()
        st = j.get("status")
        if st == STATUS_SUCCESS:
            download_url = j.get("instanceId")
            break
        if st in (STATUS_FAILED, STATUS_DELETED):
            sys.exit(f"ERROR: Equativ report failed (status={st}): {json.dumps(j)[:300]}")
        log(f"  status={st} (in progress); waiting {POLL_EVERY}s…")
        time.sleep(POLL_EVERY)
    if not download_url:
        sys.exit(f"ERROR: report did not finish within {POLL_MAX_S}s — aborting.")

    log("Downloading report CSV…")
    raw = requests.get(download_url, timeout=90).content
    try:
        text = gzip.decompress(raw).decode("utf-8", "replace")
    except Exception:
        text = raw.decode("utf-8", "replace")
    rows = list(csv.reader(io.StringIO(text), delimiter=";"))
    if not rows:
        _empty_data_exit("downloaded report was empty")
    hdr = [h.strip() for h in rows[0]]
    log(f"CSV header: {hdr} ({len(rows)-1} data rows)")
    records = [dict(zip(hdr, r)) for r in rows[1:] if len(r) == len(hdr)]
    return records


# ── Step 2 — Validate, split, shape ───────────────────────────────────────────

def _col(df, *candidates):
    norm = {c.strip().lower(): c for c in df.columns}
    for cand in candidates:
        if cand.strip().lower() in norm:
            return norm[cand.strip().lower()]
    return None


def process_and_split(records: list):
    """Return (server_df, client_df) shaped to Date/Domain/Impressions/Revenue."""
    log("Shaping + splitting report rows (server vs client)…")
    if not records:
        _empty_data_exit("API returned no rows for the current month")

    df = pd.DataFrame(records)
    date_col = _col(df, "Day", "Date")
    dom_col  = _col(df, "App Or Site Domain", "AppOrSiteDomain", "Domain")
    caller_col = _col(df, "Server Side Bidding Caller Name", "ServerSideBiddingCallerName")
    imp_col  = _col(df, "Impressions")
    rev_col  = _col(df, "Estimated Net Revenue", "EstimatedPublisherNetRevenuePublisherCurrency")
    missing = [n for n, c in [("Day", date_col), ("AppOrSiteDomain", dom_col),
                              ("ServerSideBiddingCallerName", caller_col), ("Impressions", imp_col),
                              ("Estimated Net Revenue", rev_col)] if c is None]
    if missing:
        sys.exit(f"ERROR: Could not find columns {missing} in the report.\n"
                 f"       Found: {list(df.columns)} — aborting to protect the sheets.")

    df = df.dropna(subset=[dom_col, date_col])
    df["__date"] = pd.to_datetime(df[date_col], errors="coerce")
    df = df[df["__date"].notna()].copy()
    if df.empty:
        _empty_data_exit("no valid rows after date parsing")

    # Freshness guard (current-month-tolerant).
    newest = df["__date"].max()
    age_days = (datetime.now() - newest).days
    newest_in_month = newest.strftime("%Y-%m") == datetime.now().strftime("%Y-%m")
    if age_days > MAX_ALLOWED_AGE_DAYS and not newest_in_month:
        sys.exit(f"ERROR: Newest date is {newest.date()} ({age_days} days ago) and is "
                 "outside the current month — looks stuck/broken, aborting.")

    # Split by the server-side-bidding caller:
    #   PrebidServer → S2S (7% fee) · N/A → client (no fee) · everything else
    #   (Google = open bidding, and any other caller) → IGNORED, in neither sheet.
    caller = df[caller_col].astype(str).str.strip().str.lower()
    is_server = caller.str.contains("prebid")
    is_client = caller.isin(["n/a", "na", ""])
    ignored = ~(is_server | is_client)
    if ignored.any():
        from collections import Counter
        vals = Counter(df.loc[ignored, caller_col].astype(str))
        log(f"Ignoring {int(ignored.sum())} row(s) by caller (open bidding / other, in "
            f"NEITHER sheet): {dict(vals)}")

    def shape(mask):
        d = df[mask]
        if d.empty:
            return pd.DataFrame(columns=["Date", "Domain", "Impressions", "Revenue"])
        out = pd.DataFrame({
            "Date":        d["__date"].dt.strftime("%Y-%m-%d"),
            "Domain":      d[dom_col].astype(str).str.strip(),
            "Impressions": pd.to_numeric(d[imp_col], errors="coerce").fillna(0),
            "Revenue":     pd.to_numeric(d[rev_col], errors="coerce").fillna(0),
        })
        out = out[~out["Domain"].str.lower().isin(["", "nan", "none"])]
        out = out[(pd.to_numeric(out["Impressions"], errors="coerce").fillna(0) > 0)
                  | (pd.to_numeric(out["Revenue"], errors="coerce").fillna(0) > 0)]
        return out

    server_df = shape(is_server)        # PrebidServer
    client_df = shape(is_client)        # N/A  (Google/other dropped entirely)
    log(f"Split: {len(server_df)} server-side (PrebidServer) rows, "
        f"{len(client_df)} client-side (N/A) rows.")
    if server_df.empty and client_df.empty:
        _empty_data_exit("no rows with activity after split")
    return server_df, client_df


# ── Step 3 — Write a sheet (preserve previous-month history) ──────────────────

def write_sheet(df: pd.DataFrame, creds, spreadsheet_id: str, fee: float, label: str) -> None:
    if df.empty:
        log(f"[{label}] no fresh rows — leaving sheet untouched.")
        return
    log(f"[{label}] Connecting to Google Sheets…")
    service = build("sheets", "v4", credentials=creds, cache_discovery=False)

    _sheet_id = 0
    try:
        _meta = service.spreadsheets().get(
            spreadsheetId=spreadsheet_id, fields="sheets.properties(title,sheetId)").execute()
        _props = _meta["sheets"][0]["properties"]
        _tab = _props["title"]; _sheet_id = _props["sheetId"]
    except Exception:
        _tab = SHEET_NAME
    log(f"[{label}] Using sheet tab: {_tab!r}")

    rows = [
        [str(r["Domain"]).strip(), str(r["Date"]).strip(),
         _fmt_money(pd.to_numeric(r["Revenue"], errors="coerce") or 0),
         int(pd.to_numeric(r["Impressions"], errors="coerce") or 0), 0.0]
        for _, r in df.iterrows()
    ]
    rows = _normalize_and_aggregate(rows, fee=fee)
    rows.sort(key=lambda r: (r[0], r[1]))
    rows.sort(key=lambda r: r[1], reverse=True)

    existing = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=_tab).execute().get("values", [])
    current_month = datetime.now().strftime("%Y-%m")
    preserved = []
    for _r in existing:
        if not _r or _r == HEADER:
            continue
        if len(_r) >= 2 and str(_r[1]).strip().startswith(current_month):
            continue
        preserved.append(_r)

    merged = rows + preserved
    _seen, _deduped = set(), []
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
    log(f"[{label}] Writing {len(merged)} rows = {len(rows)} new MTD + {len(preserved)} preserved…")
    service.spreadsheets().values().clear(spreadsheetId=spreadsheet_id, range=_tab, body={}).execute()
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=f"{_tab}!A1",
        valueInputOption="USER_ENTERED", body={"values": final_rows}).execute()

    try:
        fmt_rows = [{"values": [{"userEnteredFormat": {"numberFormat": {
            "type": "NUMBER", "pattern": REVENUE_NUMBER_FORMAT}}}]} for _ in range(len(merged))]
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": [{
            "updateCells": {"start": {"sheetId": _sheet_id, "rowIndex": 1, "columnIndex": REVENUE_COL_INDEX},
                            "rows": fmt_rows, "fields": "userEnteredFormat.numberFormat"}}]}).execute()
    except Exception as e:
        log(f"[{label}] WARNING: could not stamp revenue number format (display only): {e}")
    log(f"[{label}] Sheet updated successfully.")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    log("=== Equativ → Google Sheets daily sync (client + server split) ===")
    cfg   = load_config()
    creds = load_service_account()

    records = fetch_report(cfg)
    server_df, client_df = process_and_split(records)
    write_sheet(server_df, creds, SERVER_SPREADSHEET_ID, SERVER_SIDE_FEE, "SERVER/S2S")
    write_sheet(client_df, creds, CLIENT_SPREADSHEET_ID, 0.0, "CLIENT")
    log("=== Done. ===")


if __name__ == "__main__":
    main()
