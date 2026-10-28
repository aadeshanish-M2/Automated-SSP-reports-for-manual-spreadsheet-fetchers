#!/usr/bin/env python3
"""
Rich Audience (platform.richaudience.com) → Google Sheets daily sync.

Pure API integration — no browser. Authenticates against the platform's standard
login form with `requests` (GET the login page for the CSRF token, POST the
credentials to get a PHP session), then calls the Global Report's own JSON
endpoint (/report/global/pbd/get-report) that the dashboard itself uses, and
syncs to the destination sheet in the uniform format
(Domain | Date | Revenue | Impression | CPM). Previous-month history in the sheet
is preserved; only current-month rows are refreshed.

This replaced a brittle Playwright flow (the reporting UI kept changing: menus,
selectize date presets, a new 2FA security modal, a required "Run report" click).
Driving the JSON API directly removes all of that UI surface.

NOTE: Rich Audience announced it will enforce two-factor auth (an emailed one-time
code) at login. When that lands, password-only sign-in stops working and the login
step here must be extended to supply the OTP — fetch_report() detects the challenge
and exits with a clear message rather than failing cryptically.
"""

import email
import imaplib
import json
import os
import re
import sys
import time
from datetime import date, datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path

import pandas as pd
import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ── Configuration ─────────────────────────────────────────────────────────────

SCRIPT_DIR     = Path(__file__).parent
CONFIG_PATH    = SCRIPT_DIR / "config.json"
SA_PATH        = SCRIPT_DIR / "service_account.json"
SPREADSHEET_ID = os.environ.get("SHEETS_SPREADSHEET_ID") or "1xDRYSjZm9hO5tAnAVxihTp73XPQOvabsa7jf2lQXLIw"
SHEET_NAME     = "Sheet1"
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]

LOGIN_PAGE     = "https://platform.richaudience.com/login"
REPORT_PAGE    = "https://platform.richaudience.com/report/global"
GET_REPORT_URL = "https://platform.richaudience.com/report/global/pbd/get-report"
# Rolling window pulled from the API. write_sheet still keeps only current-month
# rows fresh and preserves older history, so a 30-day pull spans the month boundary
# without ever overwriting prior-month data.
REPORT_DAYS    = 30

# ── Two-factor auth (email OTP) ───────────────────────────────────────────────
# Rich Audience is rolling out mandatory 2FA: at login it emails a one-time code
# to the account's address (addemand@monetizemore.com) which must be submitted to
# finish signing in. We read that code over IMAP. These credentials are OPTIONAL —
# only needed once 2FA is active; until then login is password-only and this is
# unused. The email arrives FROM info@richaudience.com (same sender as RA's other
# mail). App-password IMAP; creds live in config.json / Streamlit secrets as
# richaudience_otp_imap_user / richaudience_otp_imap_password.
IMAP_HOST           = "imap.gmail.com"
IMAP_PORT           = 993
OTP_SENDER          = "info@richaudience.com"
OTP_EMAIL_WAIT_SECS = 150   # poll the inbox up to this long for the code to arrive
OTP_POLL_EVERY      = 6

# Rich Audience reports low-volume accounts with a genuine multi-day lag — the
# "Current month" range is requested correctly (verified: Start=1st, End=today),
# but their data backfills several days behind. Observed ~7-day lag, so 5 was too
# tight and blocked valid month-to-date data from being written. The MTD filter
# already drops any last-month rows, so this guard only needs to catch a truly
# stuck/broken fetch (data weeks old).
MAX_ALLOWED_AGE_DAYS = 14
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


# At the start of a month, SSPs haven't reported the new month's data yet
# (reporting lag — Rich Audience runs several days behind on low-volume accounts),
# so an empty current-month pull is EXPECTED, not a failure. Within this window we
# exit successfully WITHOUT touching the sheet: the previous month's data already
# preserved there keeps showing on the dashboard until new data starts flowing.
# Outside the window, an empty pull means something genuinely broke → hard error.
MONTH_START_GRACE_DAYS = 10


def _empty_data_exit(reason: str) -> None:
    day = datetime.now().day
    if day <= MONTH_START_GRACE_DAYS:
        log(f"No current-month data yet ({reason}); within first "
            f"{MONTH_START_GRACE_DAYS} days of the month → treating as normal "
            "reporting lag, not a failure. Leaving existing sheet data in place. "
            "Skipping this run.")
        sys.exit(0)
    sys.exit(f"ERROR: {reason} — aborting.")


_REQUIRED_KEYS = ("richaudience_username", "richaudience_password")
# Optional — only consumed when 2FA is active (see the OTP section above).
_OPTIONAL_KEYS = ("richaudience_otp_imap_user", "richaudience_otp_imap_password")


def load_config() -> dict:
    """Credentials: env vars first (UPPER_CASE), then config.json fallback."""
    cfg = {}
    all_keys = _REQUIRED_KEYS + _OPTIONAL_KEYS
    for k in all_keys:
        v = os.environ.get(k.upper())
        if v:
            cfg[k] = v
    if CONFIG_PATH.exists():
        try:
            file_cfg = json.loads(CONFIG_PATH.read_text())
            for k in all_keys:
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


# ── Step 1 — Fetch the report via the platform's JSON API ─────────────────────

def _new_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/124.0.0.0 Safari/537.36"),
        "Accept": "text/html,application/json,*/*",
    })
    return s


def _extract_csrf(html: str):
    """Yii2 renders the CSRF token both as a <meta> tag and a hidden form input."""
    for pat in (
        r'<meta name="csrf-token" content="([^"]+)"',
        r'name="_csrf"[^>]*value="([^"]+)"',
        r'value="([^"]+)"[^>]*name="_csrf"',
    ):
        m = re.search(pat, html)
        if m:
            return m.group(1)
    return None


def _looks_like_otp_challenge(text: str) -> bool:
    t = (text or "").lower()
    return any(n in t for n in (
        "one-time password", "one time password", "verification code",
        "two-factor", "two factor", "2fa", "authentication code",
        "enter the code", "código", "otp",
    ))


def _request_with_retry(sess, method, url, **kw):
    """GET/POST with exponential backoff on 5xx and connection errors."""
    attempts = 4
    kw.setdefault("timeout", 60)
    r = None
    for i in range(1, attempts + 1):
        try:
            r = sess.request(method, url, **kw)
        except requests.RequestException as e:
            if i == attempts:
                sys.exit(f"ERROR: Rich Audience request failed (connection error: {e}). "
                         "Likely a temporary network/platform issue; sheet left unchanged.")
            log(f"  attempt {i}/{attempts}: connection error ({e}); retrying…")
            time.sleep(min(2 ** i, 20))
            continue
        if r.status_code < 500:
            return r
        if i == attempts:
            sys.exit(f"ERROR: Rich Audience returned HTTP {r.status_code} after {attempts} "
                     "attempts — a platform-side server error, not the sync. Sheet left unchanged.")
        log(f"  attempt {i}/{attempts}: HTTP {r.status_code}; retrying…")
        time.sleep(min(2 ** i, 20))
    return r


def _email_subject(msg) -> str:
    from email.header import make_header, decode_header
    try:
        return str(make_header(decode_header(msg.get("Subject", "") or "")))
    except Exception:
        return msg.get("Subject", "") or ""


def _email_text(msg) -> str:
    """Return the plain-text body (HTML stripped as a fallback)."""
    try:
        if msg.is_multipart():
            for want in ("text/plain", "text/html"):
                for part in msg.walk():
                    if part.get_content_type() == want:
                        raw = part.get_payload(decode=True) or b""
                        txt = raw.decode(part.get_content_charset() or "utf-8", "replace")
                        return re.sub("<[^>]+>", " ", txt) if want == "text/html" else txt
            return ""
        raw = msg.get_payload(decode=True)
        return (raw.decode(msg.get_content_charset() or "utf-8", "replace")
                if raw else (msg.get_payload() or ""))
    except Exception:
        return ""


def _extract_code(subject: str, body: str) -> str:
    """Pull the OTP from the email: prefer a number next to a keyword, then a
    6-digit number, then any 4–8 digit number."""
    hay = f"{subject}\n{body}"
    m = re.search(r'(?:code|c[oó]digo|otp|password|contrase[nñ]a|pin)\D{0,20}(\d{4,8})', hay, re.I)
    if m:
        return m.group(1)
    m = re.search(r'(?<!\d)(\d{6})(?!\d)', hay)
    if m:
        return m.group(1)
    m = re.search(r'(?<!\d)(\d{4,8})(?!\d)', hay)
    return m.group(1) if m else ""


def _read_otp_from_email(imap_user: str, imap_pw: str, since_dt: datetime) -> str:
    """Poll the mailbox for the newest Rich Audience OTP email received around/after
    `since_dt` and return its code. Raises RuntimeError on timeout."""
    deadline = time.time() + OTP_EMAIL_WAIT_SECS
    since_imap = (since_dt - timedelta(days=1)).strftime("%d-%b-%Y")  # IMAP date is day-granular
    cutoff = since_dt - timedelta(minutes=2)
    last_err = None
    while time.time() < deadline:
        try:
            M = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
            M.login(imap_user, imap_pw)
            M.select("INBOX")
            typ, ids = M.search(None, 'FROM', f'"{OTP_SENDER}"', 'SINCE', since_imap)
            id_list = ids[0].split() if ids and ids[0] else []
            for num in reversed(id_list):          # newest first
                typ, d = M.fetch(num, "(RFC822)")
                if not d or not d[0]:
                    continue
                msg = email.message_from_bytes(d[0][1])
                try:
                    recv = parsedate_to_datetime(msg.get("Date"))
                    if recv and recv.tzinfo is None:
                        recv = recv.replace(tzinfo=timezone.utc)
                    if recv and recv < cutoff:     # older than this login attempt
                        break
                except Exception:
                    pass
                code = _extract_code(_email_subject(msg), _email_text(msg))
                if code:
                    try: M.logout()
                    except Exception: pass
                    return code
            try: M.logout()
            except Exception: pass
        except Exception as e:
            last_err = e
        time.sleep(OTP_POLL_EVERY)
    raise RuntimeError(
        f"no OTP email from {OTP_SENDER} arrived within {OTP_EMAIL_WAIT_SECS}s"
        + (f" (last IMAP error: {last_err})" if last_err else "")
    )


def _submit_otp(sess, challenge_html: str, challenge_url: str, otp: str) -> None:
    """Adaptively submit the OTP into whatever challenge form RA presents: parse the
    form action + all inputs (incl. _csrf), discover the code field by name/type,
    fill it, and POST. Field names are read from the page, not hardcoded, because the
    exact 2FA form isn't visible until RA turns it on — so this is best-effort and
    must be verified once 2FA is live."""
    from urllib.parse import urljoin
    forms = re.findall(r'<form\b[^>]*>.*?</form>', challenge_html, re.I | re.S)
    form = challenge_html
    for f in forms:
        if re.search(r'LoginForm\[email\]', f):   # skip the plain login form
            continue
        form = f
        if re.search(r'(code|otp|token|pin|verif|auth|two|factor|2fa)', f, re.I):
            break

    am = re.search(r'<form\b[^>]*\baction="([^"]*)"', form, re.I)
    action = urljoin(challenge_url, am.group(1)) if am and am.group(1) else challenge_url

    fields, otp_field = {}, None
    input_tags = re.findall(r'<input\b[^>]*>', form, re.I)
    for tag in input_tags:
        nm = re.search(r'\bname="([^"]+)"', tag, re.I)
        if not nm:
            continue
        name = nm.group(1)
        val = re.search(r'\bvalue="([^"]*)"', tag, re.I)
        itype = re.search(r'\btype="([^"]*)"', tag, re.I)
        itype = (itype.group(1).lower() if itype else "text")
        fields[name] = val.group(1) if val else ""
        if (otp_field is None and itype in ("text", "tel", "number", "password", "")
                and re.search(r'(code|otp|token|pin|verif|auth|digit|factor)', tag, re.I)):
            otp_field = name
    if otp_field is None:                          # fall back: first text-ish input
        for tag in input_tags:
            itype = re.search(r'\btype="([^"]*)"', tag, re.I)
            itype = (itype.group(1).lower() if itype else "text")
            nm = re.search(r'\bname="([^"]+)"', tag, re.I)
            if nm and itype in ("text", "tel", "number", ""):
                otp_field = nm.group(1)
                break
    if otp_field is None:
        sys.exit("ERROR: 2FA challenge detected but the code input field could not be "
                 "located on the page — the live challenge needs inspecting to finish "
                 "wiring. Sheet left unchanged.")

    fields[otp_field] = otp
    log(f"Submitting OTP into field {otp_field!r} → {action}")
    _request_with_retry(sess, "POST", action, data=fields,
                        headers={"Referer": challenge_url}, allow_redirects=True)


def fetch_report(username: str, password: str,
                 imap_user: str = None, imap_pw: str = None) -> list:
    """Log in (handling 2FA email-OTP if required) and pull the Global Report JSON
    for a rolling REPORT_DAYS window."""
    sess = _new_session()

    log("Authenticating with Rich Audience…")
    r = _request_with_retry(sess, "GET", LOGIN_PAGE)
    csrf = _extract_csrf(r.text)
    if not csrf:
        sys.exit("ERROR: could not find the CSRF token on the login page — the login "
                 "page structure may have changed. Sheet left unchanged.")
    login_at = datetime.now(timezone.utc)
    login_resp = _request_with_retry(sess, "POST", LOGIN_PAGE,
                        data={
                            "_csrf": csrf,
                            "LoginForm[email]": username,
                            "LoginForm[password]": password,
                            "login-button": "",
                        },
                        headers={"Referer": LOGIN_PAGE},
                        allow_redirects=True)

    today = date.today()
    start = today - timedelta(days=REPORT_DAYS)
    params = [
        ("startDate", start.isoformat()),
        ("endDate", today.isoformat()),
        ("dimensions[]", "site"),
        ("dimensions[]", "date"),
        ("currency", "usd"),
        ("openTrans", "0"),
    ]

    def _get_report():
        return _request_with_retry(sess, "GET", GET_REPORT_URL, params=params,
                                   headers={"X-Requested-With": "XMLHttpRequest",
                                            "Accept": "application/json",
                                            "Referer": REPORT_PAGE})

    log(f"Requesting Global Report JSON (last {REPORT_DAYS} days, dims=site+date)…")
    rr = _get_report()

    # If we're not authenticated the API returns HTML (login/OTP page), not JSON.
    if "application/json" not in rr.headers.get("content-type", ""):
        on_otp = _looks_like_otp_challenge(login_resp.text)
        if on_otp or _looks_like_otp_challenge(rr.text):
            log("Two-factor challenge detected — retrieving the emailed OTP…")
            if not (imap_user and imap_pw):
                sys.exit(
                    "ERROR: Rich Audience now requires two-factor authentication (an emailed "
                    "one-time code), but no OTP mailbox is configured. Set "
                    "richaudience_otp_imap_user / richaudience_otp_imap_password. "
                    "Sheet left unchanged."
                )
            try:
                otp = _read_otp_from_email(imap_user, imap_pw, login_at)
            except Exception as e:
                sys.exit(f"ERROR: could not retrieve the Rich Audience OTP from email: {e}. "
                         "Sheet left unchanged.")
            log(f"OTP retrieved ({'*' * max(len(otp) - 1, 0)}{otp[-1:]}); submitting…")
            challenge_html = login_resp.text if on_otp else rr.text
            challenge_url  = login_resp.url if on_otp else GET_REPORT_URL
            _submit_otp(sess, challenge_html, challenge_url, otp)
            rr = _get_report()
            if "application/json" not in rr.headers.get("content-type", ""):
                sys.exit(
                    "ERROR: still not authenticated after submitting the OTP — the live 2FA "
                    "challenge form likely differs from what the adaptive submitter expected "
                    "and needs inspecting. Sheet left unchanged."
                )
        else:
            sys.exit(
                "ERROR: Rich Audience login did not succeed — the report API returned HTML "
                "instead of JSON (most likely invalid credentials or a changed login flow). "
                "The sheet was left unchanged."
            )
    try:
        j = rr.json()
    except ValueError:
        sys.exit(f"ERROR: Rich Audience report was not valid JSON: {rr.text[:200]}")
    if j.get("status") != "success":
        sys.exit(f"ERROR: Rich Audience report returned status={j.get('status')!r}, "
                 f"message={j.get('message')!r}.")

    data = j.get("data") or []
    log(f"Rich Audience returned {len(data)} rows ({start.isoformat()} → {today.isoformat()}).")
    return data


# ── Step 2 — Validate & shape ─────────────────────────────────────────────────

def process_report(data: list) -> pd.DataFrame:
    log("Shaping report rows…")
    if not data:
        _empty_data_exit("API returned no rows for the requested window")

    df = pd.DataFrame(data)
    # JSON field names from /report/global/pbd/get-report.
    date_col, site_col = "date", "site"
    impr_col, rev_col, ecpm_col = "impressions", "revenue", "ecpm"

    missing = [k for k in (date_col, site_col, impr_col, rev_col, ecpm_col)
               if k not in df.columns]
    if missing:
        sys.exit(
            f"ERROR: Could not find fields {missing} in the API response.\n"
            f"       Got fields: {list(df.columns)}\n"
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
        _empty_data_exit("no valid rows after date parsing")

    newest = df["__date"].max()
    age_days = (datetime.now() - newest).days
    # The age guard exists to catch a fetch STUCK on old/previous-month data.
    # But Rich Audience is a very low-volume account that regularly goes quiet
    # for stretches (verified: zero activity for weeks at a time), so its newest
    # date can be legitimately old while still being valid current-month data.
    # Only hard-fail when the newest date is NOT in the current month (a genuine
    # stuck/wrong-month signal); otherwise accept the low-volume data with a
    # warning rather than failing the sync every day.
    newest_in_current_month = newest.strftime("%Y-%m") == datetime.now().strftime("%Y-%m")
    if age_days > MAX_ALLOWED_AGE_DAYS and not newest_in_current_month:
        sys.exit(
            f"ERROR: Newest date is {newest.date()} ({age_days} days ago) and is "
            f"outside the current month — looks stuck/broken, aborting."
        )
    if age_days > MAX_ALLOWED_AGE_DAYS:
        log(f"WARNING: newest date {newest.date()} is {age_days} days old but is "
            "within the current month — Rich Audience has no recent activity "
            "(low-volume account). Writing the available data instead of failing.")

    # Keep the FULL pulled window (it spans a few days into the previous month near a
    # boundary) — do NOT MTD-filter. The write refreshes exactly the pulled dates via
    # its fresh-first dedup and preserves older history, so the final day of the
    # previous month gets FINALIZED instead of frozen partial.
    log(f"Keeping full pulled window: {len(df)} rows, "
        f"{df['__date'].min().date()} → {df['__date'].max().date()}.")

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
    log("=== Rich Audience → Google Sheets daily sync ===")
    cfg   = load_config()
    creds = load_service_account()

    data = fetch_report(cfg["richaudience_username"], cfg["richaudience_password"],
                        cfg.get("richaudience_otp_imap_user"),
                        cfg.get("richaudience_otp_imap_password"))
    df   = process_report(data)
    write_sheet(df, creds)
    log("=== Done. ===")


if __name__ == "__main__":
    main()
