"""
SSP Revenue Reports — local web dashboard.  Matrix edition.

Run with:
    .venv/bin/streamlit run app.py
"""

from __future__ import annotations

import os
import random
import string
import subprocess
import sys
from pathlib import Path

import streamlit as st

from runner import (
    SSPS,
    run_one,
    load_history,
    save_run,
    load_overrides,
    save_overrides,
    extract_sheet_id,
)


# ── First-boot: install Playwright Chromium if missing ────────────────────────
# Streamlit Cloud doesn't run a setup script after pip install, so we lazily
# install the Chromium binary the first time the app boots. ~30s one-time cost
# per fresh container; subsequent reboots see the cached binary.
@st.cache_resource
def _ensure_chromium_installed() -> None:
    cache_dir = Path(os.environ.get("PLAYWRIGHT_BROWSERS_PATH",
                                    str(Path.home() / ".cache" / "ms-playwright")))
    if cache_dir.exists() and any(cache_dir.glob("chromium*-*")):
        return
    try:
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            check=True,
            capture_output=True,
            timeout=300,
        )
    except Exception as e:
        # Don't crash the dashboard — the SSPs that need a browser will fail
        # with a clearer error of their own.
        print(f"Warning: could not install Playwright Chromium: {e}")


_ensure_chromium_installed()

st.set_page_config(
    page_title="SSP Revenue Reports",
    page_icon="◉",
    layout="centered",
    initial_sidebar_state="collapsed",
)


# ── Optional password gate (active when SHARE_PASSWORD env var is set) ────────
_SHARE_PASSWORD = os.environ.get("SHARE_PASSWORD", "").strip()
if _SHARE_PASSWORD:
    if "authed" not in st.session_state:
        st.session_state.authed = False
    if not st.session_state.authed:
        st.markdown(
            """
            <style>
              html, body, .stApp { background:#000 !important; color:#00ff41 !important;
                font-family:'Courier New', monospace !important; }
              h2 { color:#00ff41 !important; text-align:center;
                text-shadow:0 0 10px rgba(0,255,65,0.6);
                font-family:'Courier New', monospace !important; }
              .stTextInput input { background:rgba(0,20,5,0.7) !important;
                color:#00ff41 !important; border:1px solid #00ff41 !important;
                font-family:'Courier New', monospace !important; }
              .stButton button, .stFormSubmitButton button {
                background:transparent !important; color:#00ff41 !important;
                border:1px solid #00ff41 !important;
                font-family:'Courier New', monospace !important;
                text-transform:uppercase; letter-spacing:2px; }
            </style>
            <h2>◉ Access required</h2>
            """,
            unsafe_allow_html=True,
        )
        with st.form("login"):
            pwd = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Enter")
            if submitted:
                if pwd == _SHARE_PASSWORD:
                    st.session_state.authed = True
                    st.rerun()
                else:
                    st.error("Wrong password.")
        st.stop()

# ── Matrix rain background ────────────────────────────────────────────────────

KATAKANA = (
    "アイウエオカキクケコサシスセソタチツテトナニヌネノ"
    "ハヒフヘホマミムメモヤユヨラリルレロワヲン"
)
MATRIX_CHARS = KATAKANA + string.ascii_uppercase + string.digits + "@#$%&*+=<>"


def matrix_rain(num_columns: int = 18, chars_per_column: int = 30) -> str:
    cols = []
    for i in range(num_columns):
        chars_html = "<br>".join(random.choice(MATRIX_CHARS) for _ in range(chars_per_column))
        left      = (100 / num_columns) * i + random.uniform(-0.8, 0.8)
        duration  = random.uniform(8, 18)
        delay     = -random.uniform(0, duration)
        # Dim enough that the foreground text still reads cleanly, but visible.
        opacity   = random.uniform(0.18, 0.32)
        font_size = random.choice([13, 14, 16])
        cols.append(
            f'<div class="m-col" style="'
            f'left:{left:.1f}%;'
            f'font-size:{font_size}px;'
            f'opacity:{opacity:.2f};'
            f'animation-duration:{duration:.1f}s;'
            f'animation-delay:{delay:.1f}s;'
            f'">{chars_html}</div>'
        )
    return f'<div class="matrix-bg">{"".join(cols)}</div>'


# ── Theme CSS ─────────────────────────────────────────────────────────────────

st.markdown(
    """
    <style>
      :root {
        --matrix-green:  #00ff41;
        --matrix-dim:    #008f11;
        --matrix-glow:   rgba(0, 255, 65, 0.55);
        --matrix-bright: #d8ffe0;
        --matrix-fail:   #ff003c;
        --matrix-warn:   #fbbf24;
        --matrix-bg:     #000000;
      }

      html, body, .stApp {
        background: var(--matrix-bg) !important;
        color: var(--matrix-green) !important;
        font-family: 'JetBrains Mono', 'Courier New', monospace !important;
      }

      /* Matrix rain layer — sits BEHIND everything */
      .matrix-bg {
        position: fixed;
        top: 0; left: 0;
        width: 100vw;
        height: 100vh;
        overflow: hidden;
        z-index: 0;
        pointer-events: none;
        color: var(--matrix-green);
        font-family: 'Courier New', monospace;
        text-shadow: 0 0 4px var(--matrix-glow);
      }
      .matrix-bg .m-col {
        position: absolute;
        top: -150%;
        line-height: 1.05;
        animation-name: matrixfall;
        animation-timing-function: linear;
        animation-iteration-count: infinite;
        text-align: center;
        width: 1.4em;
      }
      @keyframes matrixfall {
        from { transform: translateY(0); }
        to   { transform: translateY(250vh); }
      }

      /* Make sure all main content sits ABOVE the rain */
      [data-testid="stAppViewContainer"],
      .block-container,
      [data-testid="stHeader"] {
        position: relative;
        z-index: 1;
        background: transparent !important;
      }

      /* Content darkening behind text so rain doesn't fight readability */
      .block-container {
        padding-top: 3rem !important;
        padding-bottom: 4rem !important;
        max-width: 900px !important;
        background:
          radial-gradient(ellipse at top, rgba(0,15,3,0.96), rgba(0,0,0,0.98) 70%) !important;
        border: 1px solid rgba(0, 255, 65, 0.18);
        border-radius: 4px;
        box-shadow:
          0 0 30px rgba(0, 255, 65, 0.05) inset,
          0 0 60px rgba(0, 255, 65, 0.1);
      }

      /* Headings — green phosphor glow */
      h1, h2, h3, h4 {
        color: var(--matrix-green) !important;
        font-family: 'Courier New', monospace !important;
        font-weight: 700 !important;
        letter-spacing: 1px !important;
        text-shadow: 0 0 10px var(--matrix-glow);
        text-transform: uppercase;
      }
      h1 { font-size: 1.9rem !important; }

      p, .stCaption, [data-testid="stCaptionContainer"], label {
        color: var(--matrix-dim) !important;
        font-family: 'Courier New', monospace !important;
        font-size: 0.85rem !important;
      }

      /* Dividers — green ASCII line */
      hr {
        border: none !important;
        height: 1px !important;
        background: repeating-linear-gradient(
          90deg,
          var(--matrix-green) 0 8px,
          transparent 8px 14px
        ) !important;
        opacity: 0.4 !important;
        margin: 1.8rem 0 !important;
      }

      /* RUN ALL — primary button, looks like a terminal command */
      .stButton > button[kind="primary"] {
        background: rgba(0, 30, 5, 0.85) !important;
        color: var(--matrix-green) !important;
        border: 1px solid var(--matrix-green) !important;
        border-radius: 0 !important;
        padding: 0.6rem 1.5rem !important;
        font-family: 'Courier New', monospace !important;
        font-weight: 700 !important;
        letter-spacing: 2px !important;
        text-transform: uppercase;
        box-shadow:
          0 0 8px var(--matrix-glow),
          inset 0 0 6px rgba(0, 255, 65, 0.15) !important;
        transition: all 0.15s ease !important;
      }
      .stButton > button[kind="primary"]:hover {
        background: rgba(0, 80, 15, 0.8) !important;
        color: var(--matrix-bright) !important;
        box-shadow:
          0 0 20px var(--matrix-green),
          inset 0 0 10px rgba(0, 255, 65, 0.3) !important;
      }

      /* Secondary buttons (per-SSP play, save settings) */
      .stButton > button:not([kind="primary"]) {
        background: transparent !important;
        color: var(--matrix-green) !important;
        border: 1px solid var(--matrix-dim) !important;
        border-radius: 0 !important;
        font-family: 'Courier New', monospace !important;
        font-weight: 700 !important;
        text-transform: uppercase;
        letter-spacing: 1px !important;
        transition: all 0.15s ease !important;
      }
      .stButton > button:not([kind="primary"]):hover {
        background: rgba(0, 50, 10, 0.5) !important;
        border-color: var(--matrix-green) !important;
        box-shadow: 0 0 12px var(--matrix-glow) !important;
      }

      /* Text inputs */
      .stTextInput > div > div > input {
        background: rgba(0, 20, 5, 0.7) !important;
        color: var(--matrix-green) !important;
        border: 1px solid var(--matrix-dim) !important;
        border-radius: 0 !important;
        font-family: 'Courier New', monospace !important;
        caret-color: var(--matrix-green);
      }
      .stTextInput > div > div > input:focus {
        border-color: var(--matrix-green) !important;
        box-shadow: 0 0 8px var(--matrix-glow) !important;
      }

      /* Expanders */
      [data-testid="stExpander"] {
        background: rgba(0, 20, 5, 0.45) !important;
        border: 1px solid var(--matrix-dim) !important;
        border-radius: 0 !important;
      }
      [data-testid="stExpander"] summary,
      [data-testid="stExpander"] summary span {
        color: var(--matrix-green) !important;
        font-family: 'Courier New', monospace !important;
        text-transform: uppercase;
        letter-spacing: 1px;
      }
      /* Hide Streamlit's chevron icon (falls back to text "keyboard_arrow_…"
         on machines where the Material Symbols font fails to load) and
         replace with a Unicode arrow we control. */
      [data-testid="stExpander"] summary svg,
      [data-testid="stExpander"] summary [data-testid="stIconMaterial"],
      [data-testid="stExpander"] summary span[class*="icon"]:not(.mx-cursor) {
        display: none !important;
      }
      [data-testid="stExpander"] summary::after {
        content: " ▼";
        color: var(--matrix-green);
        font-size: 0.7em;
        margin-left: 0.5rem;
        transition: transform 0.2s ease;
        display: inline-block;
      }
      [data-testid="stExpander"] details[open] summary::after {
        transform: rotate(180deg);
      }

      /* Alerts (success/error banners) */
      .stAlert {
        background: rgba(0, 30, 5, 0.75) !important;
        border: 1px solid var(--matrix-green) !important;
        border-radius: 0 !important;
        color: var(--matrix-green) !important;
        font-family: 'Courier New', monospace !important;
      }
      .stAlert [data-testid="stAlertContentError"],
      .stAlert[data-baseweb="notification"] {
        color: var(--matrix-green) !important;
      }

      /* Code blocks */
      .stCodeBlock, pre, code {
        background: rgba(0, 10, 0, 0.85) !important;
        color: var(--matrix-green) !important;
        border: 1px solid var(--matrix-dim) !important;
        border-radius: 0 !important;
        font-family: 'Courier New', monospace !important;
      }

      /* Links */
      a {
        color: var(--matrix-bright) !important;
        text-decoration: underline !important;
        text-shadow: 0 0 4px var(--matrix-glow);
      }

      [data-testid="stHeader"] { background: transparent !important; }

      /* ═══════════════ MATRIX ANIMATIONS ═══════════════ */

      /* Blinking cursor */
      @keyframes blink { 50% { opacity: 0; } }
      .mx-cursor {
        display: inline-block;
        animation: blink 1s steps(2) infinite;
        color: var(--matrix-green);
        text-shadow: 0 0 6px var(--matrix-glow);
      }

      /* Glitch flicker on running text */
      @keyframes glitch {
        0%, 100% { text-shadow: 0 0 6px var(--matrix-glow); transform: translateX(0); }
        20%      { text-shadow: 2px 0 #ff003c, -2px 0 var(--matrix-green); transform: translateX(-1px); }
        40%      { text-shadow: -2px 0 #ff003c, 2px 0 var(--matrix-green); transform: translateX(1px); }
        60%      { text-shadow: 0 0 6px var(--matrix-glow); transform: translateX(0); }
      }
      .mx-glitch {
        display: inline-block;
        animation: glitch 1.4s ease-in-out infinite;
        color: var(--matrix-green);
        font-weight: 700;
      }

      /* Decoding text — wave of brightness */
      @keyframes decode {
        0%, 100% { color: var(--matrix-dim);    text-shadow: none; }
        50%      { color: var(--matrix-bright); text-shadow: 0 0 10px var(--matrix-glow); }
      }
      .mx-decode {
        animation: decode 1.6s ease-in-out infinite;
      }

      /* Successful breach — pulse green halo */
      @keyframes breach {
        0%   { text-shadow: 0 0 0 transparent; transform: scale(1); }
        25%  { text-shadow: 0 0 22px var(--matrix-green); transform: scale(1.2); }
        100% { text-shadow: 0 0 8px var(--matrix-glow); transform: scale(1); }
      }
      .mx-breach {
        display: inline-block;
        animation: breach 0.7s ease-out 1;
        color: var(--matrix-green);
        font-weight: 700;
      }

      /* Failed access — red jitter */
      @keyframes denied {
        0%, 100% { transform: translateX(0); color: var(--matrix-fail); }
        20%      { transform: translateX(-3px); }
        40%      { transform: translateX(3px); }
        60%      { transform: translateX(-2px); }
        80%      { transform: translateX(2px); }
      }
      .mx-denied {
        display: inline-block;
        animation: denied 0.4s ease-in-out 1;
        color: var(--matrix-fail);
        font-weight: 700;
        text-shadow: 0 0 8px rgba(255, 0, 60, 0.7);
      }
    </style>
    """,
    unsafe_allow_html=True,
)

# Drop the rain layer (renders once per page load; columns are randomised each time)
st.markdown(matrix_rain(), unsafe_allow_html=True)


# ── Header ────────────────────────────────────────────────────────────────────

_header_img = Path(__file__).parent.parent / "assets" / "matrix-header.png"
if _header_img.exists():
    # Source is 1254×627 (2:1). width=400 gives ≈200px tall. Centered via columns.
    _l, _c, _r = st.columns([1, 1.2, 1])
    with _c:
        st.image(str(_header_img), width=400)

st.markdown(
    "<h1 style='margin-top:1rem;'>◉ &nbsp; SSP Revenue Reports <span class='mx-cursor'>_</span></h1>"
    f"<p style='color:var(--matrix-dim);'>"
    f"{len(SSPS)} ad networks · monthly revenue sync"
    "</p>",
    unsafe_allow_html=True,
)

# ── Settings: destination Google Sheet per SSP ────────────────────────────────

overrides = load_overrides()

with st.expander("Settings — destination Google Sheet per SSP", expanded=False):
    st.caption(
        "Paste a Google Sheets URL or just the sheet ID. "
        "Leave blank to use the script's built-in default. "
        "Remember to share the sheet with the service account "
        "(`theseus@dfp-api-157606.iam.gserviceaccount.com`) as Editor."
    )

    new_overrides = {}
    for name, _ in SSPS:
        current = overrides.get(name, {}).get("spreadsheet_id", "")
        value = st.text_input(
            f"{name}",
            value=current,
            key=f"sheet_input_{name}",
            placeholder="https://docs.google.com/spreadsheets/d/...",
        )
        cleaned = extract_sheet_id(value)
        if cleaned:
            new_overrides[name] = {"spreadsheet_id": cleaned}
            st.caption(
                f"↳ writes to [`{cleaned}`](https://docs.google.com/spreadsheets/d/{cleaned}/edit)"
            )
        else:
            st.caption("↳ using built-in default")

    if st.button("Save settings"):
        save_overrides(new_overrides)
        st.success("Settings saved.")
        overrides = new_overrides

st.divider()

# ── Run controls ──────────────────────────────────────────────────────────────

col_run, _ = st.columns([1.4, 2.6])
run_all_clicked = col_run.button(
    "▶ Run All Reports", type="primary", use_container_width=True
)

st.divider()

# Per-SSP row: name | execute | status | time
play_clicked_for: str | None = None
status_placeholders: dict[str, dict] = {}

for name, _ in SSPS:
    cols = st.columns([1.5, 0.6, 2.4, 0.8])
    cols[0].markdown(
        f"<div style='padding-top:0.45rem; font-family:Courier New,monospace; "
        f"color:var(--matrix-green); letter-spacing:1px; font-weight:700;'>"
        f"{name}</div>",
        unsafe_allow_html=True,
    )
    if cols[1].button("▶", key=f"play_{name}", help=f"Run only {name}"):
        play_clicked_for = name
    status_placeholders[name] = {
        "status": cols[2].empty(),
        "time":   cols[3].empty(),
    }
    status_placeholders[name]["status"].markdown(
        "<span style='color:var(--matrix-dim); font-family:Courier New,monospace;'>Idle</span>",
        unsafe_allow_html=True,
    )
    status_placeholders[name]["time"].markdown("")

summary_placeholder = st.empty()
detail_placeholder = st.empty()


# ── Run logic ─────────────────────────────────────────────────────────────────

# Matrix-flavored sysop messages — one randomly chosen per running SSP.
MATRIX_LINES = [
    "decrypting publisher signature",
    "tracing connection through 7 proxies",
    "bypassing firewall layer 3",
    "decoding ad-server hex stream",
    "loading kung fu",
    "following the white rabbit",
    "interfacing with the construct",
    "recompiling the matrix",
    "consulting the oracle",
    "wake up neo",
    "the spoon does not exist",
    "there is no impression count",
    "free your mind",
    "deja vu — patch deployed",
    "phasing through CPM gates",
    "establishing handshake with the source",
    "extracting payload from carrier signal",
    "spawning subroutines",
    "rendering the desert of the real",
    "choose the red pill",
    "cracking 256-bit publisher hash",
    "dodging packet collisions",
    "rerouting via Zion uplink",
]

ssps_to_run: list[tuple[str, object]] = []
if run_all_clicked:
    ssps_to_run = list(SSPS)
elif play_clicked_for:
    ssps_to_run = [(n, f) for n, f in SSPS if n == play_clicked_for]

if ssps_to_run:
    results = []
    for name, folder in ssps_to_run:
        line = random.choice(MATRIX_LINES)
        status_placeholders[name]["status"].markdown(
            f"<span class='mx-glitch'>Running…</span> "
            f"<span class='mx-decode' style='font-style:italic;'>"
            f"{line}<span class='mx-cursor'>_</span></span>",
            unsafe_allow_html=True,
        )
        status_placeholders[name]["time"].markdown(
            "<div style='padding-top:0.45rem; color:var(--matrix-warn); "
            "font-family:Courier New,monospace;' class='mx-glitch'>///</div>",
            unsafe_allow_html=True,
        )

        result = run_one(name, folder)
        results.append(result)

        if result["ok"]:
            status_placeholders[name]["status"].markdown(
                "<span class='mx-breach'>● Success</span>",
                unsafe_allow_html=True,
            )
        else:
            status_placeholders[name]["status"].markdown(
                f"<span class='mx-denied'>● Failed</span> "
                f"<code style='color:var(--matrix-fail); font-size:0.78em; "
                f"background:transparent; border:none;'>{result['error'][:60]}</code>",
                unsafe_allow_html=True,
            )
        status_placeholders[name]["time"].markdown(
            f"<div style='color:var(--matrix-dim); font-size:0.85em; padding-top:0.45rem; "
            f"font-family:Courier New,monospace;'>{result['elapsed']:.1f}s</div>",
            unsafe_allow_html=True,
        )

    save_run(results)
    ok = sum(1 for r in results if r["ok"])
    total = len(results)
    if ok == total:
        if total == 1:
            summary_placeholder.success(f"{results[0]['name']} succeeded.")
        else:
            summary_placeholder.success(f"All {ok}/{total} reports succeeded.")
    else:
        summary_placeholder.error(f"{ok}/{total} succeeded — details below.")

    with detail_placeholder.expander("Show full output", expanded=False):
        for r in results:
            mark = "●" if r["ok"] else "●"
            color = "var(--matrix-green)" if r["ok"] else "var(--matrix-fail)"
            st.markdown(
                f"<h3 style='color:{color};'>{mark} {r['name']} "
                f"<span style='color:var(--matrix-dim); font-size:0.7em;'>"
                f"· {r['elapsed']:.1f}s</span></h3>",
                unsafe_allow_html=True,
            )
            st.code(r["output"] or "(no output)", language="text")


# ── History ───────────────────────────────────────────────────────────────────

st.divider()
st.markdown(
    "<h3>Recent runs</h3>",
    unsafe_allow_html=True,
)

history = load_history()
if not history:
    st.caption("No runs recorded yet.")
else:
    for entry in history[:15]:
        ts = entry["timestamp"].replace("T", " ")
        ok_count = sum(1 for r in entry["results"] if r["ok"])
        total    = len(entry["results"])
        per_ssp_parts = []
        for r in entry["results"]:
            color = "var(--matrix-green)" if r["ok"] else "var(--matrix-fail)"
            mark  = "✓" if r["ok"] else "✗"
            per_ssp_parts.append(
                f"<span style='color:{color}; font-size:0.85em; font-family:Courier New,monospace;'>"
                f"{mark} {r['name']}</span>"
            )
        per_ssp = "  ".join(per_ssp_parts)
        st.markdown(
            f"<div style='padding:0.4rem 0; border-bottom:1px dashed rgba(0,255,65,0.12); "
            f"font-family:Courier New,monospace;'>"
            f"<span style='color:var(--matrix-dim); font-size:0.85em;'>{ts}</span>  "
            f"{per_ssp}  "
            f"<span style='color:var(--matrix-dim); font-size:0.8em; float:right;'>"
            f"[{ok_count}/{total}]</span>"
            f"</div>",
            unsafe_allow_html=True,
        )

