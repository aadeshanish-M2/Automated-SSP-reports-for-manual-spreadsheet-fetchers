"""
Shared runner logic — used by both run_all.py (CLI) and app.py (web dashboard).
"""

import json
import os
import subprocess
import time
from datetime import datetime
from pathlib import Path

# ── Project layout ────────────────────────────────────────────────────────────
# Single repo, ready for cloud deploy:
#   project_root/
#   ├── master/        ← this file, app.py, run_all.py
#   ├── ssps/<slug>/   ← each SSP's main.py
#   └── .venv/         ← single shared venv at the project root

HERE         = Path(__file__).parent        # …/master
PROJECT_ROOT = HERE.parent                  # project root
SSPS_DIR     = PROJECT_ROOT / "ssps"

SSPS = [
    ("Teal",       SSPS_DIR / "teal"),
    ("Nexxen",     SSPS_DIR / "nexxen"),
    ("MGID",       SSPS_DIR / "mgid"),
    ("Kuantyx",    SSPS_DIR / "kuantyx"),
    ("Adsolut",    SSPS_DIR / "adsolut"),
    ("Yandex",     SSPS_DIR / "yandex"),
    ("Insticator", SSPS_DIR / "insticator"),
    ("Vidoomy",    SSPS_DIR / "vidoomy"),
    ("NoBid",      SSPS_DIR / "nobid"),
    ("RichAudience", SSPS_DIR / "richaudience"),
    ("MediaFuse",  SSPS_DIR / "mediafuse"),
    ("Ogury",      SSPS_DIR / "ogury"),
    ("OptiDigital", SSPS_DIR / "optidigital"),
    ("Minute Media", SSPS_DIR / "minutemedia"),
    # Connatix runs last because its refresh can take 5+ min and we'd rather
    # not block the fast ones behind it.
    ("Connatix",   SSPS_DIR / "connatix"),
]

HISTORY_PATH    = HERE / "history.json"
MAX_HISTORY     = 50


# ── Per-SSP run ───────────────────────────────────────────────────────────────

def run_one(name: str, folder: Path) -> dict:
    # Use the same Python interpreter that's running this process — single
    # shared venv at the project root (or whatever is on PATH on cloud).
    import sys as _sys
    python = Path(_sys.executable)
    script = folder / "main.py"

    if not script.exists():
        return {"name": name, "ok": False, "elapsed": 0.0,
                "error": f"Missing script: {script}", "output": ""}

    env = os.environ.copy()

    start = time.time()
    proc = subprocess.run(
        [str(python), str(script)],
        cwd=str(folder),
        capture_output=True,
        text=True,
        env=env,
    )
    elapsed = time.time() - start

    output = (proc.stdout or "") + ("\n" + proc.stderr if proc.stderr else "")
    if proc.returncode == 0:
        return {"name": name, "ok": True, "elapsed": elapsed,
                "error": "", "output": output}

    lines = [l for l in (proc.stderr or proc.stdout).strip().splitlines() if l.strip()]
    last = lines[-1] if lines else "(no output)"
    return {"name": name, "ok": False, "elapsed": elapsed,
            "error": last, "output": output}


# ── History persistence ───────────────────────────────────────────────────────

def load_history() -> list[dict]:
    if not HISTORY_PATH.exists():
        return []
    try:
        return json.loads(HISTORY_PATH.read_text())
    except Exception:
        return []


def save_run(results: list[dict]) -> None:
    history = load_history()
    entry = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "results": [
            {"name": r["name"], "ok": r["ok"],
             "elapsed": round(r["elapsed"], 1),
             "error": r["error"]}
            for r in results
        ],
    }
    history.insert(0, entry)
    history = history[:MAX_HISTORY]
    HISTORY_PATH.write_text(json.dumps(history, indent=2))
