#!/usr/bin/env python3
"""
Master reporting script (CLI).

Runs every SSP's individual sync script in sequence and prints a summary.
Shares its SSP list with the Streamlit dashboard (app.py) via runner.py.
"""

import sys
from datetime import datetime

from runner import SSPS, run_one, save_run


def banner(msg: str) -> None:
    print()
    print("═" * 70)
    print(f"  {msg}")
    print("═" * 70, flush=True)


def main() -> int:
    banner(f"Daily SSP Sync — {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print(f"Running {len(SSPS)} SSP scripts in sequence.\n", flush=True)

    results = []
    for name, folder in SSPS:
        banner(f"▶  {name}")
        result = run_one(name, folder)
        sys.stdout.write(result["output"])
        sys.stdout.flush()
        results.append(result)

    save_run(results)

    banner("Summary")
    ok_count = sum(1 for r in results if r["ok"])
    for r in results:
        mark = "✓" if r["ok"] else "✗"
        line = f"  {mark}  {r['name']:<10}  {r['elapsed']:5.1f}s"
        if not r["ok"]:
            line += f"   FAILED — {r['error']}"
        print(line)
    print()
    print(f"  {ok_count} / {len(results)} succeeded.")
    print()
    return 0 if ok_count == len(results) else 1


if __name__ == "__main__":
    sys.exit(main())
