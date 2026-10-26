#!/usr/bin/env python3
"""
reset_sap_steps.py
==================
DEV / TESTING UTILITY — resets SAP automation steps so you can re-run them.

What it does:
  Default mode:
  • Finds FAILED sap_automation_* steps that have reached max retries (>= 3)
  • Resets their status → 'pending'

  --broken mode:
  • Finds DONE sap_automation_* steps where no checklist was captured
    (the step ran with empty vendor data or SAP save failed)
  • Resets them → 'pending'

  --all mode:
  • Resets ALL sap_automation_* steps (done/failed/pending) → 'pending'

Usage:
    python reset_sap_steps.py                   # reset failed with max retries
    python reset_sap_steps.py --broken          # reset done-but-broken steps
    python reset_sap_steps.py --all             # reset ALL SAP steps
    python reset_sap_steps.py --run-id <ID>     # scope to one run
    python reset_sap_steps.py --dry-run         # preview only, no changes
"""

import argparse
import json
import os
import sqlite3
import sys
from datetime import datetime

ROOT    = os.path.dirname(os.path.abspath(__file__))
NIS_DIR = os.path.join(ROOT, "nis_booking")
sys.path.insert(0, NIS_DIR)

import db_manager as db   # gives us the DB path via config


def _connect() -> sqlite3.Connection:
    conn = sqlite3.connect(db._DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    return conn


def _ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def reset_sap_steps(run_id: str = "", dry_run: bool = False, mode: str = "failed"):
    print(f"\n{'═'*65}")
    print(f"  SAP STEP RESET  [{mode.upper()}]{'  [DRY RUN]' if dry_run else ''}  —  {_ts()}")
    print(f"  DB: {db._DB_PATH}")
    print(f"{'═'*65}\n")

    conn = _connect()

    # ── Build query based on mode ───────────────────────────────────────────
    if mode == "all":
        if run_id:
            rows = conn.execute(
                "SELECT * FROM run_steps "
                "WHERE run_id = ? AND step_name LIKE 'sap_automation_%'",
                (str(run_id),)
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM run_steps WHERE step_name LIKE 'sap_automation_%'"
            ).fetchall()
    elif mode == "broken":
        # Done steps where detail doesn't contain a checklist number (7+ digits)
        if run_id:
            rows = conn.execute(
                "SELECT * FROM run_steps "
                "WHERE run_id = ? AND step_name LIKE 'sap_automation_%' "
                "AND status = 'done'",
                (str(run_id),)
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM run_steps "
                "WHERE step_name LIKE 'sap_automation_%' AND status = 'done'"
            ).fetchall()
        # Filter: only keep rows where the detail doesn't have a 7-digit checklist
        import re
        rows = [r for r in rows
                if not re.search(r'\b\d{7,}\b', str(dict(r).get("detail") or ""))]
    else:
        # Default: failed with max retries
        if run_id:
            rows = conn.execute(
                "SELECT * FROM run_steps "
                "WHERE run_id = ? AND step_name LIKE 'sap_automation_%' "
                "AND status = 'failed' AND retry_count >= 3",
                (str(run_id),)
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM run_steps "
                "WHERE step_name LIKE 'sap_automation_%' "
                "AND status = 'failed' AND retry_count >= 3"
            ).fetchall()

    rows = [dict(r) for r in rows]

    if not rows:
        print("  ℹ️   No matching sap_automation_* steps found.\n")
        return

    print(f"  Found {len(rows)} SAP step(s):\n")

    reset_count = 0

    for r in rows:
        step_name   = r.get("step_name", "")
        status      = r.get("status", "")
        retry_count = int(r.get("retry_count") or 0)
        detail      = str(r.get("detail") or "")[:100]
        error       = str(r.get("error") or "")[:80]
        r_run_id    = str(r.get("run_id") or "")

        status_icon = {"done": "✅", "skipped": "⏭ ", "pending": "⏳",
                       "running": "🔄", "failed": "❌"}.get(status, "❓")

        print(f"  {status_icon}  {step_name}")
        print(f"       run_id={r_run_id[:50]}  status={status}  retries={retry_count}")
        if detail:
            print(f"       detail: {detail}")
        if error:
            print(f"       error: {error}")

        if dry_run:
            print(f"       [DRY-RUN] would reset → pending, retry_count=0\n")
            reset_count += 1
            continue

        conn.execute(
            """UPDATE run_steps
               SET status      = 'pending',
                   retry_count = 0,
                   error       = '',
                   started_at  = '',
                   finished_at = '',
                   detail      = ''
               WHERE run_id = ? AND step_name = ?
            """,
            (r_run_id, step_name),
        )
        print(f"       ➡️   Reset → pending, retry_count=0\n")
        reset_count += 1

    if not dry_run:
        conn.commit()

    conn.close()

    print(f"{'─'*65}")
    print(f"  {'Would reset' if dry_run else 'Reset'}: {reset_count} step(s)")
    print(f"{'═'*65}\n")

    if not dry_run:
        print("  ✅  Done. Now run:  python run_sap_automation.py\n")
    else:
        print("  ℹ️   Re-run without --dry-run to apply.\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Reset sap_automation_* steps to pending with retry_count=0."
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Preview only — no DB changes."
    )
    parser.add_argument(
        "--run-id", default="",
        help="Reset only steps belonging to this specific run_id."
    )
    parser.add_argument(
        "--broken", action="store_true",
        help="Reset 'done' steps that have no checklist number (ran with empty vendor data)."
    )
    parser.add_argument(
        "--all", action="store_true", dest="reset_all",
        help="Reset ALL sap_automation_* steps regardless of status."
    )
    args = parser.parse_args()

    if args.reset_all:
        mode = "all"
    elif args.broken:
        mode = "broken"
    else:
        mode = "failed"

    reset_sap_steps(run_id=args.run_id, dry_run=args.dry_run, mode=mode)
