#!/usr/bin/env python3
"""
backfill_sap_steps.py
=====================
Recovery / maintenance script.

Finds booking_rows where NIS is complete (status='nis_done') but SAP
automation steps may be missing or stuck, then:

1. For each row, inspects negative_fields_json for non-zero amounts
2. For each negative field, checks if the corresponding
   sap_automation_<field>_row_<N>_<spv> step exists in run_steps
3. If MISSING  → creates it with db.ensure_dynamic_step()
4. If EXISTS but is 'done' with no checklist number → resets to 'pending'
5. Prints a summary

Safe to run multiple times — ensure_dynamic_step is a no-op for existing steps.

Usage:
    python backfill_sap_steps.py           # find and create missing steps
    python backfill_sap_steps.py --dry-run # preview only
"""

import argparse
import json
import os
import re
import sqlite3
import sys
from datetime import datetime

ROOT    = os.path.dirname(os.path.abspath(__file__))
NIS_DIR = os.path.join(ROOT, "nis_booking")
sys.path.insert(0, NIS_DIR)

import db_manager as db

# ── SPV amount fields that require SAP automation ────────────────────────────
SPV_AMOUNT_FIELDS = [
    "total_dsm_charges_payable",
    "drawl_charges_payable",
    "revenue_diff",
    "revenue_loss",
]


# ── Helpers ──────────────────────────────────────────────────────────────────
def _row_step_name(base_step: str, row_index: int, spv_name: str) -> str:
    """Build the dynamic step name — mirrors sap_main.py convention."""
    safe_spv = re.sub(r"[^A-Za-z0-9]+", "_", str(spv_name or "spv")).strip("_")[:40] or "spv"
    return f"{base_step}_row_{row_index}_{safe_spv}"


def _safe_float(val) -> float:
    try:
        return float(str(val).replace(",", "").strip())
    except Exception:
        return 0.0


def _json_dict(raw) -> dict:
    try:
        v = json.loads(raw or "{}")
        return v if isinstance(v, dict) else {}
    except Exception:
        return {}


def _has_checklist_number(detail: str) -> bool:
    """Check if the step detail contains a 7+ digit checklist number."""
    return bool(re.search(r'\b\d{7,}\b', str(detail or "")))


def _ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ── Main ─────────────────────────────────────────────────────────────────────
def main(dry_run: bool = False):
    db.init_db()

    print(f"\n{'═'*65}")
    print(f"  SAP STEP BACKFILL{'  [DRY RUN — no writes]' if dry_run else ''}")
    print(f"  {_ts()}")
    print(f"  DB: {db._DB_PATH}")
    print(f"{'═'*65}\n")

    # Fetch booking rows with status = 'nis_done' (NIS complete, SAP pending)
    with db._LOCK:
        conn = db._connect()
        rows = conn.execute(
            "SELECT * FROM booking_rows WHERE status = 'nis_done' "
            "ORDER BY run_id, row_index"
        ).fetchall()
    rows = [dict(r) for r in rows]

    print(f"  Found {len(rows)} booking row(s) with status='nis_done'.\n")

    if not rows:
        print("  ℹ️   Nothing to backfill — no rows pending SAP automation.\n")
        return

    created = 0
    reset_count = 0
    already_ok = 0
    skipped = 0

    for row in rows:
        run_id    = str(row.get("run_id") or "")
        row_index = int(row.get("row_index") or 0)
        spv_name  = str(row.get("spv_name") or "").strip()
        negative  = _json_dict(row.get("negative_fields_json"))

        if not negative:
            skipped += 1
            continue

        print(f"  📋  row#{row_index} [{spv_name}]  run_id={run_id[:30]}...")

        for field_key in SPV_AMOUNT_FIELDS:
            amt = negative.get(field_key)
            if amt is None:
                continue

            abs_amt = abs(_safe_float(amt))
            if abs_amt <= 0:
                continue

            # Build the expected step name
            step_name = _row_step_name(db.step_sap(field_key), row_index, spv_name)
            existing = db.get_step(run_id, step_name)

            if existing is None:
                # Step is MISSING — create it
                if dry_run:
                    print(f"        [DRY-RUN] would create: {step_name}")
                else:
                    db.ensure_dynamic_step(run_id, step_name)
                    print(f"        ➕  Created pending step: {step_name}")
                created += 1

            elif existing.get("status") == db.STATUS_DONE:
                # Step exists and is done — check if it has a checklist number
                detail = str(existing.get("detail") or "")
                if not _has_checklist_number(detail):
                    # Done but no checklist → reset to pending
                    if dry_run:
                        print(f"        [DRY-RUN] would reset (done, no checklist): {step_name}")
                    else:
                        with db._LOCK:
                            conn = db._connect()
                            with conn:
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
                                    (run_id, step_name),
                                )
                        print(f"        🔄  Reset (done, no checklist): {step_name}")
                    reset_count += 1
                else:
                    # Done with a valid checklist — nothing to do
                    already_ok += 1
            else:
                # Step exists and is pending/running/failed — leave it alone
                already_ok += 1

    # ── Summary ──────────────────────────────────────────────────────────────
    print(f"\n{'─'*65}")
    print(f"  SUMMARY")
    print(f"{'─'*65}")
    print(f"  Rows scanned (nis_done)  : {len(rows)}")
    print(f"  Rows with no negative    : {skipped}")
    print(f"  Steps {'would create' if dry_run else 'created'}       : {created}")
    print(f"  Steps {'would reset' if dry_run else 'reset'}         : {reset_count}")
    print(f"  Steps already OK         : {already_ok}")
    print(f"{'═'*65}\n")

    if dry_run:
        if created or reset_count:
            print("  ℹ️   Re-run without --dry-run to apply these changes.\n")
        else:
            print("  ℹ️   Nothing to backfill — all steps are in order.\n")
    else:
        if created or reset_count:
            print("  ✅  Backfill complete. Now run:  python run_sap_automation.py\n")
        else:
            print("  ℹ️   Nothing to backfill — all steps are in order.\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Backfill missing SAP automation steps for nis_done booking rows."
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Preview only — no DB changes."
    )
    args = parser.parse_args()
    main(dry_run=args.dry_run)
