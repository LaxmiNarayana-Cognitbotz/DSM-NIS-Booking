#!/usr/bin/env python3
"""
reset_sap_checklist.py
======================
DEV / TESTING UTILITY — removes SAP checklist numbers and related data
so you can re-run SAP automation tests from scratch.

What it resets:
  • email_runs: checklist_value, sap_checklist_pdf_uploaded
  • booking_rows: checklist_value, sap_checklist_pdf_uploaded
  • run_steps: all sap_automation_* steps → status='pending', retry_count=0
  • debit_credit_pdfs: sap_status → 'pending', clears sap_checklist_number,
    sap_pdf_name, sap_pdf_sp_path

Modes:
  Default (no flags): reset ALL runs that have SAP checklist data
  --run-id <ID>     : reset only a specific run
  --dc-only         : reset only debit/credit PDF SAP records
  --dry-run         : preview what would be reset, no changes applied

Usage:
    python reset_sap_checklist.py                  # reset all SAP checklist data
    python reset_sap_checklist.py --run-id <ID>    # reset specific run only
    python reset_sap_checklist.py --dc-only        # reset only debit/credit PDFs
    python reset_sap_checklist.py --dry-run        # preview, no changes
"""

import argparse
import os
import sqlite3
import sys
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
NIS_DIR = os.path.join(ROOT, "nis_booking")
sys.path.insert(0, NIS_DIR)

import db_manager as db


def _connect() -> sqlite3.Connection:
    db.init_db()
    conn = sqlite3.connect(str(db._DB_PATH), timeout=30)
    conn.row_factory = sqlite3.Row
    return conn


def _ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def reset_sap_checklist(run_id: str = "", dry_run: bool = False, dc_only: bool = False):
    print(f"\n{'═' * 70}")
    print(f"  SAP CHECKLIST DATA RESET  {'[DRY RUN]' if dry_run else ''}  —  {_ts()}")
    print(f"  DB: {db._DB_PATH}")
    if run_id:
        print(f"  Scope: run_id = {run_id}")
    if dc_only:
        print(f"  Scope: Debit/Credit PDFs only")
    print(f"{'═' * 70}\n")

    conn = _connect()

    # ─────────────────────────────────────────────────────────────────────────
    # 1. Reset email_runs (checklist_value, sap_checklist_pdf_uploaded)
    # ─────────────────────────────────────────────────────────────────────────
    if not dc_only:
        print("  ┌─ EMAIL RUNS ─────────────────────────────────────────────────")
        if run_id:
            runs = conn.execute(
                "SELECT * FROM email_runs WHERE id = ? OR message_id = ?",
                (run_id, run_id),
            ).fetchall()
        else:
            runs = conn.execute(
                "SELECT * FROM email_runs "
                "WHERE COALESCE(TRIM(checklist_value), '') != '' "
                "   OR sap_checklist_pdf_uploaded = 1"
            ).fetchall()

        runs = [dict(r) for r in runs]
        if runs:
            for r in runs:
                rid = r.get("id", "")[:50]
                spv = r.get("spv_name", "")
                old_val = r.get("checklist_value", "")
                old_pdf = r.get("sap_checklist_pdf_uploaded", 0)
                print(f"  │  {rid}")
                print(f"  │    SPV: {spv}")
                print(f"  │    checklist_value: '{old_val}' → ''")
                print(f"  │    sap_checklist_pdf_uploaded: {old_pdf} → 0")
                print(f"  │")

            if not dry_run:
                if run_id:
                    conn.execute(
                        "UPDATE email_runs "
                        "SET checklist_value = '', sap_checklist_pdf_uploaded = 0 "
                        "WHERE id = ? OR message_id = ?",
                        (run_id, run_id),
                    )
                else:
                    conn.execute(
                        "UPDATE email_runs "
                        "SET checklist_value = '', sap_checklist_pdf_uploaded = 0 "
                        "WHERE COALESCE(TRIM(checklist_value), '') != '' "
                        "   OR sap_checklist_pdf_uploaded = 1"
                    )
            print(f"  │  {'Would reset' if dry_run else 'Reset'}: {len(runs)} email_runs row(s)")
        else:
            print(f"  │  No email_runs with SAP checklist data found.")
        print(f"  └─────────────────────────────────────────────────────────────\n")

    # ─────────────────────────────────────────────────────────────────────────
    # 2. Reset booking_rows (checklist_value, sap_checklist_pdf_uploaded)
    # ─────────────────────────────────────────────────────────────────────────
    if not dc_only:
        print("  ┌─ BOOKING ROWS ───────────────────────────────────────────────")
        if run_id:
            brows = conn.execute(
                "SELECT * FROM booking_rows WHERE run_id = ? OR message_id = ?",
                (run_id, run_id),
            ).fetchall()
        else:
            brows = conn.execute(
                "SELECT * FROM booking_rows "
                "WHERE COALESCE(TRIM(checklist_value), '') != '' "
                "   OR sap_checklist_pdf_uploaded = 1"
            ).fetchall()

        brows = [dict(r) for r in brows]
        if brows:
            for r in brows:
                row_id = r.get("row_id")
                spv = r.get("spv_name", "")
                old_val = r.get("checklist_value", "")
                old_pdf = r.get("sap_checklist_pdf_uploaded", 0)
                status = r.get("status", "")
                print(f"  │  row_id={row_id}  spv={spv}  status={status}")
                print(f"  │    checklist_value: '{old_val}' → ''")
                print(f"  │    sap_checklist_pdf_uploaded: {old_pdf} → 0")

                # If the row was marked 'done' solely because SAP finished,
                # reset it back to 'nis_done' so the SAP service picks it up again
                if status == "done" and old_val:
                    print(f"  │    status: 'done' → 'nis_done' (so SAP re-processes it)")
                print(f"  │")

            if not dry_run:
                if run_id:
                    # Reset checklist data
                    conn.execute(
                        "UPDATE booking_rows "
                        "SET checklist_value = '', sap_checklist_pdf_uploaded = 0 "
                        "WHERE run_id = ? OR message_id = ?",
                        (run_id, run_id),
                    )
                    # Reset status from 'done' back to 'nis_done' for rows that
                    # had SAP checklist data (meaning SAP completion was what
                    # marked them done)
                    conn.execute(
                        "UPDATE booking_rows "
                        "SET status = 'nis_done' "
                        "WHERE (run_id = ? OR message_id = ?) "
                        "AND status = 'done' "
                        "AND COALESCE(TRIM(checklist_value), '') != ''",
                        (run_id, run_id),
                    )
                else:
                    conn.execute(
                        "UPDATE booking_rows "
                        "SET checklist_value = '', sap_checklist_pdf_uploaded = 0 "
                        "WHERE COALESCE(TRIM(checklist_value), '') != '' "
                        "   OR sap_checklist_pdf_uploaded = 1"
                    )
                    conn.execute(
                        "UPDATE booking_rows "
                        "SET status = 'nis_done' "
                        "WHERE status = 'done' "
                        "AND COALESCE(TRIM(checklist_value), '') != ''"
                    )
            print(f"  │  {'Would reset' if dry_run else 'Reset'}: {len(brows)} booking_rows row(s)")
        else:
            print(f"  │  No booking_rows with SAP checklist data found.")
        print(f"  └─────────────────────────────────────────────────────────────\n")

    # ─────────────────────────────────────────────────────────────────────────
    # 3. Reset run_steps (all sap_automation_* steps → pending)
    # ─────────────────────────────────────────────────────────────────────────
    if not dc_only:
        print("  ┌─ SAP AUTOMATION STEPS ───────────────────────────────────────")
        if run_id:
            steps = conn.execute(
                "SELECT * FROM run_steps "
                "WHERE run_id = ? AND step_name LIKE 'sap_automation_%'",
                (run_id,),
            ).fetchall()
        else:
            steps = conn.execute(
                "SELECT * FROM run_steps WHERE step_name LIKE 'sap_automation_%'"
            ).fetchall()

        steps = [dict(r) for r in steps]
        if steps:
            status_icons = {"done": "✅", "skipped": "⏭ ", "pending": "⏳",
                            "running": "🔄", "failed": "❌"}
            for s in steps:
                icon = status_icons.get(s.get("status"), "❓")
                print(f"  │  {icon}  {s.get('step_name', '')}")
                print(f"  │       run_id={str(s.get('run_id', ''))[:50]}  "
                      f"status={s.get('status')}  retries={s.get('retry_count', 0)}")
                if s.get("detail"):
                    print(f"  │       detail: {str(s.get('detail', ''))[:80]}")
                print(f"  │       → status='pending', retry_count=0, detail/error cleared")
                print(f"  │")

            if not dry_run:
                if run_id:
                    conn.execute(
                        "UPDATE run_steps "
                        "SET status = 'pending', retry_count = 0, "
                        "    error = '', detail = '', started_at = '', finished_at = '' "
                        "WHERE run_id = ? AND step_name LIKE 'sap_automation_%'",
                        (run_id,),
                    )
                else:
                    conn.execute(
                        "UPDATE run_steps "
                        "SET status = 'pending', retry_count = 0, "
                        "    error = '', detail = '', started_at = '', finished_at = '' "
                        "WHERE step_name LIKE 'sap_automation_%'"
                    )
            print(f"  │  {'Would reset' if dry_run else 'Reset'}: {len(steps)} step(s)")
        else:
            print(f"  │  No sap_automation_* steps found.")
        print(f"  └─────────────────────────────────────────────────────────────\n")

    # ─────────────────────────────────────────────────────────────────────────
    # 4. Reset debit_credit_pdfs SAP data
    # ─────────────────────────────────────────────────────────────────────────
    print("  ┌─ DEBIT/CREDIT PDF SAP RECORDS ───────────────────────────────")
    dc_rows = conn.execute(
        "SELECT * FROM debit_credit_pdfs "
        "WHERE sap_status != 'pending' "
        "   OR COALESCE(TRIM(sap_checklist_number), '') != '' "
        "ORDER BY id"
    ).fetchall()
    dc_rows = [dict(r) for r in dc_rows]

    if dc_rows:
        print(f"  │  {'ID':<5} {'PDF Name':<28} {'Vendor':<8} {'Amount':<10} "
              f"{'SAP Status':<10} {'Checklist#'}")
        print(f"  │  {'─'*5} {'─'*28} {'─'*8} {'─'*10} {'─'*10} {'─'*12}")
        for r in dc_rows:
            print(f"  │  {r['id']:<5} "
                  f"{str(r.get('pdf_name') or '')[:27]:<28} "
                  f"{str(r.get('vendor_code') or '')[:7]:<8} "
                  f"{str(r.get('amount') or '')[:9]:<10} "
                  f"{str(r.get('sap_status') or '')[:9]:<10} "
                  f"{str(r.get('sap_checklist_number') or '')[:12]}")
        print(f"  │")
        print(f"  │  All → sap_status='pending', sap_checklist_number='', "
              f"sap_pdf_name='', sap_pdf_sp_path='', error_message=''")

        if not dry_run:
            conn.execute(
                "UPDATE debit_credit_pdfs "
                "SET sap_status = 'pending', "
                "    sap_checklist_number = '', "
                "    sap_pdf_name = '', "
                "    sap_pdf_sp_path = '', "
                "    error_message = '', "
                "    updated_at = ? "
                "WHERE sap_status != 'pending' "
                "   OR COALESCE(TRIM(sap_checklist_number), '') != ''",
                (_ts(),),
            )
        print(f"  │  {'Would reset' if dry_run else 'Reset'}: {len(dc_rows)} record(s)")
    else:
        print(f"  │  No debit_credit_pdfs with SAP data found.")
    print(f"  └─────────────────────────────────────────────────────────────\n")

    # ─────────────────────────────────────────────────────────────────────────
    # 5. Clean up local SAP-completed marker files
    # ─────────────────────────────────────────────────────────────────────────
    marker_dir = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "sap_automation", "temp-book", "_completed_sap_tasks"
    )
    if os.path.isdir(marker_dir):
        markers = [f for f in os.listdir(marker_dir) if f.endswith(".json")]
        if markers:
            print(f"  ┌─ LOCAL MARKER FILES ─────────────────────────────────────────")
            for m in markers:
                print(f"  │  {m}")
            if not dry_run:
                import shutil
                shutil.rmtree(marker_dir, ignore_errors=True)
                print(f"  │  Deleted {len(markers)} marker file(s)")
            else:
                print(f"  │  Would delete {len(markers)} marker file(s)")
            print(f"  └─────────────────────────────────────────────────────────────\n")

    # ─────────────────────────────────────────────────────────────────────────
    # Commit
    # ─────────────────────────────────────────────────────────────────────────
    if not dry_run:
        conn.commit()
    conn.close()

    print(f"{'═' * 70}")
    if dry_run:
        print(f"  ℹ️   DRY RUN complete — no changes applied.")
        print(f"  Re-run without --dry-run to apply the reset.")
    else:
        print(f"  ✅  SAP checklist data cleared successfully!")
        print(f"  You can now re-run:  python run_sap_automation.py")
    print(f"{'═' * 70}\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Reset SAP checklist numbers and related data for re-testing.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python reset_sap_checklist.py                  # reset everything
  python reset_sap_checklist.py --dry-run        # preview only
  python reset_sap_checklist.py --run-id <ID>    # reset one run
  python reset_sap_checklist.py --dc-only        # only debit/credit PDFs
        """,
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Preview what would be reset — no DB changes.",
    )
    parser.add_argument(
        "--run-id", default="",
        help="Reset only a specific run (by run_id or message_id).",
    )
    parser.add_argument(
        "--dc-only", action="store_true",
        help="Reset only debit/credit PDF SAP records (skip email_runs/booking_rows/steps).",
    )
    args = parser.parse_args()

    reset_sap_checklist(
        run_id=args.run_id,
        dry_run=args.dry_run,
        dc_only=args.dc_only,
    )
