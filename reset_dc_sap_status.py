#!/usr/bin/env python3
"""
reset_dc_sap_status.py
======================
DEV / TESTING UTILITY — resets debit/credit PDF SAP records back to "pending"
so they can be retried by the SAP automation service.

Modes:
  • Default (no flags): resets only "broken" records — those marked as "done"
    but with no checklist number and no PDF (incorrectly marked complete).
  • --all : resets ALL debit_credit_pdfs records to pending (full re-test).
  • --ids 1,2,3 : resets only the specified record IDs.

Usage:
    python reset_dc_sap_status.py              # reset only broken ones
    python reset_dc_sap_status.py --all        # reset ALL to pending
    python reset_dc_sap_status.py --ids 1,2,3  # reset specific IDs
"""

import argparse
import os
import sqlite3
import sys
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
NIS_DIR = os.path.join(ROOT, "nis_booking")
sys.path.insert(0, NIS_DIR)

import db_manager as db  # noqa: E402 — gives us the shared DB path


def _connect() -> sqlite3.Connection:
    # Ensure schema exists (creates tables if missing)
    db.init_db()
    conn = sqlite3.connect(str(db._DB_PATH), timeout=30)
    conn.row_factory = sqlite3.Row
    return conn


def _ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def reset_dc_sap_status(mode: str = "broken", ids: list = None):
    """Reset debit/credit PDF SAP records.

    mode:
        "broken" — only records where sap_status='done' but no checklist/pdf
        "all"    — every record in debit_credit_pdfs
        "ids"    — only the specified record IDs
    """
    print(f"\n{'═' * 65}")
    print(f"  DEBIT/CREDIT PDF SAP STATUS RESET  —  {_ts()}")
    print(f"  DB: {db._DB_PATH}")
    print(f"  Mode: {mode}")
    print(f"{'═' * 65}\n")

    conn = _connect()

    # ── Fetch target rows ────────────────────────────────────────────────────
    if mode == "all":
        rows = conn.execute(
            "SELECT * FROM debit_credit_pdfs ORDER BY id"
        ).fetchall()
    elif mode == "ids":
        if not ids:
            print("  ⚠️  No IDs provided. Nothing to do.\n")
            conn.close()
            return
        placeholders = ", ".join("?" for _ in ids)
        rows = conn.execute(
            f"SELECT * FROM debit_credit_pdfs WHERE id IN ({placeholders}) ORDER BY id",
            ids,
        ).fetchall()
    else:
        # Default: "broken" — done but missing checklist number AND pdf name
        rows = conn.execute(
            "SELECT * FROM debit_credit_pdfs "
            "WHERE sap_status = 'done' "
            "AND (sap_checklist_number IS NULL OR TRIM(sap_checklist_number) = '') "
            "AND (sap_pdf_name IS NULL OR TRIM(sap_pdf_name) = '') "
            "ORDER BY id"
        ).fetchall()

    rows = [dict(r) for r in rows]

    if not rows:
        print("  ℹ️  No matching debit_credit_pdfs records found.\n")
        conn.close()
        return

    # ── Display what was found ───────────────────────────────────────────────
    print(f"  Found {len(rows)} record(s):\n")
    print(f"  {'ID':<5} {'PDF Name':<30} {'Vendor':<10} {'Amount':<12} {'Ref#':<15} {'Doc Date':<12} {'SAP Status'}")
    print(f"  {'─' * 5} {'─' * 30} {'─' * 10} {'─' * 12} {'─' * 15} {'─' * 12} {'─' * 10}")

    for r in rows:
        pdf_name = str(r.get("pdf_name") or "")[:29]
        vendor = str(r.get("vendor_code") or "")[:9]
        amount = str(r.get("amount") or "")[:11]
        ref_num = str(r.get("ref_number") or "")[:14]
        doc_date = str(r.get("doc_date") or "")[:11]
        sap_status = str(r.get("sap_status") or "")
        print(f"  {r['id']:<5} {pdf_name:<30} {vendor:<10} {amount:<12} {ref_num:<15} {doc_date:<12} {sap_status}")

    print()

    # ── Apply reset ──────────────────────────────────────────────────────────
    id_list = [r["id"] for r in rows]
    placeholders = ", ".join("?" for _ in id_list)
    conn.execute(
        f"UPDATE debit_credit_pdfs "
        f"SET sap_status = 'pending', error_message = '', updated_at = ? "
        f"WHERE id IN ({placeholders})",
        [_ts()] + id_list,
    )
    conn.commit()
    conn.close()

    print(f"  ✅  Reset {len(rows)} record(s) → sap_status='pending', error_message cleared.")
    print(f"\n{'═' * 65}\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Reset debit/credit PDF SAP records to 'pending' for retry."
    )
    parser.add_argument(
        "--all", action="store_true", dest="reset_all",
        help="Reset ALL debit_credit_pdfs records to pending (full re-test).",
    )
    parser.add_argument(
        "--ids", default="",
        help="Comma-separated list of specific record IDs to reset (e.g. --ids 1,2,3).",
    )
    args = parser.parse_args()

    if args.reset_all:
        reset_dc_sap_status(mode="all")
    elif args.ids:
        id_list = [int(x.strip()) for x in args.ids.split(",") if x.strip().isdigit()]
        reset_dc_sap_status(mode="ids", ids=id_list)
    else:
        reset_dc_sap_status(mode="broken")
