"""
clean_db.py
===========
Wipe the local NIS booking SQLite database (``config.SQLITE_DB_PATH``).

This ONLY touches the local SQLite tracking DB (email_runs, booking_rows,
run_steps). It never touches SharePoint (DSM-NIS-INPUT / DSM-NIS-OUTPUT /
nis_completion_recipients).

Usage
-----
    python clean_db.py                 # clear all rows (asks for confirmation)
    python clean_db.py --yes           # clear all rows, no prompt
    python clean_db.py --delete-file   # delete the .db file entirely (+ WAL/SHM)
    python clean_db.py --delete-file --yes
    python clean_db.py --show          # just show current row counts and exit
"""

from __future__ import annotations

import argparse
import os
import sqlite3
import sys

import config

DB_PATH = str(getattr(config, "SQLITE_DB_PATH", "nis_booking.db"))
_TABLES = ["debit_credit_pdfs", "booking_rows", "run_steps", "debit_credit_emails", "email_runs"]


def _existing_tables(conn: sqlite3.Connection) -> list:
    rows = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'"
    ).fetchall()
    return [r[0] for r in rows]


def show_counts() -> bool:
    if not os.path.exists(DB_PATH):
        print(f"[CLEAN-DB] No database file at: {DB_PATH}")
        return False
    conn = sqlite3.connect(DB_PATH)
    try:
        tables = _existing_tables(conn)
        if not tables:
            print(f"[CLEAN-DB] {DB_PATH} has no tables.")
            return True
        print(f"[CLEAN-DB] Row counts in {DB_PATH}:")
        for t in tables:
            count = conn.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
            print(f"   - {t:<16} {count}")
    finally:
        conn.close()
    return True


def clear_tables() -> None:
    conn = sqlite3.connect(DB_PATH)
    try:
        tables = _existing_tables(conn)
        with conn:
            for t in tables:
                conn.execute(f"DELETE FROM {t}")
            # Reset AUTOINCREMENT counters if the sequence table exists
            if conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name='sqlite_sequence'"
            ).fetchone():
                conn.execute("DELETE FROM sqlite_sequence")
        conn.execute("VACUUM")
        print(f"[CLEAN-DB] Cleared all rows from {len(tables)} table(s): {', '.join(tables)}")
    finally:
        conn.close()


def delete_file() -> None:
    removed = []
    for suffix in ("", "-wal", "-shm", "-journal"):
        path = DB_PATH + suffix
        if os.path.exists(path):
            os.remove(path)
            removed.append(path)
    if removed:
        print(f"[CLEAN-DB] Deleted: {', '.join(removed)}")
    else:
        print(f"[CLEAN-DB] Nothing to delete (no file at {DB_PATH})")


def _confirm(action: str) -> bool:
    print(f"[CLEAN-DB] About to {action}")
    print(f"[CLEAN-DB] Target: {DB_PATH}")
    answer = input("[CLEAN-DB] Type 'yes' to continue: ").strip().lower()
    return answer in ("yes", "y")


def main() -> None:
    parser = argparse.ArgumentParser(description="Clean the NIS booking SQLite database.")
    parser.add_argument("--yes", action="store_true", help="Skip the confirmation prompt.")
    parser.add_argument("--delete-file", action="store_true",
                        help="Delete the .db file entirely (plus WAL/SHM) instead of clearing rows.")
    parser.add_argument("--show", action="store_true", help="Show current row counts and exit.")
    args = parser.parse_args()

    if args.show:
        show_counts()
        return

    if not os.path.exists(DB_PATH):
        print(f"[CLEAN-DB] No database file at: {DB_PATH} — nothing to do.")
        return

    action = "DELETE the database file" if args.delete_file else "DELETE ALL ROWS from the database"
    if not args.yes and not _confirm(action):
        print("[CLEAN-DB] Aborted.")
        sys.exit(1)

    if args.delete_file:
        delete_file()
    else:
        clear_tables()
    print("[CLEAN-DB] Done.")


if __name__ == "__main__":
    main()
