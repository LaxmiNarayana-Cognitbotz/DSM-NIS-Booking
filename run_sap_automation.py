#!/usr/bin/env python3
"""
run_sap_automation.py  (repo-root entry point)
==============================================
Independent launcher for the **SAP automation** service.

Responsibilities of this service ONLY:
  * read the shared SQLite database (nis_booking.db) for pending SAP work
  * process the negative (-ve) SPV rows → SAP checklist creation
  * (Debit/Credit flow — staged) send the 5-row request email, monitor the
    reply, extract the reply PDFs and create their SAP checklists
  * upload generated checklist PDFs to SharePoint and record results in the
    shared DB

It runs completely independently of the NIS booking service and never triggers
it. Both services share ONE SQLite database.

sys.path is arranged so the sap_automation package's own modules take
precedence, with nis_booking/ appended only so the shared, standalone
``db_manager`` (the single source of truth) is importable.

Usage:
    python run_sap_automation.py                 # one-shot: process all pending
    python run_sap_automation.py --poll --interval 60
"""

import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
SAP_DIR = os.path.join(ROOT, "sap_automation")
NIS_DIR = os.path.join(ROOT, "nis_booking")

# Clear stale __pycache__ to ensure fresh modules are loaded after code updates.
import shutil
for _cache_dir in (
    os.path.join(NIS_DIR, "__pycache__"),
    os.path.join(SAP_DIR, "__pycache__"),
):
    if os.path.isdir(_cache_dir):
        shutil.rmtree(_cache_dir, ignore_errors=True)

# Order matters: sap_automation FIRST (its own copies of read_vendor_master_data,
# sharepoint_upload, etc. win), nis_booking appended ONLY so `db_manager` — the
# shared SQLite layer — resolves. db_manager is import-safe / standalone so it
# does not drag in nis_booking's other (duplicated) modules.
sys.path.insert(0, NIS_DIR)   # lower priority
sys.path.insert(0, SAP_DIR)   # higher priority (index 0)
os.chdir(SAP_DIR)

import sap_main  # noqa: E402


if __name__ == "__main__":
    sap_main.main()
