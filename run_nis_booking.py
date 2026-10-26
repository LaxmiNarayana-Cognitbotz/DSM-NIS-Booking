#!/usr/bin/env python3
"""
run_nis_booking.py  (repo-root entry point)
===========================================
Independent launcher for the **NIS booking** automation.

Responsibilities of this service ONLY:
  * read DSM emails, extract tables, classify amounts
  * NIS-book the positive (+ve) amounts
  * upload documents to SharePoint
  * update the Excel tracker and the shared SQLite database
  * seed the pending SAP / Debit-Credit work into the shared DB

It NEVER runs SAP checklist automation. The SAP side is a completely separate
process — see run_sap_automation.py.

Both services share ONE SQLite database (nis_booking.db). This launcher lives
outside the nis_booking/ package and puts it on sys.path so the module's flat
imports keep working.

Usage:
    python run_nis_booking.py --lookback 2
"""

import os
import sys

# ── Global SSL bypass (corporate proxy / self-signed cert in chain) ────────────
# Must happen BEFORE any other import so it covers requests, MSAL, and urllib3.
import ssl
ssl._create_default_https_context = ssl._create_unverified_context
os.environ.setdefault("PYTHONHTTPSVERIFY", "0")
os.environ.setdefault("CURL_CA_BUNDLE", "")
os.environ.setdefault("REQUESTS_CA_BUNDLE", "")

import urllib3 as _urllib3
_urllib3.disable_warnings(_urllib3.exceptions.InsecureRequestWarning)

import requests as _requests
_real_session_init = _requests.Session.__init__
def _patched_session_init(self, *args, **kwargs):
    _real_session_init(self, *args, **kwargs)
    self.verify = False
_requests.Session.__init__ = _patched_session_init
# ──────────────────────────────────────────────────────────────────────────────

ROOT = os.path.dirname(os.path.abspath(__file__))
NIS_DIR = os.path.join(ROOT, "nis_booking")

# Clear stale __pycache__ to ensure fresh modules are loaded after code updates.
import shutil
for _cache_dir in (
    os.path.join(NIS_DIR, "__pycache__"),
    os.path.join(ROOT, "sap_automation", "__pycache__"),
):
    if os.path.isdir(_cache_dir):
        shutil.rmtree(_cache_dir, ignore_errors=True)

# Put the module directory on sys.path so its flat imports resolve, and run
# with that as the working directory (keeps temp-book/ and logs/ relative paths
# identical to running nis_booking/main.py directly).
sys.path.insert(0, NIS_DIR)
os.chdir(NIS_DIR)

import main as nis_main_module  # noqa: E402


if __name__ == "__main__":
    nis_main_module.nis_main()
