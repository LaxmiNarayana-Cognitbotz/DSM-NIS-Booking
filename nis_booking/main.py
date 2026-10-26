"""
main.py  (nis_booking)
======================
Email → Extract → Classify → Validate → Vendor → NIS Booking → Tracker Update.
SAP automation steps are SEEDED as "pending" but NOT executed here.
The sap_automation service handles those independently.

Usage:
    python main.py --lookback 2
"""

# ── Global SSL bypass (corporate proxy / self-signed cert in chain) ────────────
# Must happen BEFORE any other import so it covers requests, MSAL, and urllib3.
import ssl
import os as _os
ssl._create_default_https_context = ssl._create_unverified_context
_os.environ.setdefault("PYTHONHTTPSVERIFY", "0")
_os.environ.setdefault("CURL_CA_BUNDLE", "")
_os.environ.setdefault("REQUESTS_CA_BUNDLE", "")

import urllib3 as _urllib3
_urllib3.disable_warnings(_urllib3.exceptions.InsecureRequestWarning)

import requests as _requests
_requests.Session.__init__.__wrapped__ if hasattr(_requests.Session.__init__, "__wrapped__") else None
_real_session_init = _requests.Session.__init__
def _patched_session_init(self, *args, **kwargs):
    _real_session_init(self, *args, **kwargs)
    self.verify = False
_requests.Session.__init__ = _patched_session_init
# ──────────────────────────────────────────────────────────────────────────────

import argparse, html, json, os, re, shutil, sys, time, traceback, subprocess, threading
from datetime import datetime
from pathlib import Path
import requests

import db_manager as db
from read_email import fetch_emails_last_n_days, process_email, is_dsm_email, has_required_tables, send_html_email
from classified_html_table_parser import validate_mandatory_columns
from read_vendor_master_data import read_vendor_data, read_vendor_data_by_pdf_filename, spv_name_from_pdf_filename

# ── Compatibility shim: ensure STATUS_NIS_DONE exists on db_manager ──────────
# If the user deployed a new main.py but has an older db_manager.py, this
# prevents AttributeError crashes. The value is just a string — no schema change.
if not hasattr(db, "STATUS_NIS_DONE"):
    db.STATUS_NIS_DONE = "nis_done"
if not hasattr(db, "mark_booking_row_nis_done"):
    db.mark_booking_row_nis_done = lambda row_id: db.update_booking_row(
        row_id, {"status": "nis_done"}
    )
if not hasattr(db, "finalize_booking_row_if_ready"):
    # Stub: just return False — the SAP service will handle finalization
    db.finalize_booking_row_if_ready = lambda run_id, row_index, spv_name: False
if not hasattr(db, "are_all_sap_steps_done_for_row"):
    db.are_all_sap_steps_done_for_row = lambda run_id, row_index, spv_name: True
# ──────────────────────────────────────────────────────────────────────────────
from nis_booking import book_nis
from invoice_rules import is_email_pdf_spv
from add_new_row_data_nis import add_incremental_week_row, add_sap_checklist_reply_row
from pdf_merger import merge_pdfs
import sharepoint_upload as sp
import debit_credit_email as dce
import debit_credit_pdf_extractor as dcpdf
import finance_recipients_sp as finance_sp
import config
import log_manager as lm
import sys
sys.stdout.reconfigure(encoding='utf-8')


DEFAULT_LOOKBACK_DAYS = 2
MAX_STEP_RETRIES = 5

SPV_AMOUNT_FIELDS = [
    ("total_dsm_charges_payable", "Total DSM Charges Payable"),
    ("drawl_charges_payable",     "Drawl Charges Payable"),
    ("revenue_diff",              "Revenue Diff"),
    ("revenue_loss",              "Revenue Loss"),
]

TEMP_BOOKING_ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp-book")

# Serializes writes to the shared DSM-NIS-OUTPUT Excel tracker. The tracker
# update is a read-modify-write (read existing rows → dedup → find blank row →
# write). Parallel booking workers MUST NOT run it concurrently or they race —
# each reads before the others' writes land and ends up appending duplicate
# rows for the same email. This lock makes those updates one-at-a-time.
_TRACKER_LOCK = threading.Lock()

# Serializes writes to the run-level vendor_data_json field. Parallel booking
# workers each do a read-modify-write (read map → add their SPV → write back).
# Without this lock a later write can overwrite an earlier one, losing vendor
# data for other SPVs in the same email run.
_VENDOR_MAP_LOCK = threading.Lock()

NIS_CHECKLIST_SUBJECT = re.compile(r"NIS\s*Checklist\s*Print\s*-\s*(\d+)", re.IGNORECASE)
SAP_CHECKLIST_SUBJECT = re.compile(r"SAP\s*Checklist\s*Print\s*-\s*(\d+)", re.IGNORECASE)
CHECKLIST_NOT_MATCHED = "not_checklist"
CHECKLIST_PROCESSED = "processed"
CHECKLIST_SKIPPED = "skipped"

# ── HOD / FC Approval subject patterns ───────────────────────────────────────
HOD_APPROVAL_SUBJECT = re.compile(r"NIS\s*Approval\s*-\s*(\d+)", re.IGNORECASE)
FC_APPROVAL_SUBJECT = re.compile(r"NIS\s*Checklist\s*Print\s*-\s*(\d+)", re.IGNORECASE)
# HOD/FC handler outcomes
APPROVAL_NOT_MATCHED = "not_approval"
APPROVAL_PROCESSED = "processed"
APPROVAL_SKIPPED = "skipped"

# Debit/Credit Note finance-reply handling outcomes
DC_NOT_MATCHED = "not_debit_credit"
DC_PROCESSED = "processed"
DC_SKIPPED = "skipped"
_REPLY_PREFIX_RE = re.compile(r"^\s*(re|fw|fwd|aw|wg)\s*:\s*", re.IGNORECASE)


# ── Helpers ──────────────────────────────────────────────────────────────────
def safe_float(val):
    if val is None: return 0.0
    if isinstance(val, (int, float)): return float(val)
    try: return float(str(val).replace(",", "").strip())
    except: return 0.0

def _ts(): return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def _sanitize(name):
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    return re.sub(r"\s+", " ", name).strip()

def _spv_tmp_dir(spv_name):
    safe = _sanitize(spv_name) if spv_name else "unknown_spv"
    path = os.path.join(TEMP_BOOKING_ROOT, safe)
    os.makedirs(path, exist_ok=True)
    return path

def _row_tmp_dir(row_id):
    """Per-booking-row working directory.

    Each parallel worker gets its own folder (keyed by the unique row_id) so
    concurrent bookings never share PDFs or other files — their values can
    never overlap.
    """
    path = os.path.join(TEMP_BOOKING_ROOT, f"row_{row_id}")
    os.makedirs(path, exist_ok=True)
    return path

def _cleanup_tmp(path):
    try:
        if path and os.path.isdir(path):
            shutil.rmtree(path, ignore_errors=True)
    except: pass

def _message_tmp_dir(message_id):
    safe_id = re.sub(r"[^a-zA-Z0-9_-]", "_", str(message_id or ""))[:40]
    return os.path.join(TEMP_BOOKING_ROOT, safe_id) if safe_id else ""

def _is_truthy(value):
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    return str(value).strip().lower() in {"1", "true", "yes", "y", "done", "uploaded"}

def _extract_checklist_number(subject, pattern):
    match = pattern.search(subject or "")
    return match.group(1) if match else ""

def _collect_pdf_paths(email_result):
    pdfs = [p for p in (email_result.get("downloaded_files") or []) if str(p).lower().endswith(".pdf")]
    if not pdfs:
        attachment_dir = email_result.get("attachment_dir") or ""
        if attachment_dir and os.path.isdir(attachment_dir):
            pdfs = [str(p) for p in Path(attachment_dir).glob("*.pdf")]
    return pdfs


def _is_invoice_upload_pdf_name(name: str) -> bool:
    """True for the source invoice/DSM attachment PDF used as NIS header upload."""
    lower = str(name or "").strip().lower()
    if not lower.endswith(".pdf"):
        return False
    return not lower.startswith((
        "email",
        "sap_checklist_",
        "upload_",
        "nis_checklist_",
        "checklist_",
    ))


def _sorted_invoice_pdf_paths(paths) -> list:
    return sorted(
        [str(p) for p in (paths or []) if _is_invoice_upload_pdf_name(os.path.basename(str(p)))],
        key=lambda p: os.path.basename(p).lower(),
    )


def _forward_approval_email_to_recipients(msg, nis_id: str, approval_type: str):
    """Forward an approval email (HOD or FC/NIS Checklist) to the NIS
    completion recipients so they are notified when approvals come in.

    Args:
        msg: The raw Graph API message dict
        nis_id: The NIS checklist number extracted from the subject
        approval_type: Label like "HOD" or "FC (NIS Checklist)" for the notification
    """
    try:
        recipients = db.get_nis_completion_recipients()
        if not recipients:
            print(f"  [{approval_type}] Forward not sent: no recipients in nis_completion_recipients.emails")
            return

        original_subject = msg.get("subject", "")
        original_body = msg.get("body", {}).get("content", "")
        from_addr = ""
        from_obj = msg.get("from", {}).get("emailAddress", {})
        if from_obj:
            from_addr = from_obj.get("address", "")

        # Build forwarding notification
        fwd_subject = f"Fwd: {original_subject}"
        fwd_body = f"""<p><b>[DSM-NIS Automation]</b> {html.escape(approval_type)} approval received for NIS {html.escape(nis_id)}.</p>
<hr>
<p><b>Original Subject:</b> {html.escape(original_subject)}</p>
<p><b>From:</b> {html.escape(from_addr)}</p>
<p><b>Received:</b> {html.escape(msg.get('receivedDateTime', ''))}</p>
<hr>
{original_body if original_body else '<p>(no body)</p>'}
"""

        if send_html_email(fwd_subject, fwd_body, recipients):
            print(f"  [{approval_type}] ✅ Forwarded to {', '.join(recipients)}")
        else:
            print(f"  [{approval_type}] ⚠ Forward email send failed")
    except Exception as exc:
        print(f"  [{approval_type}] ⚠ Could not forward approval email: {exc}")


def _ensure_fc_approval_for_nis(nis_id: str, message_id: str = ""):
    """Mark FC approval for booking rows matching this NIS ID (if HOD is done).

    Called both when uploading the checklist PDF for the first time AND when
    re-encountering an already-uploaded checklist. Idempotent — skips rows
    already FC approved.
    """
    rows = db.get_booking_rows_by_nis_checklist_id(nis_id)
    for row in rows:
        row_id = row.get("row_id")
        if db.is_fc_approved(row_id):
            continue
        if row.get("hod_approval_status") != db.STATUS_DONE:
            continue
        db.mark_fc_approved(row_id, message_id)
        db.mark_fc_checklist_pdf_uploaded(row_id)
        print(f"  [FC] ✅ FC approved: row_id={row_id} SPV={row.get('spv_name')} NIS={nis_id}")
        _try_pdf_merge_for_row(row)

def _handle_checklist_email(msg, headers, base_url):
    subject = msg.get("subject", "")
    nis_id = _extract_checklist_number(subject, NIS_CHECKLIST_SUBJECT)
    sap_id = _extract_checklist_number(subject, SAP_CHECKLIST_SUBJECT)
    if not nis_id and not sap_id:
        return CHECKLIST_NOT_MATCHED

    run = None
    mark_uploaded = None
    pdf_name_key = ""
    checklist_label = ""
    uploaded_flag_key = ""
    if nis_id:
        run = db.get_run_by_nis_checklist_id(nis_id)
        checklist_label = f"NIS {nis_id}"
        uploaded_flag_key = "nis_checklist_pdf_uploaded"
        if run and _is_truthy(run.get(uploaded_flag_key)):
            print(f"  [CHECKLIST] NIS checklist PDF already uploaded for {nis_id}")
            # Even though PDF is already uploaded, ensure FC approval is marked
            # (this handles the case where checklist was uploaded before FC logic existed)
            _ensure_fc_approval_for_nis(nis_id, msg.get("id", ""))
            return CHECKLIST_SKIPPED
        mark_uploaded = db.mark_nis_checklist_pdf_uploaded
        pdf_name_key = "nis_checklist_pdf"
    elif sap_id:
        run = db.get_run_by_sap_checklist_value(sap_id)
        checklist_label = f"SAP {sap_id}"
        uploaded_flag_key = "sap_checklist_pdf_uploaded"
        if run and _is_truthy(run.get(uploaded_flag_key)):
            print(f"  [CHECKLIST] SAP checklist PDF already uploaded for {sap_id}")
            return CHECKLIST_SKIPPED
        mark_uploaded = db.mark_sap_checklist_pdf_uploaded
        pdf_name_key = "sap_checklist_pdf"

    if not run:
        print(f"  [CHECKLIST] No run found for {checklist_label} – skipping")
        return CHECKLIST_SKIPPED

    booking_folder = str(run.get("booking_folder_path") or "").strip()
    if not booking_folder:
        print(f"  [CHECKLIST] Missing booking folder for {checklist_label} – skipping")
        return CHECKLIST_SKIPPED

    # Collect ALL unique booking folders for rows sharing this checklist number.
    # Each SPV has its own folder — we upload the checklist PDF to ALL of them.
    all_booking_folders = set()
    all_booking_folders.add(booking_folder)
    if nis_id:
        for brow in db.get_booking_rows_by_nis_checklist_id(nis_id):
            bf = str(brow.get("booking_folder_path") or "").strip()
            if bf:
                all_booking_folders.add(bf)
    elif sap_id:
        run_id = str(run.get("id") or "")
        if run_id:
            for brow in db.get_booking_rows(run_id):
                bf = str(brow.get("booking_folder_path") or "").strip()
                if bf:
                    all_booking_folders.add(bf)

    email_result = {}
    try:
        email_result = process_email(msg, headers, base_url)
    except Exception as exc:
        print(f"  [CHECKLIST] Could not process checklist email for {checklist_label}: {exc}")
        _cleanup_tmp(_message_tmp_dir(msg.get("id")))
        return CHECKLIST_SKIPPED
    try:
        # Re-check after download in case another run uploaded it while this
        # process was working. Keep the local cleanup in the finally below.
        latest_run = db.get_run_by_id(run.get("id") or "") or run
        if uploaded_flag_key and _is_truthy(latest_run.get(uploaded_flag_key)):
            print(f"  [CHECKLIST] PDF became uploaded while processing {checklist_label} – skipping upload")
            return CHECKLIST_SKIPPED

        pdf_paths = _collect_pdf_paths(email_result)
        if not pdf_paths:
            print(f"  [CHECKLIST] No PDF attachments found for {checklist_label}")
            return CHECKLIST_SKIPPED

        uploaded_names = []
        for pdf_path in pdf_paths:
            try:
                pdf_name = os.path.basename(pdf_path)
                # Upload to ALL booking folders (each SPV has its own folder)
                for bf in all_booking_folders:
                    sp.upload_file_to_sp(pdf_path, bf, pdf_name)
                uploaded_names.append(pdf_name)
            except Exception as exc:
                print(f"  [CHECKLIST] Upload failed for {pdf_path}: {exc}")

        if mark_uploaded and uploaded_names:
            mark_uploaded(run.get("id") or "", True)
            if pdf_name_key:
                db.save_downloaded_pdf_name(run.get("id") or "", pdf_name_key, uploaded_names)
            print(f"  [CHECKLIST] Uploaded PDF(s) for {checklist_label}")

            # ── FC Approval: NIS Checklist Print IS the FC approval event ────
            # When a NIS checklist PDF is successfully uploaded, this also means
            # FC approval is done. Mark matching booking rows as FC approved and
            # trigger the PDF merge.
            if nis_id:
                _ensure_fc_approval_for_nis(nis_id, msg.get("id", ""))
                _forward_approval_email_to_recipients(msg, nis_id, "FC (NIS Checklist)")

            return CHECKLIST_PROCESSED

        return CHECKLIST_SKIPPED
    finally:
        _cleanup_tmp(email_result.get("attachment_dir", ""))

# ── HOD / FC Approval email handlers ──────────────────────────────────────────
def _handle_hod_approval_email(msg, headers, base_url):
    """Detect an HOD approval email (subject: 'NIS Approval - <nis_number>').

    When matched, mark the corresponding booking row(s) as HOD approved and
    forward the email to NIS completion recipients.
    Returns APPROVAL_NOT_MATCHED / APPROVAL_PROCESSED / APPROVAL_SKIPPED.
    """
    subject = msg.get("subject", "")
    match = HOD_APPROVAL_SUBJECT.search(subject)
    if not match:
        return APPROVAL_NOT_MATCHED

    nis_id = match.group(1)
    rows = db.get_booking_rows_by_nis_checklist_id(nis_id)
    if not rows:
        print(f"  [HOD] No booking row found for NIS {nis_id} — skipping")
        return APPROVAL_SKIPPED

    message_id = msg.get("id", "")
    processed_any = False
    for row in rows:
        row_id = row.get("row_id")
        if db.is_hod_approved(row_id):
            print(f"  [HOD] Already approved for row_id={row_id} NIS={nis_id}")
            continue
        # Only approve rows that are done (NIS + SAP complete)
        if row.get("status") != db.STATUS_DONE:
            print(f"  [HOD] Row {row_id} not yet done (status={row.get('status')}) — skipping HOD mark")
            continue
        db.mark_hod_approved(row_id, message_id)
        print(f"  [HOD] ✅ HOD approved: row_id={row_id} SPV={row.get('spv_name')} NIS={nis_id}")
        processed_any = True

    # Forward the HOD approval email to NIS completion recipients
    if processed_any:
        _forward_approval_email_to_recipients(msg, nis_id, "HOD")

    return APPROVAL_PROCESSED if processed_any else APPROVAL_SKIPPED


def _handle_fc_approval_email(msg, headers, base_url):
    """Detect an FC approval email (subject: 'NIS Checklist Print - <nis_number>').

    When matched:
      1. Mark the corresponding booking row(s) as FC approved
      2. Download the attached NIS checklist PDF
      3. Upload it to the row's booking folder

    Returns APPROVAL_NOT_MATCHED / APPROVAL_PROCESSED / APPROVAL_SKIPPED.
    """
    subject = msg.get("subject", "")
    match = FC_APPROVAL_SUBJECT.search(subject)
    if not match:
        return APPROVAL_NOT_MATCHED

    nis_id = match.group(1)
    rows = db.get_booking_rows_by_nis_checklist_id(nis_id)
    if not rows:
        print(f"  [FC] No booking row found for NIS {nis_id} — skipping")
        return APPROVAL_SKIPPED

    # Only process rows that already have HOD approval
    eligible_rows = [r for r in rows
                     if r.get("hod_approval_status") == db.STATUS_DONE
                     and r.get("fc_approval_status") != db.STATUS_DONE]
    if not eligible_rows:
        # Check if already FC approved
        if all(r.get("fc_approval_status") == db.STATUS_DONE for r in rows):
            print(f"  [FC] Already FC approved for NIS={nis_id}")
            return APPROVAL_SKIPPED
        print(f"  [FC] No eligible rows (HOD not done yet?) for NIS={nis_id}")
        return APPROVAL_SKIPPED

    message_id = msg.get("id", "")

    # Download the attached NIS checklist PDF
    email_result = {}
    try:
        email_result = process_email(msg, headers, base_url)
    except Exception as exc:
        print(f"  [FC] Could not process FC approval email for NIS {nis_id}: {exc}")
        _cleanup_tmp(_message_tmp_dir(msg.get("id")))
        return APPROVAL_SKIPPED

    try:
        pdf_paths = _collect_pdf_paths(email_result)

        for row in eligible_rows:
            row_id = row.get("row_id")
            booking_folder = str(row.get("booking_folder_path") or "").strip()

            # Mark FC approved
            db.mark_fc_approved(row_id, message_id)
            print(f"  [FC] ✅ FC approved: row_id={row_id} SPV={row.get('spv_name')} NIS={nis_id}")

            # Upload the NIS checklist PDF to the row's booking folder
            if booking_folder and pdf_paths:
                try:
                    sp.ensure_sp_folder(booking_folder)
                except Exception:
                    pass
                for pdf_path in pdf_paths:
                    pdf_name = os.path.basename(pdf_path)
                    # Ensure consistent naming
                    if not pdf_name.lower().startswith("nis_checklist_"):
                        pdf_name = f"nis_checklist_{nis_id}.pdf"
                    sp_path = sp.upload_file_to_sp(pdf_path, booking_folder, pdf_name)
                    if sp_path:
                        db.mark_fc_checklist_pdf_uploaded(row_id)
                        print(f"  [FC] Uploaded {pdf_name} → {booking_folder}")
                        break  # Only need one PDF per row

        # After FC approval, trigger PDF merge for eligible rows
        for row in eligible_rows:
            _try_pdf_merge_for_row(row)

        return APPROVAL_PROCESSED
    finally:
        _cleanup_tmp(email_result.get("attachment_dir", ""))


def _try_pdf_merge_for_row(row):
    """Attempt PDF merge for a booking row after FC approval.

    Downloads ALL documents from the row's booking folder, merges them into
    a single PDF, and uploads it back. Each SPV has its own folder, so each
    row's merge produces an independent merged PDF in its own folder.
    """
    row_id = row.get("row_id")
    if db.is_pdf_merged(row_id):
        return
    if not db.is_fc_approved(row_id):
        return

    booking_folder = str(row.get("booking_folder_path") or "").strip()
    if not booking_folder:
        print(f"  [PDF-MERGE] No booking folder for row_id={row_id} — skipping merge")
        return

    metadata = _json_dict(row.get("metadata_json"))
    to_date = str(metadata.get("to_date") or "").strip().replace(" ", "_").replace("/", "-")
    # Use the folder's SPV label (from folder path) for a consistent merged name
    # regardless of which row triggers it. E.g. "merged_AHEJ4L_24-May-2026.pdf"
    folder_spv = booking_folder.rstrip("/").split("/")[-2] if "/" in booking_folder else "unknown"
    safe_folder_spv = _sanitize(folder_spv) if folder_spv else "unknown"
    merged_name = f"merged_{safe_folder_spv}_{to_date}.pdf"

    try:
        # Check if merged PDF already exists in the folder (another row did it)
        sp_items = sp.list_sp_folder(booking_folder)
        existing_merged = [f for f in sp_items
                           if not f["is_folder"]
                           and f["name"].lower().startswith("merged_")
                           and f["name"].lower().endswith(".pdf")]
        if existing_merged:
            # Merged PDF already in folder — just mark this row as merged
            db.mark_pdf_merged(row_id, existing_merged[0]["name"])
            print(f"  [PDF-MERGE] Merged PDF already exists in folder — "
                  f"marking row_id={row_id} as merged ({existing_merged[0]['name']})")
            return

        # Collect ALL PDFs in the folder (except any old merged_ files)
        pdf_items = [f for f in sp_items
                     if not f["is_folder"]
                     and f["name"].lower().endswith(".pdf")
                     and not f["name"].lower().startswith("merged_")]

        if not pdf_items:
            print(f"  [PDF-MERGE] No PDFs found in {booking_folder} for row_id={row_id}")
            return

        # Download all PDFs to a temp directory
        work_dir = _row_tmp_dir(f"merge_{row_id}")
        downloaded_paths = []
        for item in pdf_items:
            local_path = os.path.join(work_dir, item["name"])
            dl = sp.download_file_from_sp(f"{booking_folder}/{item['name']}", local_path)
            if dl and os.path.isfile(dl):
                downloaded_paths.append(dl)

        if not downloaded_paths:
            print(f"  [PDF-MERGE] Could not download any PDFs for row_id={row_id}")
            _cleanup_tmp(work_dir)
            return

        # Sort: NIS checklist first, then SAP checklist, then invoice PDFs, then email.pdf
        def _sort_key(p):
            name = os.path.basename(p).lower()
            if name.startswith("nis_checklist"):
                return (0, name)
            elif name.startswith("sap_checklist"):
                return (1, name)
            elif name == "email.pdf":
                return (3, name)
            else:
                return (2, name)  # invoice/other PDFs (DSM_*.pdf)

        downloaded_paths.sort(key=_sort_key)

        # Merge all into one PDF
        merged_path = os.path.join(work_dir, merged_name)
        from pdf_merger import merge_pdfs
        merge_pdfs(downloaded_paths, merged_path)

        # Upload merged PDF
        if os.path.isfile(merged_path):
            sp.upload_file_to_sp(merged_path, booking_folder, merged_name)
            print(f"  [PDF-MERGE] ✅ Merged {len(downloaded_paths)} PDFs → {merged_name}")

            # Mark ALL rows sharing this folder as merged
            all_rows = db.get_booking_rows_by_nis_checklist_id(
                str(row.get("nis_checklist_id") or "").strip()
            )
            for r in all_rows:
                rid = r.get("row_id")
                if str(r.get("booking_folder_path") or "").strip() == booking_folder:
                    if not db.is_pdf_merged(rid):
                        db.mark_pdf_merged(rid, merged_name)
            # Also mark the current row (in case it wasn't caught above)
            if not db.is_pdf_merged(row_id):
                db.mark_pdf_merged(row_id, merged_name)
        else:
            print(f"  [PDF-MERGE] ❌ Merge output not found for row_id={row_id}")

        _cleanup_tmp(work_dir)
    except Exception as exc:
        print(f"  [PDF-MERGE] ❌ Failed for row_id={row_id}: {exc}")


# ── Debit/Credit Note flow (5-row DSM table trigger) ──────────────────────────
def _strip_reply_prefixes(subject):
    s = str(subject or "")
    prev = None
    while prev != s:
        prev = s
        s = _REPLY_PREFIX_RE.sub("", s, count=1)
    return s.strip()

def _normalize_subject(subject):
    return re.sub(r"\s+", " ", _strip_reply_prefixes(subject)).strip().lower()

def _looks_like_reply(subject):
    return bool(_REPLY_PREFIX_RE.match(str(subject or "")))

def _maybe_register_debit_credit_email(run_id, metadata_row, spv_dsm_rows, booking_folder):
    """If the SPV table has the trigger row-count, register (idempotently) a
    pending debit/credit-email record. Sending happens later via
    ``send_ready_debit_credit_emails`` so it is retryable and fully tracked.
    """
    if not getattr(config, "DEBIT_CREDIT_ENABLED", True):
        return
    try:
        if not dce.should_send(spv_dsm_rows):
            return
        run = db.get_run_by_id(run_id) or {}
        message_id = str(run.get("message_id") or run_id)
        info = dce.build_email(metadata_row, spv_dsm_rows)
        valid_rows = dce.valid_spv_rows(spv_dsm_rows)
        spv_name = str((valid_rows[0].get("spv_name") if valid_rows else "") or "").strip()
        dc_id = db.register_debit_credit_email(run_id, message_id, {
            "spv_name": spv_name,
            "week_no": info.get("week_no", ""),
            "from_date": info.get("from_date", ""),
            "to_date": info.get("to_date", ""),
            "dsm_amount": info.get("dsm_amount", ""),
            "company_code": info.get("company_code", ""),
            "due_date_cn_dn": info.get("due_date_cn_dn", ""),
            "booking_folder_path": str(booking_folder or ""),
        })
        print(f"  [DEBIT-CREDIT] Registered debit/credit email "
              f"(id={dc_id}) for {len(valid_rows)}-row table — week {info.get('week_no')}")
    except Exception as exc:
        print(f"  [DEBIT-CREDIT] Could not register debit/credit email: {exc}")

def _send_one_debit_credit_email(record):
    dc_id = record.get("id")
    run_id = str(record.get("run_id") or "")
    try:
        recipients = finance_sp.get_finance_recipients()
        if not recipients:
            print("  [DEBIT-CREDIT] No finance_team recipients found — cannot send (sheet 'finance_team', column 'emails').")
            db.mark_debit_credit_email_failed(dc_id, "no finance_team recipients")
            return False

        run = db.get_run_by_id(run_id) or {}
        metadata_row = _json_dict(run.get("metadata_json"))
        spv_rows = _json_list(run.get("spv_dsm_json"))
        if not spv_rows:
            db.mark_debit_credit_email_failed(dc_id, "no spv rows available to build email")
            return False

        subject = dce.build_subject(metadata_row, spv_rows)
        body_html = dce.build_html(metadata_row, spv_rows)

        if send_html_email(subject, body_html, recipients):
            db.mark_debit_credit_email_sent(dc_id, subject, recipients)
            print(f"  [DEBIT-CREDIT] Sent debit/credit email to {', '.join(recipients)}")
            return True
        db.mark_debit_credit_email_failed(dc_id, "send_html_email returned False")
        return False
    except Exception as exc:
        print(f"  [DEBIT-CREDIT] Send failed for id={dc_id}: {exc}")
        try:
            db.mark_debit_credit_email_failed(dc_id, str(exc))
        except Exception:
            pass
        return False

def send_ready_debit_credit_emails():
    """Send debit/credit emails for runs whose booking work is fully complete.

    A 5-row email's debit/credit note is only sent AFTER every SPV row has been
    booked (NIS for +ve amounts) and all SAP checklist steps (-ve amounts) are
    complete. Until then the record stays pending and is retried next loop.
    """
    if not getattr(config, "DEBIT_CREDIT_ENABLED", True):
        return 0
    sent = 0
    for record in db.get_pending_debit_credit_emails():
        run_id = str(record.get("run_id") or "")
        if not _is_run_ready_for_debit_credit(run_id):
            print(f"  [DEBIT-CREDIT] Holding email for run_id={run_id[:30]} "
                  f"— NIS/SAP booking not complete yet.")
            continue
        if _send_one_debit_credit_email(record):
            sent += 1
    return sent

def _is_run_ready_for_debit_credit(run_id):
    """True only when ALL booking work for the run has finished: every SPV
    booking row is fully done (not nis_done/pending/failed) or skipped, and
    every SAP checklist step (-ve) is complete. This enforces the order
    'book NIS + SAP checklist first, then debit/credit'. Rows in nis_done
    state (awaiting SAP) will block this gate until SAP finalizes them."""
    if not run_id:
        return False
    run = db.get_run_by_id(run_id)
    if not run:
        return False
    if not db.are_all_booking_rows_done(run_id):
        return False
    if not db.are_all_sap_steps_complete(run_id):
        return False
    return True

def _find_debit_credit_record_for_reply(subject):
    """Match an incoming reply subject to a sent debit/credit record."""
    norm_in = _normalize_subject(subject)
    if not norm_in:
        return None
    for rec in db.get_debit_credit_emails_awaiting_reply():
        sent_subject = _normalize_subject(rec.get("sent_subject"))
        if sent_subject and sent_subject in norm_in:
            return rec
        # Fallback: keyword + week + period tokens all present
        wk = str(rec.get("week_no") or "").strip().lower()
        frm = str(rec.get("from_date") or "").strip().lower()
        to = str(rec.get("to_date") or "").strip().lower()
        if wk and frm and to and wk in norm_in and frm in norm_in and to in norm_in:
            return rec
    return None

def _handle_debit_credit_reply(msg, headers, base_url):
    """Detect a finance-team reply to a debit/credit email and store its PDFs.

    Returns DC_NOT_MATCHED / DC_PROCESSED / DC_SKIPPED.
    """
    if not getattr(config, "DEBIT_CREDIT_ENABLED", True):
        return DC_NOT_MATCHED
    subject = msg.get("subject", "")
    keyword = str(getattr(config, "DEBIT_CREDIT_SUBJECT_KEYWORD", "Debit/Credit Notes")).lower()
    if keyword not in (subject or "").lower():
        return DC_NOT_MATCHED
    # Only replies/forwards carry the finance PDFs — this also excludes our own
    # outgoing copy (saved to Sent Items, no reply prefix).
    if not _looks_like_reply(subject):
        return DC_NOT_MATCHED

    record = _find_debit_credit_record_for_reply(subject)
    if not record:
        return DC_NOT_MATCHED
    if _is_truthy(record.get("reply_pdfs_uploaded")):
        print(f"  [DEBIT-CREDIT] Reply PDFs already stored for id={record.get('id')}")
        return DC_SKIPPED

    booking_folder = str(record.get("booking_folder_path") or "").strip()
    if not booking_folder:
        run = db.get_run_by_id(str(record.get("run_id") or "")) or {}
        booking_folder = str(run.get("booking_folder_path") or "").strip()
    if not booking_folder:
        print(f"  [DEBIT-CREDIT] No booking folder for id={record.get('id')} — skipping reply")
        return DC_SKIPPED

    email_result = {}
    try:
        email_result = process_email(msg, headers, base_url)
    except Exception as exc:
        print(f"  [DEBIT-CREDIT] Could not process reply email: {exc}")
        _cleanup_tmp(_message_tmp_dir(msg.get("id")))
        return DC_SKIPPED

    try:
        pdf_paths = _collect_pdf_paths(email_result)
        if not pdf_paths:
            print(f"  [DEBIT-CREDIT] Reply has no PDF attachments yet (id={record.get('id')})")
            return DC_SKIPPED

        target_folder = f"{booking_folder}/{getattr(config, 'DEBIT_CREDIT_REPLY_SUBFOLDER', 'Debit-Credit-Notes')}"
        try:
            sp.ensure_sp_folder(target_folder)
        except Exception as exc:
            print(f"  [DEBIT-CREDIT] Could not ensure folder {target_folder}: {exc}")

        uploaded_names = []
        for pdf_path in pdf_paths:
            try:
                pdf_name = os.path.basename(pdf_path)
                # process_email prefixes downloads with "DSM_"; keep the
                # finance note's original name for stored Debit/Credit PDFs.
                if pdf_name.startswith("DSM_"):
                    pdf_name = pdf_name[len("DSM_"):]
                sp_path = sp.upload_file_to_sp(pdf_path, target_folder, pdf_name)
                if sp_path:
                    uploaded_names.append(pdf_name)
                # Extract note data + resolve vendor, and persist a per-PDF row.
                _register_reply_pdf(record, pdf_path, pdf_name)
            except Exception as exc:
                print(f"  [DEBIT-CREDIT] Upload/extract failed for {pdf_path}: {exc}")

        if uploaded_names:
            db.mark_debit_credit_reply_received(
                record.get("id"), msg.get("id", ""), uploaded_names,
                uploaded=True, received_at=email_result.get("received_at", ""),
            )
            print(f"  [DEBIT-CREDIT] Stored {len(uploaded_names)} reply PDF(s) for id={record.get('id')}: "
                  f"{', '.join(uploaded_names)}")
            return DC_PROCESSED
        return DC_SKIPPED
    finally:
        _cleanup_tmp(email_result.get("attachment_dir", ""))

def _register_reply_pdf(record, pdf_path, pdf_name):
    """Extract a reply PDF's fields, resolve its vendor from DSM-NIS-INPUT by
    keyword-matching the filename SPV, and upsert a debit_credit_pdfs row.

    Best-effort: extraction/vendor errors are logged and stored on the row so
    the SAP step can be retried, but they never block the reply upload."""
    dc_id = record.get("id")
    fields = {
        "workflow_id": str(record.get("workflow_id") or ""),
        "run_id": str(record.get("run_id") or ""),
        "reply_message_id": str(record.get("reply_message_id") or ""),
        "spv_from_filename": spv_name_from_pdf_filename(pdf_name),
        "from_date": str(record.get("from_date") or ""),
        "to_date": str(record.get("to_date") or ""),
    }
    errors = []

    # 1) Extract ref / amount / date from the PDF text.
    try:
        data = dcpdf.extract_dc_pdf_data(pdf_path)
        fields.update({
            "ref_number": data.get("ref_number", ""),
            "amount": data.get("amount", 0.0),
            "amount_text": data.get("amount_text", ""),
            "doc_date": data.get("doc_date", ""),
            "note_type": data.get("note_type", ""),
            "dsm_week": data.get("dsm_week", ""),
        })
        fields["extract_status"] = db.STATUS_DONE if data.get("ref_number") else db.STATUS_FAILED
        if not data.get("ref_number"):
            errors.append("ref_number not found in PDF")
    except Exception as exc:
        fields["extract_status"] = db.STATUS_FAILED
        errors.append(f"extract: {exc}")

    # 2) Resolve vendor / company / cost-center / plant from DSM-NIS-INPUT.
    try:
        vendor = read_vendor_data_by_pdf_filename(pdf_name)
        if vendor:
            fields.update({
                "matched_spv_name": vendor.get("Matched SPV", ""),
                "vendor_code": vendor.get("Vendor", ""),
                "company_code": vendor.get("Company", ""),
                "cost_center": vendor.get("Cost Center", ""),
                "plant": vendor.get("Plant", ""),
            })
        else:
            errors.append("no SPV match in DSM-NIS-INPUT")
    except Exception as exc:
        errors.append(f"vendor: {exc}")

    if errors:
        fields["error_message"] = "; ".join(errors)[:2000]
    db.upsert_debit_credit_pdf(dc_id, pdf_name, fields)
    print(f"  [DEBIT-CREDIT] PDF tracked: {pdf_name} → spv={fields.get('matched_spv_name') or fields.get('spv_from_filename')} "
          f"vendor={fields.get('vendor_code','')} ref={fields.get('ref_number','')} amount={fields.get('amount',0)}")

def _import_run_sap_automation():
    """Lazily import run_sap_automation from the sibling sap_automation package.

    Kept lazy so the nis_booking service imports fine on non-Windows hosts;
    the actual SAP GUI automation only loads on the Windows machine that has
    win32com + SAP GUI available.
    """
    sap_dir = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "sap_automation"
    )
    if sap_dir not in sys.path:
        sys.path.insert(0, sap_dir)
    from sap_automation import run_sap_automation  # noqa: E402
    return run_sap_automation

def _write_debit_credit_output_row(pdf_row: dict, checklist_number: str):
    """Write a DSM-NIS-OUTPUT row for a finance-reply Debit/Credit SAP checklist.

    The row goes on the matched SPV's sheet, carries NO NIS booking number, and
    stores the SAP checklist number in column K. Shared header fields (week,
    dates, due date) are copied from the parent debit/credit email so the row
    belongs to the original DSM email. Best-effort: failures are logged but do
    not fail the SAP checklist itself (it is already recorded in the DB)."""
    checklist_number = str(checklist_number or "").strip()
    if not checklist_number:
        return
    try:
        sheet_name = (str(pdf_row.get("matched_spv_name") or "").strip()
                      or str(pdf_row.get("spv_from_filename") or "").strip())
        if not sheet_name:
            print("  [DC-SAP] No SPV name for reply row — skipping OUTPUT write.")
            return

        parent = db.get_debit_credit_email_by_id(pdf_row.get("dc_email_id")) or {}
        metadata_row = {
            "week_no":   str(parent.get("week_no") or pdf_row.get("dsm_week") or "").strip(),
            "from_date": str(pdf_row.get("from_date") or parent.get("from_date") or "").strip(),
            "to_date":   str(pdf_row.get("to_date") or parent.get("to_date") or "").strip(),
            "due_date":  str(parent.get("due_date_cn_dn") or "").strip(),
            "dsm_statement_published_date": str(parent.get("dsm_statement_published_date") or "").strip(),
        }
        cn_dn_date = _format_dc_date(pdf_row.get("doc_date")) or datetime.utcnow().strftime("%d-%b-%Y")

        with _TRACKER_LOCK:
            add_sap_checklist_reply_row(
                metadata_row=metadata_row,
                sheet_name=sheet_name,
                sap_checklist_number=checklist_number,
                cn_dn_date=cn_dn_date,
            )
    except Exception as exc:
        print(f"  [DC-SAP] Could not write OUTPUT reply row (checklist {checklist_number}): {exc}")


def _format_dc_date(raw) -> str:
    """Format a debit/credit doc date to the tracker's DD-Mon-YYYY, best-effort."""
    raw = str(raw or "").strip()
    if not raw:
        return ""
    for fmt in ("%d.%m.%Y", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d",
                "%d-%b-%Y", "%d %b %Y", "%d %B %Y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%d-%b-%Y")
        except Exception:
            continue
    return raw


def process_debit_credit_sap_checklists():
    """Create a SAP checklist for each finance-reply Debit/Credit PDF that has
    its extracted values, then store the checklist number and upload the
    generated checklist PDF to booking/<from_date>_<to_date>/ on SharePoint.

    Uses vendor/company/cost-center/plant resolved from DSM-NIS-INPUT (by the
    filename SPV) and ref/amount/date extracted from the PDF. Returns the number
    of checklists completed this pass. Rows missing data or unable to run SAP
    stay pending/failed for the next loop (retryable).
    """
    if not getattr(config, "DEBIT_CREDIT_ENABLED", True):
        return 0
    if not getattr(config, "DEBIT_CREDIT_SAP_ENABLED", True):
        return 0
    pending = db.get_pending_sap_debit_credit_pdfs()
    if not pending:
        return 0

    run_sap = None
    done = 0
    for row in pending:
        pdf_id = row.get("id")
        # Validate the values SAP needs before attempting a checklist.
        missing = [f for f in ("vendor_code", "company_code", "ref_number", "doc_date")
                   if not str(row.get(f) or "").strip()]
        if safe_float(row.get("amount")) <= 0:
            missing.append("amount")
        if missing:
            db.mark_debit_credit_pdf_sap_failed(pdf_id, "missing values: " + ", ".join(missing))
            print(f"  [DC-SAP] PDF id={pdf_id} not ready — missing {missing}")
            continue

        if run_sap is None:
            try:
                run_sap = _import_run_sap_automation()
            except Exception as exc:
                print(f"  [DC-SAP] SAP automation unavailable on this host ({exc}) — "
                      f"leaving {len(pending)} PDF(s) pending for retry.")
                return done

        try:
            folder = sp.build_sap_checklist_folder_path(row.get("from_date"), row.get("to_date"))
            try:
                sp.ensure_sp_folder(folder)
            except Exception as exc:
                print(f"  [DC-SAP] Could not ensure folder {folder}: {exc}")

            out_dir = os.path.join(TEMP_BOOKING_ROOT, f"dc_sap_{pdf_id}")
            os.makedirs(out_dir, exist_ok=True)

            print(f"  [DC-SAP] Checklist for {row.get('pdf_name')} "
                  f"(vendor={row.get('vendor_code')} company={row.get('company_code')} "
                  f"ref={row.get('ref_number')} amount={row.get('amount')} date={row.get('doc_date')})")

            # Strip trailing "-DSM" from ref_number — SAP expects just the
            # note reference (e.g. "CN/26-27/308") without the DSM suffix that
            # appears in the finance-reply PDF text.
            _raw_ref = str(row.get("ref_number") or "").strip()
            if _raw_ref.upper().endswith("-DSM"):
                _raw_ref = _raw_ref[:-4].strip()

            checklist_number, sap_pdf_path = run_sap(
                company_code=str(row.get("company_code") or ""),
                vendor_code=str(row.get("vendor_code") or ""),
                amount=safe_float(row.get("amount")),
                cost_center=str(row.get("cost_center") or ""),
                plant=str(row.get("plant") or ""),
                deb_not_date=str(row.get("doc_date") or ""),
                deb_not_ref=_raw_ref,
                pdf_output_dir=out_dir,
            )

            sap_pdf_name, sap_sp_path = "", ""
            if sap_pdf_path and os.path.isfile(sap_pdf_path):
                sap_pdf_name = os.path.basename(sap_pdf_path)
                sap_sp_path = sp.upload_file_to_sp(sap_pdf_path, folder, sap_pdf_name)

            db.update_debit_credit_pdf(pdf_id, {"sap_folder_path": folder})

            # Only mark as done if SAP actually saved successfully
            # (at minimum we need either a checklist number OR a PDF file)
            if not checklist_number and not sap_pdf_name:
                db.mark_debit_credit_pdf_sap_failed(
                    pdf_id,
                    "SAP save failed: no checklist number captured and no PDF generated. "
                    "Check if all required SAP fields are filled correctly."
                )
                print(f"  [DC-SAP] ❌ id={pdf_id} SAP save failed — no checklist and no PDF. "
                      f"NOT marking as done.")
            else:
                db.mark_debit_credit_pdf_sap_done(pdf_id, checklist_number or "", sap_pdf_name, sap_sp_path)
                print(f"  [DC-SAP] ✅ id={pdf_id} checklist={checklist_number or 'n/a'} pdf={sap_pdf_name or 'n/a'}")
                _write_debit_credit_output_row(row, checklist_number or "")
                done += 1
            _cleanup_tmp(out_dir)
        except Exception as exc:
            db.mark_debit_credit_pdf_sap_failed(pdf_id, str(exc))
            print(f"  [DC-SAP] ❌ id={pdf_id} failed: {exc}")

    return done

def normalize_vendor_data(raw):
    out = {}
    for k, v in raw.items():
        key = k.strip().upper()
        if key in ("VENDOR", "VENDOR CODE"): out["vendor_code"] = str(v)
        elif key in ("COMPANY", "COMPANY CODE"): out["company_code"] = str(v)
        elif key == "COST CENTER": out["cost_center"] = str(v)
        elif key == "PLANT": out["plant"] = str(v)
        elif key == "BANK KEY": out["bank_key"] = str(v)
        elif key == "GL": out["gl_account"] = str(v)
        elif key == "PURPOSE": out["purpose"] = str(v)
    return out

def _format_invoice_date(raw):
    from datetime import datetime as _dt
    raw = str(raw or "").strip()
    if not raw: return raw
    for fmt in ["%d.%m.%Y","%d-%m-%Y","%d/%m/%Y","%Y-%m-%d","%d-%b-%y","%d-%b-%Y","%d %b %Y","%d %b %y","%d %B %Y"]:
        try: return _dt.strptime(raw, fmt).strftime("%d.%m.%Y")
        except: continue
    return raw

def _has_all_extracted_values(metadata_row, spv_dsm_rows):
    for f in ("week_no","from_date","to_date","due_date"):
        if not str(metadata_row.get(f,"")).strip(): return False
    if not spv_dsm_rows: return False
    return any(str(r.get("spv_name","")).strip() for r in spv_dsm_rows)

# ── Failure email (ONE per email/run, aggregating every SPV row) ─────────────
# Required booking fields — used to point out exactly which value is missing.
_FAILURE_REQUIRED_FIELDS = [
    ("vendor_data_json",      "Vendor master data (vendor code / company code / cost center / plant)"),
    ("invoice_date",          "Invoice date"),
    ("booking_folder_path",   "SharePoint booking folder"),
    ("positive_fields_json",  "Charge amount classification"),
    ("metadata_json",         "Email metadata (week / dates / due date)"),
    ("spv_data_json",         "SPV charge values"),
]


def _row_missing_fields(brow: dict) -> list:
    """Return the human labels of required booking fields that are empty on *brow*."""
    missing = []
    for field, label in _FAILURE_REQUIRED_FIELDS:
        val = (brow or {}).get(field)
        if field == "vendor_data_json":
            if not _json_dict(val).get("vendor_code"):
                missing.append(label)
        elif field in ("positive_fields_json", "metadata_json", "spv_data_json"):
            if not _json_dict(val):
                missing.append(label)
        else:
            if not str(val or "").strip():
                missing.append(label)
    return missing


def _collect_run_failure_details(run_id):
    """Gather a per-row status snapshot for every SPV row tried in this email.

    Returns ``(rows, any_failed)`` where each entry is a dict with row_index,
    spv_name, status, error_message and the list of missing required fields.
    """
    details = []
    any_failed = False
    for brow in db.get_booking_rows(run_id):
        status = str(brow.get("status") or "").strip() or "pending"
        if status == db.STATUS_FAILED:
            any_failed = True
        details.append({
            "row_index": brow.get("row_index"),
            "spv_name": str(brow.get("spv_name") or "").strip() or "(unnamed SPV)",
            "status": status,
            "error_message": str(brow.get("error_message") or "").strip(),
            "missing_fields": _row_missing_fields(brow),
        })
    return details, any_failed


def _llm_explain_failure(spv_name, from_date, to_date, step_name, row_details, error_message):
    """Ask the LLM for a short, business-friendly explanation of the failure.

    Focuses on WHICH column/value caused the problem (value not found vs. value
    mismatch) in plain language a business user can act on. Falls back to a
    simple rule-based summary if the LLM is unavailable.
    """
    # Build a compact, factual context for the model.
    lines = [
        f"Email SPV: {spv_name or 'unknown'}",
        f"Period: {from_date or '?'} to {to_date or '?'}",
        f"Failed step: {step_name or 'unknown'}",
        "",
        "Per-row outcomes:",
    ]
    if row_details:
        for d in row_details:
            lines.append(
                f"- Row {d['row_index']} | SPV '{d['spv_name']}' | status={d['status']}"
            )
            if d.get("missing_fields"):
                lines.append(f"    Missing values: {', '.join(d['missing_fields'])}")
            if d.get("error_message"):
                lines.append(f"    Error: {d['error_message'][:300]}")
    else:
        lines.append(f"- (no per-row data) Error: {str(error_message or '')[:400]}")
    context = "\n".join(lines)

    system_prompt = (
        "You are an assistant that explains automation failures to a NON-technical "
        "business/finance user. Given the raw per-row status of a DSM-NIS booking "
        "email, write a SHORT explanation (3-6 sentences, plain English, no code, "
        "no stack traces). Clearly state, for each affected SPV row: which column / "
        "value is the problem, and whether the value was NOT FOUND (missing) or is a "
        "MISMATCH (does not match the vendor master / expected value). If several rows "
        "share the same issue, group them. End with a one-line suggested next action. "
        "Do not invent details that are not present in the input."
    )
    try:
        from llm_curl_client import call_llm_curl
        reply = call_llm_curl(
            [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": context},
            ],
            max_tokens=350,
            timeout=60,
        )
        if reply and reply.strip():
            return reply.strip()
    except Exception as exc:
        print(f"  [EMAIL] LLM failure-explanation unavailable: {exc}")

    # ── Fallback (no LLM): simple rule-based summary ─────────────────────────
    parts = []
    for d in row_details or []:
        if d["status"] == db.STATUS_FAILED or d.get("missing_fields"):
            bits = []
            if d.get("missing_fields"):
                bits.append("missing " + ", ".join(d["missing_fields"]))
            if d.get("error_message"):
                bits.append(d["error_message"][:160])
            parts.append(f"Row {d['row_index']} ({d['spv_name']}): " + "; ".join(bits))
    if not parts and error_message:
        parts.append(str(error_message)[:300])
    return " | ".join(parts) or "The automation could not be completed for this email."


def _send_failure_email(run_id, spv_name, from_date, to_date, step_name, error_message):
    """Send ONE failure email per email/run, listing every SPV row that was tried
    with its status, plus a plain-language (LLM-generated) explanation of which
    column/value caused the problem (value not found vs. value mismatch)."""
    if db.is_failure_email_sent(run_id):
        return
    recipients = db.get_nis_completion_recipients()
    if not recipients:
        print("  [EMAIL] Failure email not sent: no recipients in nis_completion_recipients.emails")
        return
    try:
        import msal
        app = msal.ConfidentialClientApplication(config.GRAPH_CLIENT_ID, authority=config.GRAPH_AUTHORITY, client_credential=config.GRAPH_CLIENT_SECRET)
        result = app.acquire_token_for_client(scopes=config.GRAPH_SCOPE)
        token = result.get("access_token")
        if not token:
            return

        row_details, _ = _collect_run_failure_details(run_id)
        explanation = _llm_explain_failure(
            spv_name, from_date, to_date, step_name, row_details, error_message
        )

        def esc(v):
            return html.escape(str(v if v is not None else ""))

        # Per-row status table (all tried rows + their status/issue).
        if row_details:
            row_html = "".join(
                "<tr>"
                f"<td>{esc(d['row_index'])}</td>"
                f"<td>{esc(d['spv_name'])}</td>"
                f"<td>{esc(d['status'].upper())}</td>"
                f"<td>{esc('; '.join(d['missing_fields']) if d['missing_fields'] else '')}</td>"
                f"<td>{esc(d['error_message'][:300])}</td>"
                "</tr>"
                for d in row_details
            )
            rows_table = (
                "<table border='1' cellspacing='0' cellpadding='6'>"
                "<tr style='background:#f2f2f2'>"
                "<th>Row</th><th>SPV</th><th>Status</th>"
                "<th>Missing value(s)</th><th>Details</th></tr>"
                f"{row_html}</table>"
            )
        else:
            rows_table = f"<pre>{esc(str(error_message)[:1000])}</pre>"

        subject = f"[DSM-NIS Automation FAILED] {spv_name} | {from_date} to {to_date}"
        body = (
            "<h3>Automation Failure</h3>"
            f"<p><b>SPV / Email:</b> {esc(spv_name)}</p>"
            f"<p><b>Period:</b> {esc(from_date)} to {esc(to_date)}</p>"
            f"<p><b>Step:</b> {esc(step_name)}</p>"
            "<h4>What went wrong (summary)</h4>"
            f"<div style='padding:8px;background:#fff8e1;border:1px solid #ffe082'>"
            f"{esc(explanation).replace(chr(10), '<br>')}</div>"
            "<h4>All rows in this email</h4>"
            f"{rows_table}"
            f"<p style='color:#888;font-size:11px'>Run ID: {esc(run_id)}</p>"
        )
        resp = requests.post(f"https://graph.microsoft.com/v1.0/users/{config.USER_EMAIL}/sendMail",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"message": {"subject": subject, "body": {"contentType": "HTML", "content": body}, "toRecipients": [{"emailAddress": {"address": addr}} for addr in recipients]}, "saveToSentItems": "true"}, timeout=30)
        if resp.status_code in (200, 202):
            db.mark_failure_email_sent(run_id)
            print(f"  [EMAIL] Failure notification sent to {', '.join(recipients)}")
        else:
            print(f"  [EMAIL] Failure notification HTTP {resp.status_code}: {resp.text[:200]}")
    except Exception as exc:
        print(f"  [EMAIL] Could not send failure notification: {exc}")


def _json_dict(raw):
    try:
        value = json.loads(raw or "{}")
        return value if isinstance(value, dict) else {}
    except Exception:
        return {}


def _json_list(raw):
    try:
        value = json.loads(raw or "[]")
        return value if isinstance(value, list) else []
    except Exception:
        return []


def _normalized_key(value):
    return re.sub(r"\s+", " ", str(value or "").strip()).upper()


def _is_total_spv_row(row):
    return _normalized_key((row or {}).get("spv_name")) == "TOTAL"


def _valid_spv_rows(rows):
    valid = []
    for row in rows or []:
        if not isinstance(row, dict) or _is_total_spv_row(row):
            continue
        if str(row.get("spv_name") or "").strip():
            valid.append(row)
    return valid


def _classify_amounts(spv_row, emit=False):
    positive, negative = {}, {}
    for fk, label in SPV_AMOUNT_FIELDS:
        amount = safe_float((spv_row or {}).get(fk))
        if amount > 0:
            positive[fk] = amount
            if emit: print(f"    [+] {label} = {amount}  → NIS booking")
        elif amount < 0:
            negative[fk] = abs(amount)
            if emit: print(f"    [-] {label} = {amount}  → SAP automation (pending)")
        else:
            if emit: print(f"    [0] {label} = 0 / blank  → skipped")
    return positive, negative


def _sum_amount_maps(items):
    merged = {}
    for amount_map in items:
        for key, value in (amount_map or {}).items():
            merged[key] = merged.get(key, 0.0) + safe_float(value)
    return merged


def _row_step_name(base_step, row_index, spv_name):
    safe_spv = re.sub(r"[^A-Za-z0-9]+", "_", str(spv_name or "spv")).strip("_")[:40] or "spv"
    return f"{base_step}_row_{row_index}_{safe_spv}"


# ── SPV folder naming and cost-center booking grouping ───────────────────────
# Wind/Solar variants of one plant continue to share a SharePoint folder named
# after the base plant code, e.g. booking/AHEJ4L/. NIS booking groups are formed
# separately from the persisted vendor cost center.
_WIND_SOLAR_RE = re.compile(r"^\s*([A-Za-z0-9]+)\s*-\s*(Solar|Wind)\b", re.IGNORECASE)

# Display labels used when each positive amount becomes an NIS expense item.
NIS_ITEM_LABELS = {
    "total_dsm_charges_payable": "Total DSM charges payable",
    "drawl_charges_payable":     "Drawl charges payable",
    "revenue_diff":              "Revenue difference",
    "revenue_loss":              "Revenue loss",
}


def _spv_base_and_kind(spv_name):
    """Return (base_code, kind) for a Wind/Solar SPV, else ("", "").

    base_code is the leading plant code (e.g. "AHEJ4L"); kind is "solar"/"wind".
    """
    match = _WIND_SOLAR_RE.match(str(spv_name or ""))
    if not match:
        return "", ""
    return match.group(1).upper(), match.group(2).lower()


def _folder_spv_name(spv_name):
    """Folder label for a booking. Each SPV ALWAYS gets its own folder named
    after its full SPV name — even Wind/Solar variants of the same plant.
    Documents (email.pdf, invoice, checklists) are duplicated to each folder
    so every SPV has a complete, independent set of files.
    """
    return str(spv_name or "").strip()


def _booking_cost_center(row):
    """Return the normalized persisted cost center used for NIS grouping."""
    vendor_data = _json_dict((row or {}).get("vendor_data_json"))
    return _normalized_key(vendor_data.get("cost_center"))


def _booking_vendor_code(row):
    """Return the normalized persisted vendor code (from DSM-NIS-INPUT) used
    for NIS booking-unit grouping.

    Grouping is driven by VENDOR because the NIS portal will not allow the
    same vendor + same invoice number to be booked twice. All SPV rows that
    resolve to the same vendor MUST therefore be booked in ONE NIS booking,
    each SPV contributing its own expense item (its own cost center / plant /
    amount) via the add-expense-item flow.
    """
    vendor_data = _json_dict((row or {}).get("vendor_data_json"))
    return _normalized_key(vendor_data.get("vendor_code"))


def _build_booking_units(rows):
    """Group pending booking rows into units to be booked together.

    Rows that resolve to the same non-empty VENDOR within the SAME email
    become one unit, regardless of SPV name or cost center. They are booked as
    a single NIS booking (same vendor + same invoice number cannot be booked
    twice), with one expense item per SPV row carrying that row's own cost
    center, plant and positive amount.

    Rows without a resolved vendor code stay separate so missing vendor data
    can never cause unrelated rows to merge. Order is preserved (oldest email,
    lowest row_index first).
    """
    units = []
    index = {}  # (run_id, vendor_code) -> unit list
    for row in rows:
        run_id = str(row.get("run_id") or row.get("message_id") or "")
        vendor_code = _booking_vendor_code(row)
        if vendor_code:
            key = (run_id, vendor_code)
            unit = index.get(key)
            if unit is None:
                unit = []
                index[key] = unit
                units.append(unit)
            unit.append(row)
        else:
            units.append([row])
    return units


def _unit_will_book(unit):
    """True when any row in this unit will actually drive an NIS booking."""
    return any(_row_will_book(r) for r in unit)


def _is_flat_vendor_data(value):
    if not isinstance(value, dict):
        return False
    vendor_keys = {"vendor_code", "company_code", "cost_center", "plant", "bank_key", "gl_account", "purpose"}
    return any(k in value for k in vendor_keys)


def _vendor_map_from_json(raw):
    data = _json_dict(raw)
    if not data:
        return {}
    if _is_flat_vendor_data(data):
        return {"": data}
    out = {}
    for spv_name, vendor_data in data.items():
        if isinstance(vendor_data, dict):
            out[str(spv_name)] = vendor_data
    return out


def _vendor_from_map(vendor_map, spv_name):
    wanted = _normalized_key(spv_name)
    for key, value in (vendor_map or {}).items():
        if key == "" or _normalized_key(key) == wanted:
            return value
    return {}


def _failure_context_from_run(run_id):
    run = db.get_run_by_id(run_id) or {}
    metadata = _json_dict(run.get("metadata_json"))
    spv_rows = _json_list(run.get("spv_dsm_json"))
    spv_name = str(run.get("spv_name") or "").strip()
    if not spv_name:
        for row in spv_rows:
            spv_name = str((row or {}).get("spv_name") or "").strip()
            if spv_name:
                break
    return (
        spv_name or "unknown",
        str(metadata.get("from_date") or "").strip(),
        str(metadata.get("to_date") or "").strip(),
    )


def _is_run_ready_for_nis_completion_email(run_id):
    run = db.get_run_by_id(run_id)
    if not run or db.is_nis_completion_email_sent(run_id):
        return False
    positive = _json_dict(run.get("positive_fields_json"))
    if not positive:
        return False
    nis_step = db.get_step(run_id, db.STEP_NIS_BOOKING)
    if not nis_step or nis_step.get("status") != db.STATUS_DONE:
        return False
    return db.are_all_sap_steps_complete(run_id)


def _send_nis_completion_email(run_id):
    if not _is_run_ready_for_nis_completion_email(run_id):
        return False

    recipients = db.get_nis_completion_recipients()
    if not recipients:
        print("  [EMAIL] NIS completion email not sent: no recipients in nis_completion_recipients.emails")
        return False

    run = db.get_run_by_id(run_id) or {}
    metadata = _json_dict(run.get("metadata_json"))
    spv_rows = _json_list(run.get("spv_dsm_json"))
    spv_name = str(run.get("spv_name") or "").strip()
    if not spv_name:
        for row in spv_rows:
            spv_name = str((row or {}).get("spv_name") or "").strip()
            if spv_name:
                break

    from_date = str(metadata.get("from_date") or "").strip()
    to_date = str(metadata.get("to_date") or "").strip()
    week_no = str(metadata.get("week_no") or "").strip()
    nis_number = str(run.get("nis_checklist_id") or "").strip()
    sap_number = str(run.get("checklist_value") or "").strip()
    sap_number_note = (
        sap_number
        if sap_number
        else "Not captured. SAP Save was submitted, but the checklist number could not be read from SAP."
    )
    booking_folder = str(run.get("booking_folder_path") or "").strip()

    subject_parts = ["DSM-NIS Automation Completed"]
    if spv_name:
        subject_parts.append(spv_name)
    if from_date or to_date:
        subject_parts.append(f"{from_date} to {to_date}".strip())
    subject = " | ".join(subject_parts)

    def esc(value):
        return html.escape(str(value or ""))

    body = f"""
<h3>DSM-NIS automation completed</h3>
<p>The NIS booking and SAP automation steps are complete.</p>
<table border="1" cellspacing="0" cellpadding="6">
  <tr><td><b>Run ID</b></td><td>{esc(run_id)}</td></tr>
  <tr><td><b>SPV</b></td><td>{esc(spv_name)}</td></tr>
  <tr><td><b>Week</b></td><td>{esc(week_no)}</td></tr>
  <tr><td><b>Period</b></td><td>{esc(from_date)} to {esc(to_date)}</td></tr>
  <tr><td><b>NIS Booking Number</b></td><td>{esc(nis_number or "n/a")}</td></tr>
  <tr><td><b>SAP Checklist Number</b></td><td>{esc(sap_number_note)}</td></tr>
  <tr><td><b>Booking Folder</b></td><td>{esc(booking_folder)}</td></tr>
</table>
"""

    try:
        import msal
        app = msal.ConfidentialClientApplication(
            config.GRAPH_CLIENT_ID,
            authority=config.GRAPH_AUTHORITY,
            client_credential=config.GRAPH_CLIENT_SECRET,
        )
        result = app.acquire_token_for_client(scopes=config.GRAPH_SCOPE)
        token = result.get("access_token")
        if not token:
            print("  [EMAIL] NIS completion email not sent: could not acquire Graph token")
            return False

        payload = {
            "message": {
                "subject": subject,
                "body": {"contentType": "HTML", "content": body},
                "toRecipients": [{"emailAddress": {"address": addr}} for addr in recipients],
            },
            "saveToSentItems": "true",
        }
        resp = requests.post(
            f"https://graph.microsoft.com/v1.0/users/{config.USER_EMAIL}/sendMail",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json=payload,
            timeout=60,
        )
        if resp.status_code in (200, 202):
            db.mark_nis_completion_email_sent(run_id)
            print(f"  [EMAIL] NIS completion notification sent to {', '.join(recipients)}")
            return True
        print(f"  [EMAIL] NIS completion email failed: HTTP {resp.status_code} {resp.text[:300]}")
    except Exception as exc:
        print(f"  [EMAIL] Could not send NIS completion notification: {exc}")
    return False


def send_ready_nis_completion_emails():
    sent = 0
    for run in db.get_runs_pending_nis_completion_email():
        try:
            if _send_nis_completion_email(run.get("id") or ""):
                sent += 1
        except Exception as exc:
            print(f"  [EMAIL] Completion notification check failed for run_id={run.get('id')}: {exc}")
    return sent


# ── Stage Notification Emails (HOD / FC / Completion) ────────────────────────
def _send_stage_email(row, stage: str, extra_info: str = "") -> bool:
    """Send a stage notification email for a booking row.

    stage: 'hod_approved', 'fc_approved', 'all_complete'
    """
    row_id = row.get("row_id")
    spv_name = str(row.get("spv_name") or "").strip()
    metadata = _json_dict(row.get("metadata_json"))
    from_date = str(metadata.get("from_date") or "").strip()
    to_date = str(metadata.get("to_date") or "").strip()
    nis_number = str(row.get("nis_checklist_id") or "").strip()
    booking_folder = str(row.get("booking_folder_path") or "").strip()

    stage_labels = {
        "hod_approved": "HOD Approval Received",
        "fc_approved": "FC Approval Received & PDF Merged",
        "all_complete": "All Stages Complete",
    }
    stage_label = stage_labels.get(stage, stage)

    subject = f"DSM-NIS [{stage_label}] | {spv_name} | {from_date} to {to_date}"

    def esc(value):
        return html.escape(str(value or ""))

    body = f"""
<h3>DSM-NIS Stage Update: {esc(stage_label)}</h3>
<table border="1" cellspacing="0" cellpadding="6">
  <tr><td><b>SPV</b></td><td>{esc(spv_name)}</td></tr>
  <tr><td><b>Period</b></td><td>{esc(from_date)} to {esc(to_date)}</td></tr>
  <tr><td><b>NIS Number</b></td><td>{esc(nis_number)}</td></tr>
  <tr><td><b>Stage</b></td><td>{esc(stage_label)}</td></tr>
  <tr><td><b>Booking Folder</b></td><td>{esc(booking_folder)}</td></tr>
</table>
{f'<p>{esc(extra_info)}</p>' if extra_info else ''}
"""
    recipients = db.get_nis_completion_recipients()
    if not recipients:
        print(f"  [EMAIL] Stage email not sent: no recipients")
        return False
    return send_html_email(subject, body, recipients)


def send_hod_stage_emails():
    """Send HOD stage notification for rows that just got HOD approval."""
    sent = 0
    for row in db.get_rows_awaiting_fc():
        row_id = row.get("row_id")
        if db.is_hod_email_sent(row_id):
            continue
        if _send_stage_email(row, "hod_approved"):
            db.mark_hod_email_sent(row_id)
            sent += 1
    return sent


def send_fc_stage_emails():
    """Send FC stage notification for rows that have FC approval + PDF merge done."""
    sent = 0
    for row in db.get_rows_awaiting_completion_email():
        row_id = row.get("row_id")
        if db.is_fc_email_sent(row_id):
            continue
        merged_name = str(row.get("merged_pdf_name") or "").strip()
        extra = f"Merged PDF: {merged_name}" if merged_name else ""
        if _send_stage_email(row, "fc_approved", extra):
            db.mark_fc_email_sent(row_id)
            sent += 1
    return sent


def send_final_completion_emails():
    """Send the FINAL completion email only after ALL stages are done.

    Stages: NIS done → SAP done → HOD approved → FC approved (NIS Checklist
    Print email received with PDF) → PDF merged.

    The completion email includes the merged PDF as an attachment. The merged
    PDF contains: email.pdf + invoice attachment + NIS checklist PDF + SAP
    checklist PDF(s).

    Only rows that have passed ALL stages and haven't had their completion
    email sent yet will trigger the email.
    """
    sent = 0
    for row in db.get_rows_awaiting_completion_email():
        row_id = row.get("row_id")
        if db.is_row_completion_email_sent(row_id):
            continue
        if not db.is_row_all_stages_done(row_id):
            continue

        # Download the merged PDF from SharePoint to attach to the email
        merged_pdf_path = None
        booking_folder = str(row.get("booking_folder_path") or "").strip()
        merged_name = str(row.get("merged_pdf_name") or "").strip()
        if booking_folder and merged_name:
            try:
                tmp_dir = _row_tmp_dir(f"completion_{row_id}")
                local_path = os.path.join(tmp_dir, merged_name)
                dl = sp.download_file_from_sp(f"{booking_folder}/{merged_name}", local_path)
                if dl and os.path.isfile(dl):
                    merged_pdf_path = dl
            except Exception as dl_exc:
                print(f"  [EMAIL] Could not download merged PDF for row_id={row_id}: {dl_exc}")

        extra = f"Merged PDF: {merged_name}" if merged_name else ""
        if _send_completion_email_with_attachment(row, merged_pdf_path, extra):
            db.mark_row_completion_email_sent(row_id)
            sent += 1
            print(f"  [EMAIL] ✅ Final completion email sent for row_id={row_id}")

        # Cleanup temp
        if merged_pdf_path:
            try:
                _cleanup_tmp(os.path.dirname(merged_pdf_path))
            except Exception:
                pass
    return sent


def _send_completion_email_with_attachment(row, merged_pdf_path=None, extra_info=""):
    """Send the final completion email with the merged PDF attached."""
    row_id = row.get("row_id")
    spv_name = str(row.get("spv_name") or "").strip()
    metadata = _json_dict(row.get("metadata_json"))
    from_date = str(metadata.get("from_date") or "").strip()
    to_date = str(metadata.get("to_date") or "").strip()
    nis_number = str(row.get("nis_checklist_id") or "").strip()
    sap_number = str(row.get("checklist_value") or "").strip()
    booking_folder = str(row.get("booking_folder_path") or "").strip()
    merged_name = str(row.get("merged_pdf_name") or "").strip()

    subject = f"DSM-NIS [All Stages Complete] | {spv_name} | {from_date} to {to_date}"

    def esc(value):
        return html.escape(str(value or ""))

    body = f"""
<h3>DSM-NIS Automation Complete</h3>
<p>All workflow stages are done. The merged PDF is attached and available in the booking folder.</p>
<table border="1" cellspacing="0" cellpadding="6">
  <tr><td><b>SPV</b></td><td>{esc(spv_name)}</td></tr>
  <tr><td><b>Period</b></td><td>{esc(from_date)} to {esc(to_date)}</td></tr>
  <tr><td><b>NIS Number</b></td><td>{esc(nis_number)}</td></tr>
  <tr><td><b>SAP Checklist</b></td><td>{esc(sap_number or 'N/A (no negatives)')}</td></tr>
  <tr><td><b>Merged PDF</b></td><td>{esc(merged_name)}</td></tr>
  <tr><td><b>Booking Folder</b></td><td>{esc(booking_folder)}</td></tr>
</table>
{f'<p>{esc(extra_info)}</p>' if extra_info else ''}
"""
    recipients = db.get_nis_completion_recipients()
    if not recipients:
        print(f"  [EMAIL] Final completion email not sent: no recipients")
        return False

    pdf_attachments = [merged_pdf_path] if merged_pdf_path and os.path.isfile(merged_pdf_path) else None
    return send_html_email(subject, body, recipients, pdf_attachments=pdf_attachments)


def _auto_advance_completed_rows():
    """Auto-approve HOD for rows where NIS number is captured and row is done.

    The actual workflow stages are:
      1. NIS done → NIS number captured
      2. HOD approval → incoming email "NIS Approval - <nis_number>"
      3. FC approval / NIS Checklist PDF → incoming email "NIS Checklist Print - <nis_number>"
         with the NIS checklist PDF attached (downloaded + uploaded to booking folder)
      4. SAP done → SAP checklist number + PDF captured
      5. PDF merge → only when BOTH NIS checklist PDF and SAP checklist PDF
         (if applicable) are in the booking folder
      6. Final completion email → sent once all stages complete

    This function auto-approves HOD once NIS+SAP are both captured (row is done),
    so the system is ready to process the incoming "NIS Checklist Print" email
    (which serves as FC approval). SAP runs independently.
    """
    rows = db.get_rows_pending_auto_advance()

    for row in rows:
        row_id = row.get("row_id")
        nis_number = str(row.get("nis_checklist_id") or "").strip()
        sap_number = str(row.get("checklist_value") or "").strip()
        negative = _json_dict(row.get("negative_fields_json"))

        # NIS number MUST be captured for the row to advance
        if not nis_number:
            continue

        # SAP number must be captured IF the row has negative fields
        has_negatives = bool(negative and any(
            safe_float(v) > 0 for v in negative.values()
        ))
        if has_negatives and not sap_number:
            continue

        spv_name = str(row.get("spv_name") or "").strip()

        # Auto-approve HOD if not yet done (NIS + SAP both captured means
        # the row is fully done — HOD can proceed automatically)
        if row.get("hod_approval_status") != db.STATUS_DONE:
            db.mark_hod_approved(row_id, "auto_approved_on_completion")
            print(f"  [AUTO-ADVANCE] HOD auto-approved for row_id={row_id} SPV={spv_name} NIS={nis_number}")

        # FC approval comes from the "NIS Checklist Print" email — do NOT auto-approve.
        # The _handle_checklist_email / _handle_fc_approval_email functions handle it
        # when the email arrives with the NIS checklist PDF attached.


def _is_row_ready_for_merge(row) -> bool:
    """True when a row has ALL required checklist PDFs for the final merge.

    Requirements:
      - NIS checklist PDF uploaded (from "NIS Checklist Print" email)
      - SAP checklist PDF uploaded (from SAP automation) — only if row has negatives
      - FC approval done (the "NIS Checklist Print" email IS the FC approval)
    """
    if not db.is_fc_approved(row.get("row_id")):
        return False
    if not _is_truthy(row.get("fc_checklist_pdf_uploaded")):
        return False

    negative = _json_dict(row.get("negative_fields_json"))
    has_negatives = bool(negative and any(
        safe_float(v) > 0 for v in negative.values()
    ))
    if has_negatives:
        if not _is_truthy(row.get("sap_checklist_pdf_uploaded")):
            return False

    return True


def process_approval_emails_and_stages():
    """Called in the main loop — processes approval stages + PDF merge.

    The workflow order is:
      1. Auto-advance HOD for rows with NIS+SAP both captured
      2. Send HOD stage notification emails
      3. Send FC stage notification emails
      4. NIS+SAP completion emails (run-level notification)
      5. PDF merge — only when ALL PDFs are in the booking folder:
         email.pdf + invoice PDF + NIS checklist PDF + SAP checklist PDF(s)
      6. Final completion email — sent once PDF merge is done

    NOTE: FC approval + NIS checklist PDF download happen via email detection
    (_handle_checklist_email / _handle_fc_approval_email) — not here.
    """
    # Auto-advance HOD for completed rows
    _auto_advance_completed_rows()

    # NIS+SAP completion emails (triggered after SAP finishes remotely)
    send_ready_nis_completion_emails()

    # HOD stage notification emails
    send_hod_stage_emails()

    # FC stage notification emails
    send_fc_stage_emails()

    # Attempt PDF merge for rows that have ALL required checklist PDFs
    for row in db.get_rows_awaiting_pdf_merge():
        if _is_row_ready_for_merge(row):
            _try_pdf_merge_for_row(row)

    # Final completion emails — only after PDF merge is done
    send_final_completion_emails()


# ── Booking folder ───────────────────────────────────────────────────────────
def _prepare_html_for_pdf(body_html: str) -> str:
    """Wrap email HTML in a full document with explicit styles to ensure
    visibility when rendered to PDF. Outlook emails often omit background/color
    declarations, causing headless renderers to produce blank-looking PDFs."""
    # If the HTML already has a <html> root, inject our overrides into it;
    # otherwise wrap entirely.
    import re as _re
    _OVERRIDE_CSS = """
        body {
            background: #ffffff !important;
            color: #000000 !important;
            font-family: Arial, "Times New Roman", sans-serif;
            font-size: 11pt;
            padding: 20px;
            margin: 0;
        }
        table {
            border-collapse: collapse;
            width: auto;
            margin: 8px 0;
        }
        td, th {
            border: 1px solid #333;
            padding: 6px 10px;
            color: #000 !important;
            font-size: 10pt;
        }
        th {
            background: #f0f0f0 !important;
            font-weight: bold;
        }
        p { margin: 4px 0; color: #000 !important; }
        span { color: inherit !important; }
        /* Override any display:none on style tags (Outlook quirk) */
        style { display: none !important; }
        /* Ensure images don't break layout */
        img { max-width: 100%; height: auto; }
    """

    # Strip conditional comments (<!--[if ...]> ... <![endif]-->)
    clean_html = _re.sub(r'<!--\[if.*?\]>.*?<!\[endif\]-->', '', body_html, flags=_re.DOTALL | _re.IGNORECASE)
    # Remove mso-* style properties that can cause rendering issues
    clean_html = _re.sub(r'mso-[^;"\']+[:][^;"\']+[;]?', '', clean_html)

    if _re.search(r'<html[\s>]', clean_html, _re.IGNORECASE):
        # Inject override CSS into existing <head> or after <html>
        style_tag = f'<style type="text/css">{_OVERRIDE_CSS}</style>'
        if _re.search(r'<head[\s>]', clean_html, _re.IGNORECASE):
            clean_html = _re.sub(
                r'(<head[^>]*>)',
                rf'\1{style_tag}',
                clean_html, count=1, flags=_re.IGNORECASE
            )
        else:
            clean_html = _re.sub(
                r'(<html[^>]*>)',
                rf'\1<head>{style_tag}</head>',
                clean_html, count=1, flags=_re.IGNORECASE
            )
    else:
        # Wrap in a full HTML document
        clean_html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style type="text/css">{_OVERRIDE_CSS}</style>
</head>
<body>
{clean_html}
</body>
</html>"""
    return clean_html


def _convert_html_to_pdf(body_html: str, output_path: str) -> bool:
    """Convert HTML content to PDF. Tries multiple methods. Returns True on success."""
    prepared_html = _prepare_html_for_pdf(body_html)

    # Method 1: Playwright headless Chromium (best quality)
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            page = browser.new_page()
            page.set_content(prepared_html, wait_until="networkidle")
            page.pdf(path=output_path, format="A4", print_background=True,
                     margin={"top": "15mm", "bottom": "15mm", "left": "10mm", "right": "10mm"})
            browser.close()
        if os.path.isfile(output_path) and os.path.getsize(output_path) > 0:
            return True
    except Exception as exc:
        print(f"  [PDF-CONVERT] Playwright method failed: {exc}")

    # Method 2: pdfkit / wkhtmltopdf (if installed)
    try:
        import pdfkit
        pdfkit.from_string(prepared_html, output_path, options={
            "quiet": "",
            "margin-top": "15mm",
            "margin-bottom": "15mm",
            "margin-left": "10mm",
            "margin-right": "10mm",
            "encoding": "UTF-8",
            "no-stop-slow-scripts": "",
        })
        if os.path.isfile(output_path) and os.path.getsize(output_path) > 0:
            return True
    except Exception as exc:
        print(f"  [PDF-CONVERT] pdfkit method failed: {exc}")

    # Method 3: weasyprint (if installed)
    try:
        from weasyprint import HTML
        HTML(string=prepared_html).write_pdf(output_path)
        if os.path.isfile(output_path) and os.path.getsize(output_path) > 0:
            return True
    except Exception as exc:
        print(f"  [PDF-CONVERT] weasyprint method failed: {exc}")

    # Method 4: reportlab with HTML table structure preserved
    try:
        _convert_html_to_pdf_reportlab(body_html, output_path)
        if os.path.isfile(output_path) and os.path.getsize(output_path) > 0:
            return True
    except Exception as exc:
        print(f"  [PDF-CONVERT] reportlab method failed: {exc}")

    return False


def _convert_html_to_pdf_reportlab(body_html: str, output_path: str):
    """Fallback PDF conversion using reportlab that preserves table structure.
    Parses the HTML tables and renders them as proper PDF tables with borders,
    headers, and cell data — instead of just dumping stripped plain text."""
    import re as _re
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    )
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from html import unescape

    styles = getSampleStyleSheet()
    # Custom styles for email content
    body_style = ParagraphStyle(
        'EmailBody', parent=styles['Normal'],
        fontSize=9, leading=12, textColor=colors.black
    )
    header_style = ParagraphStyle(
        'TableHeader', parent=styles['Normal'],
        fontSize=8, leading=10, textColor=colors.black, alignment=1  # center
    )
    cell_style = ParagraphStyle(
        'TableCell', parent=styles['Normal'],
        fontSize=8, leading=10, textColor=colors.black, alignment=1  # center
    )

    story = []

    # ── Extract non-table text (paragraphs before/between/after tables) ──
    def _strip_tags(html_str):
        text = _re.sub(r'<br\s*/?\s*>', '\n', html_str, flags=_re.I)
        text = unescape(_re.sub(r'<[^>]+>', '', text))
        return _re.sub(r'[ \t]+', ' ', text).strip()

    # Split content into table and non-table sections
    parts = _re.split(r'(<table[\s\S]*?</table>)', body_html, flags=_re.I)

    for part in parts:
        part_stripped = part.strip()
        if not part_stripped:
            continue

        if _re.match(r'<table', part_stripped, _re.I):
            # ── Parse the HTML table into rows/cols ──
            rows_html = _re.findall(r'<tr[^>]*>([\s\S]*?)</tr>', part_stripped, _re.I)
            if not rows_html:
                continue

            table_data = []
            for row_html in rows_html:
                cells = _re.findall(r'<t[dh][^>]*>([\s\S]*?)</t[dh]>', row_html, _re.I)
                row_cells = []
                for cell_html in cells:
                    cell_text = _strip_tags(cell_html)
                    # Truncate overly long cell values
                    if len(cell_text) > 80:
                        cell_text = cell_text[:77] + "..."
                    row_cells.append(Paragraph(cell_text, cell_style))
                if row_cells:
                    table_data.append(row_cells)

            if not table_data:
                continue

            # Normalize column count (pad short rows)
            max_cols = max(len(r) for r in table_data)
            for row in table_data:
                while len(row) < max_cols:
                    row.append(Paragraph("", cell_style))

            # Build the reportlab Table
            # Use landscape-friendly column widths
            page_width = A4[0] - 20 * mm  # 20mm margins
            col_width = page_width / max_cols if max_cols > 0 else 60 * mm
            col_widths = [col_width] * max_cols

            tbl = Table(table_data, colWidths=col_widths, repeatRows=1)
            tbl_style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.9, 0.9, 0.9)),
                ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                # Bold first row (header)
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ])
            tbl.setStyle(tbl_style)
            story.append(tbl)
            story.append(Spacer(1, 8))
        else:
            # Non-table content — render as paragraphs
            text = _strip_tags(part_stripped)
            if text:
                for line in text.split('\n'):
                    line = line.strip()
                    if line:
                        # Escape special reportlab characters
                        line = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                        story.append(Paragraph(line, body_style))
                        story.append(Spacer(1, 4))

    if not story:
        # Absolute fallback: just dump whatever text is there
        plain = _strip_tags(body_html)
        if plain:
            for line in plain.split('\n')[:200]:
                line = line.strip()
                if line:
                    line = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                    story.append(Paragraph(line, body_style))
                    story.append(Spacer(1, 3))

    if story:
        doc = SimpleDocTemplate(
            output_path, pagesize=A4,
            leftMargin=10 * mm, rightMargin=10 * mm,
            topMargin=15 * mm, bottomMargin=15 * mm
        )
        doc.build(story)


def _save_email_to_booking_folder(metadata_row, spv_dsm_rows, body_html, attachment_dir, subject=""):
    from_date = str(metadata_row.get("from_date","")).strip()
    to_date   = str(metadata_row.get("to_date","")).strip()

    # ── Collect ALL unique SPV folder labels ─────────────────────────────────
    # Each SPV gets its OWN folder (using its full name). Documents (email.pdf,
    # invoice attachments) are duplicated to every folder so each SPV has a
    # complete independent set of files — even when multiple SPVs share the
    # same cost center and are booked together in one NIS booking.
    unique_labels = []
    seen = set()
    for row in (spv_dsm_rows or []):
        name = str(row.get("spv_name", "")).strip()
        if not name:
            continue
        label = _folder_spv_name(name)
        if label not in seen:
            seen.add(label)
            unique_labels.append(label)

    # Fallback: if no SPV names found, use a default
    if not unique_labels:
        unique_labels = ["unknown_spv"]

    # ── Create a SharePoint folder for EACH unique SPV label ─────────────────
    folder_map = {}  # {label: sp_folder_path}
    for label in unique_labels:
        sp_folder = sp.build_booking_folder_path(label, from_date, to_date)
        sp.ensure_sp_folder(sp_folder)
        folder_map[label] = sp_folder

    # ── Move attachment PDFs from email staging dir to local temp ─────────────
    first_label = unique_labels[0]
    local_tmp = _spv_tmp_dir(first_label)
    if attachment_dir and os.path.isdir(attachment_dir):
        for fname in os.listdir(attachment_dir):
            if fname.lower().endswith(".pdf"):
                try:
                    shutil.move(os.path.join(attachment_dir, fname), os.path.join(local_tmp, fname))
                except Exception:
                    pass
        _cleanup_tmp(attachment_dir)

    # ── Convert email body to PDF ONCE ───────────────────────────────────────
    email_pdf = os.path.join(local_tmp, "email.pdf")
    pdf_created = _convert_html_to_pdf(body_html, email_pdf)
    if pdf_created:
        # Upload email.pdf to ALL SPV folders
        for label, sp_folder in folder_map.items():
            sp.upload_file_to_sp(email_pdf, sp_folder, "email.pdf")
        print(f"  [BOOKING] email.pdf created and uploaded to {len(folder_map)} folder(s)")
    else:
        print(f"  [BOOKING] ⚠ Could not convert email to PDF — all methods failed")

    # ── Upload invoice/attachment PDFs to ALL SPV folders ────────────────────
    pdf_files = [f for f in os.listdir(local_tmp) if f.lower().endswith(".pdf") and f != "email.pdf"]
    for fname in pdf_files:
        local_path = os.path.join(local_tmp, fname)
        for label, sp_folder in folder_map.items():
            sp.upload_file_to_sp(local_path, sp_folder, fname)

    _cleanup_tmp(local_tmp)
    print(f"  [BOOKING] SharePoint folders created: {list(folder_map.values())}")
    return folder_map


# ── Pipeline steps ───────────────────────────────────────────────────────────
def _step_extraction(run_id, body_html, parsed_tables=None):
    from classified_html_table_parser import parse_html_tables, print_raw_tables
    if db.is_step_done(run_id, db.STEP_EMAIL_EXTRACTION):
        run = db.get_run_by_id(run_id)
        return json.loads(run["metadata_json"] or "{}"), _valid_spv_rows(json.loads(run["spv_dsm_json"] or "[]"))
    db.step_start(run_id, db.STEP_EMAIL_EXTRACTION)
    try:
        print_raw_tables(body_html)
        tables = parsed_tables or parse_html_tables(body_html)
        spv_dsm_rows = _valid_spv_rows(tables.get("spv_dsm_table", []))
        metadata_row = (tables.get("metadata_table", []) or [{}])[0]
        if not spv_dsm_rows: raise ValueError("spv_dsm_table is empty")
        if not metadata_row: raise ValueError("metadata_table is empty")
        db.save_extracted_data(run_id, metadata_row, spv_dsm_rows)
        db.step_done(run_id, db.STEP_EMAIL_EXTRACTION, detail=f"{len(spv_dsm_rows)} SPV row(s)")
        return metadata_row, spv_dsm_rows
    except Exception as exc:
        db.step_failed(run_id, db.STEP_EMAIL_EXTRACTION, str(exc)); raise

def _step_classification(run_id, spv_row):
    step = db.STEP_EMAIL_CLASSIFICATION
    if db.is_step_done(run_id, step):
        run = db.get_run_by_id(run_id)
        return _classify_amounts(spv_row)
    db.step_start(run_id, step)
    try:
        positive, negative = _classify_amounts(spv_row, emit=True)
        db.save_classification(run_id, positive, negative)
        if positive: db.ensure_dynamic_step(run_id, db.STEP_NIS_BOOKING)
        for fk in negative: db.ensure_dynamic_step(run_id, db.step_sap(fk))
        db.step_done(run_id, step, detail=f"positive={list(positive.keys())} negative={list(negative.keys())}")
        return positive, negative
    except Exception as exc:
        db.step_failed(run_id, step, str(exc)); raise

def _step_classification_all(run_id, spv_rows):
    step = db.STEP_EMAIL_CLASSIFICATION
    classifications = []
    if db.is_step_done(run_id, step):
        for spv_row in spv_rows:
            positive, negative = _classify_amounts(spv_row)
            classifications.append((spv_row, positive, negative))
        return classifications

    db.step_start(run_id, step)
    try:
        all_positive, all_negative = [], []
        for row_index, spv_row in enumerate(spv_rows, start=1):
            spv_name = str(spv_row.get("spv_name") or "").strip()
            print(f"\n  ── Classifying row {row_index}: {spv_name} ──")
            positive, negative = _classify_amounts(spv_row, emit=True)
            classifications.append((spv_row, positive, negative))
            all_positive.append(positive)
            all_negative.append(negative)
            if positive:
                db.ensure_dynamic_step(run_id, _row_step_name(db.STEP_NIS_BOOKING, row_index, spv_name))
            # NOTE: Do NOT create run-level SAP steps here (db.step_sap(fk)).
            # Per-row SAP steps will be created in _process_booking_row with the
            # proper _row_N_<spv> suffix. Creating run-level steps here would
            # cause duplicates — the same negative amount processed twice.

        positive_total = _sum_amount_maps(all_positive)
        negative_total = _sum_amount_maps(all_negative)
        db.save_classification(run_id, positive_total, negative_total)
        if positive_total:
            db.ensure_dynamic_step(run_id, db.STEP_NIS_BOOKING)
        db.step_done(
            run_id,
            step,
            detail=f"rows={len(classifications)} positive={list(positive_total.keys())} negative={list(negative_total.keys())}",
        )
        return classifications
    except Exception as exc:
        db.step_failed(run_id, step, str(exc)); raise

def _step_validation(run_id, metadata_row, spv_row, positive, negative, step_name=None):
    step = step_name or db.STEP_EMAIL_VALIDATION
    if db.is_step_done(run_id, step): return
    db.step_start(run_id, step)
    errors = []
    for f in ("week_no","from_date","to_date","due_date"):
        if not str(metadata_row.get(f,"")).strip(): errors.append(f"metadata.{f} missing")
    if not str(spv_row.get("spv_name","")).strip(): errors.append("spv_name missing")
    if not positive and not negative: errors.append("All amounts zero")
    if errors:
        msg = "; ".join(errors); db.step_failed(run_id, step, msg); raise ValueError(msg)
    db.step_done(run_id, step, detail=f"week={metadata_row.get('week_no')} spv={spv_row.get('spv_name')}")

def _step_vendor_master(run_id, spv_name):
    step = db.STEP_READ_VENDOR
    if db.is_step_done(run_id, step):
        run = db.get_run_by_id(run_id)
        vendor_map = _vendor_map_from_json(run.get("vendor_data_json"))
        return _vendor_from_map(vendor_map, spv_name)
    return _load_vendor_for_spv(run_id, spv_name, step)

def _load_vendor_for_spv(run_id, spv_name, step_name):
    run = db.get_run_by_id(run_id) or {}
    vendor_map = _vendor_map_from_json(run.get("vendor_data_json"))
    cached = _vendor_from_map(vendor_map, spv_name)
    if cached:
        return cached

    db.step_start(run_id, step_name, detail=f"sheet={spv_name}")
    try:
        raw = read_vendor_data(spv_name)
        vendor_data = normalize_vendor_data(raw)
        with _VENDOR_MAP_LOCK:
            # Re-read inside the lock to avoid overwriting another worker's write
            run = db.get_run_by_id(run_id) or {}
            vendor_map = _vendor_map_from_json(run.get("vendor_data_json"))
            if "" in vendor_map:
                vendor_map.pop("", None)
            vendor_map[spv_name] = vendor_data
            db.save_vendor_data(run_id, vendor_map)
        db.step_done(run_id, step_name, detail=f"{spv_name}={json.dumps(vendor_data)}")
        return vendor_data
    except Exception as exc:
        # Exact match failed — try closest/fuzzy match as fallback
        print(f"  [VENDOR] Exact match failed for '{spv_name}': {exc}")
        print(f"  [VENDOR] Trying fuzzy/closest name match...")
        try:
            from read_vendor_master_data import find_closest_spv
            closest_row = find_closest_spv(spv_name)
            if closest_row:
                matched_name = str(closest_row.get("SPV Name") or "").strip()
                print(f"  [VENDOR] Closest match: '{matched_name}' for '{spv_name}'")
                raw_fallback = {
                    "Vendor":      str(closest_row.get("Vendor") or "").strip(),
                    "Company":     str(closest_row.get("Company") or "").strip(),
                    "Cost Center": str(closest_row.get("Cost Center") or "").strip(),
                    "Plant":       str(closest_row.get("Plant") or "").strip(),
                    "Bank Key":    str(closest_row.get("Bank Key") or "").strip(),
                    "GL":          str(closest_row.get("GL") or "").strip(),
                }
                vendor_data = normalize_vendor_data(raw_fallback)
                if vendor_data.get("vendor_code"):
                    with _VENDOR_MAP_LOCK:
                        run = db.get_run_by_id(run_id) or {}
                        vendor_map = _vendor_map_from_json(run.get("vendor_data_json"))
                        if "" in vendor_map:
                            vendor_map.pop("", None)
                        vendor_map[spv_name] = vendor_data
                        db.save_vendor_data(run_id, vendor_map)
                    db.step_done(run_id, step_name,
                                 detail=f"{spv_name} (fuzzy→{matched_name})={json.dumps(vendor_data)}")
                    return vendor_data
                else:
                    print(f"  [VENDOR] Closest match '{matched_name}' has no vendor code — failing.")
            else:
                print(f"  [VENDOR] No fuzzy match found for '{spv_name}'.")
        except Exception as fuzzy_exc:
            print(f"  [VENDOR] Fuzzy match also failed: {fuzzy_exc}")
        db.step_failed(run_id, step_name, str(exc))
        raise

def _step_vendor_master_for_row(run_id, spv_name, row_index):
    step = _row_step_name(db.STEP_READ_VENDOR, row_index, spv_name)
    if db.is_step_done(run_id, step):
        run = db.get_run_by_id(run_id) or {}
        vendor_map = _vendor_map_from_json(run.get("vendor_data_json"))
        cached = _vendor_from_map(vendor_map, spv_name)
        if cached:
            return cached
    return _load_vendor_for_spv(run_id, spv_name, step)

def _step_nis_booking(run_id, positive, spv_row, negative, vendor_data, invoice_date, attachment_dir, metadata_row=None, booking_dir="", step_name=None, row_index=None, booking_row_id=None, work_dir=None, login_barrier=None, expense_items=None, extra_booking_row_ids=None):
    step = step_name or db.STEP_NIS_BOOKING
    extra_booking_row_ids = extra_booking_row_ids or []
    label = ", ".join(lbl for k, lbl in SPV_AMOUNT_FIELDS if k in positive)
    if db.is_step_done(run_id, step): return
    # ── Idempotency guard ────────────────────────────────────────────────────
    # If this booking row has already produced an NIS number, the booking was
    # done in a previous attempt/run. NEVER book it again (that would create a
    # duplicate NIS entry in the portal). Just mark the step done and return.
    if booking_row_id is not None:
        existing = db.get_booking_row(booking_row_id) or {}
        existing_nis = str(existing.get("nis_checklist_id") or "").strip()
        if existing_nis:
            print(f"  [SKIP] Row already booked (NIS={existing_nis}) — not booking again.")
            # Propagate the existing NIS number to every cost-center-grouped row
            # so their tracker rows get the same number.
            for _eid in extra_booking_row_ids:
                if not str((db.get_booking_row(_eid) or {}).get("nis_checklist_id") or "").strip():
                    db.save_booking_row_nis_checklist_id(_eid, existing_nis)
            if not db.is_step_done(run_id, step):
                db.step_done(run_id, step, detail=f"already_booked nis_number={existing_nis}")
            return
    retry_count = db.get_step_retry_count(run_id, step)
    if retry_count >= MAX_STEP_RETRIES:
        msg = f"NIS [{label}] exceeded {MAX_STEP_RETRIES} retries"
        db.step_failed(run_id, step, msg); raise RuntimeError(msg)
    # Find PDF
    pdf_path, pdf_files = "", []
    spv_name_tmp = str(spv_row.get("spv_name", "")) or str((metadata_row or {}).get("spv_name", "")) or f"run_{run_id}"
    # Use a caller-supplied per-row working directory when running in parallel
    # so concurrent bookings never read/write each other's PDFs. Falls back to
    # the per-SPV temp dir for the sequential path.
    spv_tmp = work_dir or _spv_tmp_dir(spv_name_tmp)
    pdf_files = _sorted_invoice_pdf_paths(Path(spv_tmp).glob("*.pdf"))
    if not pdf_files and attachment_dir and os.path.isdir(attachment_dir):
        pdf_files = _sorted_invoice_pdf_paths(Path(attachment_dir).glob("*.pdf"))
    if not pdf_files:
        # No invoice PDF found — check for email.pdf in the work_dir
        # (pre-downloaded as fallback by _process_booking_row)
        _email_pdf_in_work = os.path.join(spv_tmp, "email.pdf")
        if os.path.isfile(_email_pdf_in_work):
            pdf_files = [_email_pdf_in_work]
    if pdf_files: pdf_path = str(pdf_files[0])
    elif booking_dir:
        sp_items = sp.list_sp_folder(booking_dir)
        invoice_pdfs = sorted(
            [f["name"] for f in sp_items
             if not f["is_folder"] and _is_invoice_upload_pdf_name(f["name"])],
            key=str.lower,
        )
        if invoice_pdfs:
            dl = sp.download_file_from_sp(
                f"{booking_dir}/{invoice_pdfs[0]}",
                os.path.join(spv_tmp, invoice_pdfs[0]),
            )
            if dl: pdf_path = dl
    if not pdf_path:
        # ── Ultimate fallback: download email.pdf from booking folder ────────
        # Every email run has an email.pdf (email body converted to PDF) stored
        # in the booking folder. Use it as a last resort so NO row ever proceeds
        # without a PDF upload.
        if booking_dir:
            _email_pdf_fallback = os.path.join(spv_tmp, "email.pdf")
            if not os.path.isfile(_email_pdf_fallback):
                try:
                    _dl = sp.download_file_from_sp(
                        f"{booking_dir}/email.pdf", _email_pdf_fallback
                    )
                    if _dl and os.path.isfile(_dl):
                        pdf_path = _dl
                        print(f"  [PDF] Using email.pdf as final fallback for NIS upload")
                except Exception:
                    pass
            else:
                pdf_path = _email_pdf_fallback
        if not pdf_path:
            print("  [WARN] No PDF found (not even email.pdf). Proceeding without PDF upload.")

    # ── AWEK exception rule ──────────────────────────────────────────────────
    # For SPVs whose name matches any keyword in NIS_EMAIL_PDF_SPV_KEYWORDS
    # (e.g. AWEK06L, AWEK03L, AWEK05L): ignore the invoice PDF and upload the
    # email.pdf (email body converted to PDF) instead. The invoice number for
    # these SPVs is always auto-generated (handled inside book_nis).
    if is_email_pdf_spv(spv_name_tmp):
        _email_pdf_local = os.path.join(spv_tmp, "email.pdf")
        _email_pdf_found = False
        # 1) Already downloaded locally (e.g. from the work_dir copy)
        if os.path.isfile(_email_pdf_local):
            pdf_path = _email_pdf_local
            _email_pdf_found = True
        # 2) Download from the booking folder on SharePoint
        if not _email_pdf_found and booking_dir:
            _dl = sp.download_file_from_sp(
                f"{booking_dir}/email.pdf",
                _email_pdf_local,
            )
            if _dl and os.path.isfile(_dl):
                pdf_path = _dl
                _email_pdf_found = True
        if _email_pdf_found:
            print(f"  [AWEK-EXCEPTION] {spv_name_tmp}: using email.pdf for NIS upload "
                  f"(invoice PDF suppressed).")
        else:
            # Last resort: regenerate email.pdf from body_html stored in DB
            _run_data = db.get_run_by_id(run_id) or {}
            _body_html = str(_run_data.get("body_html") or "").strip()
            if _body_html:
                print(f"  [AWEK-EXCEPTION] {spv_name_tmp}: email.pdf not on SharePoint — regenerating from email body HTML")
                _regen_path = os.path.join(spv_tmp, "email.pdf")
                _regen_ok = _convert_html_to_pdf(_body_html, _regen_path)
                if _regen_ok and os.path.isfile(_regen_path):
                    pdf_path = _regen_path
                    _email_pdf_found = True
                    print(f"  [AWEK-EXCEPTION] {spv_name_tmp}: ✅ email.pdf regenerated locally")
                    # Upload to SharePoint for future retries
                    if booking_dir:
                        try:
                            sp.upload_file_to_sp(_regen_path, booking_dir, "email.pdf")
                            print(f"  [AWEK-EXCEPTION] {spv_name_tmp}: ✅ email.pdf uploaded to SharePoint")
                        except Exception as _upl_exc:
                            print(f"  [AWEK-EXCEPTION] {spv_name_tmp}: ⚠ SP upload failed: {_upl_exc}")
                else:
                    print(f"  [AWEK-EXCEPTION] {spv_name_tmp}: ❌ email.pdf regeneration failed — "
                          f"HTML-to-PDF conversion failed. Proceeding without PDF upload.")
                    pdf_path = ""
            else:
                print(f"  [AWEK-EXCEPTION] {spv_name_tmp}: ❌ email.pdf not found and no body_html in DB — "
                      f"proceeding without PDF upload.")
                pdf_path = ""
    # ────────────────────────────────────────────────────────────────────────

    nis_row = dict(spv_row)
    nis_row["spv_name"] = spv_name_tmp
    for fk in negative: nis_row[fk] = 0
    _meta = metadata_row or {}
    from_date = str(_meta.get("from_date","")).strip()
    to_date = str(_meta.get("to_date","")).strip()
    purpose = f"{from_date} TO {to_date} DSM charges" if from_date and to_date else "NIS booking for DSM"
    total_amount = sum(positive.values())
    while True:
        current = db.get_step_retry_count(run_id, step) + 1
        db.step_start(run_id, step, detail=f"attempt={current}/{MAX_STEP_RETRIES} total={total_amount}")
        try:
            print(f"\n  ▶ NIS booking  [{label}]  total={total_amount}  (attempt {current}/{MAX_STEP_RETRIES})")
            print(f"  [PDF-UPLOAD] PDF file for this row: {pdf_path or '(none — no PDF will be uploaded)'}")
            # 3-row rule: use the invoice number pre-resolved for this row (if any).
            row_invoice_no = ""
            if booking_row_id is not None:
                row_invoice_no = str((db.get_booking_row(booking_row_id) or {}).get("invoice_number") or "").strip()
            nis_number = book_nis(pdf_path=pdf_path, company_code=vendor_data.get("company_code",""),
                vendor_code=vendor_data.get("vendor_code",""), bank_key=vendor_data.get("bank_key",""),
                invoice_date=invoice_date, cost_center=vendor_data.get("cost_center",""),
                plant=vendor_data.get("plant",""), purpose=purpose, svp_dsm_row=nis_row,
                login_barrier=login_barrier, expense_items=expense_items,
                invoice_number=row_invoice_no)
            if nis_number: db.save_nis_checklist_id(run_id, nis_number, append=bool(row_index))
            if booking_row_id is not None and nis_number:
                db.save_booking_row_nis_checklist_id(booking_row_id, nis_number)
            # Share the captured NIS number with every cost-center-grouped row
            # booked together in this single submission.
            if nis_number:
                for _eid in extra_booking_row_ids:
                    db.save_booking_row_nis_checklist_id(_eid, nis_number)
            db.save_downloaded_pdfs(run_id, [pdf_path] if pdf_path else pdf_files)

            if not nis_number:
                # NIS ran but booking number was NOT captured. Retry to get it.
                print(f"  ⚠️  NIS ran but booking number NOT captured — will retry.")
                new_count = db.increment_step_retry(run_id, step,
                    "NIS booking submitted but NIS number not captured")
                if new_count >= MAX_STEP_RETRIES:
                    # Exhausted retries — mark done with a note so pipeline
                    # doesn't get permanently stuck.
                    print(f"  ⚠️  Max retries reached without capturing NIS number — marking done with note.")
                    row_detail = f"row={row_index} " if row_index else ""
                    db.step_done(run_id, step, detail=f"{row_detail}nis_number=NOT_CAPTURED")
                    print(f"  ✅ NIS done  [{label}]  NIS=(NOT CAPTURED)")
                    return
                time.sleep(10)
                continue  # retry the while loop

            row_detail = f"row={row_index} " if row_index else ""
            db.step_done(run_id, step, detail=f"{row_detail}nis_number={nis_number}")
            print(f"  ✅ NIS done  [{label}]  NIS={nis_number}")
            return
        except Exception as exc:
            new_count = db.increment_step_retry(run_id, step, str(exc))
            if new_count >= MAX_STEP_RETRIES:
                final = f"NIS [{label}] failed after {MAX_STEP_RETRIES} attempts: {str(exc)[:200]}"
                db.step_failed(run_id, step, final)
                # NOTE: do NOT send a per-row failure email here. Failures are
                # collected and emailed ONCE per email/run in _finalize_run,
                # aggregating every SPV row and its status.
                raise RuntimeError(final) from exc
            time.sleep(5 * current)

def _step_update_tracker(run_id, metadata_row, spv_row, spv_name, step_name=None, booking_row_id=None, allow_duplicate_week=False):
    step = step_name or db.STEP_UPDATE_TRACKER
    if db.is_step_done(run_id, step): return
    # NOTE: In the NIS-only service, we update the tracker with NIS info.
    # SAP checklist info will be filled by sap_automation when it runs.
    db.step_start(run_id, step, detail=f"sheet={spv_name}")
    try:
        run_row = db.get_run_by_id(run_id)
        # Prefer THIS row's own NIS number. The run-level nis_checklist_id is an
        # aggregate of every row in the email, so using it per row would write
        # the wrong number. Fall back to the run value only if the row has none.
        checklist_id = ""
        if booking_row_id is not None:
            brow = db.get_booking_row(booking_row_id) or {}
            checklist_id = str(brow.get("nis_checklist_id") or "").strip()
        if not checklist_id:
            checklist_id = run_row.get("nis_checklist_id") or ""
        sap_checklist_number = run_row.get("checklist_value") or ""
        email_received_date = ""
        raw_received = run_row.get("received_at") or ""
        if raw_received:
            try:
                dt = datetime.strptime(raw_received.replace("Z","")[:19], "%Y-%m-%dT%H:%M:%S")
                email_received_date = dt.strftime("%d-%b-%Y")
            except: email_received_date = raw_received[:10]
        nis_booking_date = datetime.utcnow().strftime("%d-%b-%Y")
        # Serialize the shared-tracker read-modify-write so parallel workers
        # can't race and create duplicate rows.
        with _TRACKER_LOCK:
            add_incremental_week_row(metadata_row=metadata_row, spv_dsm_row=spv_row, sheet_name=spv_name,
                checklist_id=checklist_id, sap_checklist_number=sap_checklist_number,
                email_received_date=email_received_date, nis_booking_date=nis_booking_date,
                allow_duplicate_week=allow_duplicate_week)
        db.step_done(run_id, step, detail=f"week={metadata_row.get('week_no')} spv={spv_name}")
    except Exception as exc:
        db.step_failed(run_id, step, str(exc)); raise


# ── Pipeline runner ──────────────────────────────────────────────────────────
def _maybe_extract_3row_invoices(spv_dsm_rows, attachment_dir, email_result):
    """3-row rule: resolve each row's invoice number from the single attached PDF.

    * INOX bill  → one invoice per SPV page; each row gets the invoice number
      from the page matching its SPV name.
    * non-INOX   → single-page bill; the one invoice number is shared by all rows.

    Only applies when the table has exactly 3 valid SPV rows and a PDF is
    attached. Returns {spv_name: invoice_number}; {} when the rule doesn't apply.
    Must be called BEFORE the attachment dir is cleaned up.
    """
    try:
        rows = [r for r in (spv_dsm_rows or [])
                if str((r or {}).get("spv_name") or "").strip()]
        if len(rows) != 3:
            return {}
        pdfs = [p for p in ((email_result or {}).get("downloaded_files") or [])
                if str(p).lower().endswith(".pdf")]
        if not pdfs and attachment_dir and os.path.isdir(attachment_dir):
            pdfs = [str(p) for p in Path(attachment_dir).glob("*.pdf")]
        if not pdfs:
            print("  [3ROW-INVOICE] 3-row table but no attached PDF — skipping pre-extraction.")
            return {}
        from invoice_extractor import extract_invoice_numbers_for_rows
        inv_map = extract_invoice_numbers_for_rows(pdfs[0], rows)
        print(f"  [3ROW-INVOICE] Resolved invoice numbers: {inv_map}")
        return inv_map or {}
    except Exception as exc:
        print(f"  [3ROW-INVOICE] Invoice pre-extraction failed: {exc}")
        return {}


def _prepare_run(run_id, email_result=None) -> bool:
    """Prepare ONE email run up to the point its booking rows exist.

    Does everything that does NOT need the NIS browser: read-email, extraction,
    save the email + PDFs to the booking folder, classification, and seeding one
    independent booking row per SPV row (with its classification snapshot).

    The actual NIS bookings are performed later by the parallel booking driver,
    pulling rows from the global pool. Returns True on success.
    """
    db.mark_run_started(run_id)
    run = db.get_run_by_id(run_id)
    # ── Skip re-extraction if this email's rows are already in the DB ─────────
    # Once an email has been extracted and its booking rows seeded, never parse
    # / seed it again — reuse the rows already persisted in SQLite.
    existing_rows = db.get_booking_rows(run_id)
    if existing_rows and db.is_step_done(run_id, db.STEP_EMAIL_EXTRACTION):
        print(f"  [PREPARE-SKIP] run_id={run_id} already extracted & seeded "
              f"({len(existing_rows)} row(s) in DB) — reusing, not re-extracting.")
        return True
    print(f"\n{'─'*65}\n  PREPARE  run_id={run_id}  →  {run['subject']}\n{'─'*65}")
    try:
        if not db.is_step_done(run_id, db.STEP_READ_EMAIL):
            db.step_start(run_id, db.STEP_READ_EMAIL)
            db.step_done(run_id, db.STEP_READ_EMAIL, detail=f"msg_id={run['message_id'][:30]}")
        body_html = run.get("body_html") or (email_result or {}).get("body_html", "")
        attachment_dir = (email_result or {}).get("attachment_dir", "")

        metadata_row, spv_dsm_rows = _step_extraction(
            run_id,
            body_html,
            parsed_tables=(email_result or {}).get("parsed_tables"),
        )
        _ext_from = str(metadata_row.get("from_date","")).strip()
        _ext_to = str(metadata_row.get("to_date","")).strip()
        _ext_spv = ""
        for _r in (spv_dsm_rows or []):
            _ext_spv = str(_r.get("spv_name","")).strip()
            if _ext_spv: break
        if _ext_spv: lm.set_context(_ext_spv, _ext_from, _ext_to)

        # 3-row rule: pre-resolve per-row invoice numbers from the attached PDF
        # BEFORE the attachment dir is cleaned by _save_email_to_booking_folder.
        row_invoice_map = _maybe_extract_3row_invoices(spv_dsm_rows, attachment_dir, email_result)

        booking_folder = str(run.get("booking_folder_path") or "").strip()
        folder_map = {}  # {label: sp_folder_path} — populated below
        if not booking_folder and _has_all_extracted_values(metadata_row, spv_dsm_rows):
            folder_map = _save_email_to_booking_folder(metadata_row, spv_dsm_rows, body_html, attachment_dir, run.get("subject",""))
            # Store the first folder as the run-level booking_folder (backward compat)
            if folder_map:
                booking_folder = next(iter(folder_map.values()))
                db.save_booking_folder_path(run_id, booking_folder)

        row_classifications = _step_classification_all(run_id, spv_dsm_rows)

        # Persist every extracted SPV as its own row. Booking units are formed
        # later by grouping these rows on same-email cost center.
        message_id = str(run.get("message_id") or run_id)
        db.seed_booking_rows(run_id, message_id, spv_dsm_rows, metadata_row)

        # Persist booking folder + classification snapshot on each pending row so
        # the booking workers can run fully from the DB row, with no shared state.
        for row_index, (spv_row, positive, negative) in enumerate(row_classifications, start=1):
            brow = db.get_booking_row_by_index(message_id, row_index) or {}
            brow_id = brow.get("row_id")
            if brow_id is None:
                continue
            # Assign per-SPV folder path: each SPV gets its own folder (full name).
            # Even SPVs booked together (same cost center) have separate folders
            # with duplicated documents so each SPV's folder is self-contained.
            if not str(brow.get("booking_folder_path") or "").strip():
                row_spv_name = str(spv_row.get("spv_name") or "").strip()
                row_folder_label = _folder_spv_name(row_spv_name) if row_spv_name else ""
                row_folder = folder_map.get(row_folder_label, booking_folder) if folder_map else booking_folder
                if row_folder:
                    db.save_booking_row_folder(brow_id, row_folder)
            if brow.get("status") not in (db.STATUS_DONE, db.STATUS_SKIPPED):
                db.save_booking_row_classification(brow_id, positive, negative)
            # 3-row rule: persist this row's pre-resolved invoice number.
            inv_no = row_invoice_map.get(str(spv_row.get("spv_name") or "").strip(), "")
            if inv_no and not str(brow.get("invoice_number") or "").strip():
                db.save_booking_row_invoice_number(brow_id, inv_no)

            # ── Pre-load vendor data for ALL rows during prepare ─────────────
            # This ensures vendor_data_json is populated on the booking row
            # immediately, so it's visible in the DB and the SAP service has it
            # available without needing to re-fetch from SharePoint.
            row_spv_name = str(spv_row.get("spv_name") or "").strip()
            if row_spv_name and not _json_dict(brow.get("vendor_data_json")):
                try:
                    vendor_data = _step_vendor_master_for_row(run_id, row_spv_name, row_index)
                    db.save_booking_row_vendor(brow_id, vendor_data)
                except Exception as vend_exc:
                    print(f"  [PREPARE-VENDOR] ⚠ Could not load vendor data for "
                          f"row {row_index} '{row_spv_name}': {vend_exc}")
                    print(f"  [PREPARE-VENDOR]   Reason: {str(vend_exc)[:200]}")
                    print(f"  [PREPARE-VENDOR]   This row's vendor_data_json will remain "
                          f"empty until booking time (will retry then).")

            # ── Pre-compute invoice_date for ALL rows during prepare ──────────
            # invoice_date is derived from metadata (dsm_statement_published_date
            # or from_date). Compute and persist it now so all booking rows have
            # this field populated immediately.
            if not str(brow.get("invoice_date") or "").strip():
                inv_date = _format_invoice_date(
                    metadata_row.get("dsm_statement_published_date")
                    or metadata_row.get("from_date")
                    or "10.02.2020"
                )
                if inv_date:
                    db.save_booking_row_invoice_date(brow_id, inv_date)

        # ── Field completeness check: highlight any missing required fields ──
        # Re-read each row from DB and warn if any field required by NIS/SAP
        # booking is still empty. This makes issues visible immediately.
        _required_fields = [
            ("vendor_data_json", "Vendor data (vendor code, company, cost center, plant)"),
            ("invoice_date", "Invoice date (from metadata: dsm_statement_published_date or from_date)"),
            ("booking_folder_path", "SharePoint booking folder path"),
            ("positive_fields_json", "Positive amount classification"),
            ("metadata_json", "Email metadata (week, dates, due date)"),
            ("spv_data_json", "SPV row data (charges, amounts)"),
        ]
        for row_index, (spv_row, positive, negative) in enumerate(row_classifications, start=1):
            brow = db.get_booking_row_by_index(message_id, row_index) or {}
            brow_id = brow.get("row_id")
            if brow_id is None:
                continue
            row_spv = str(spv_row.get("spv_name") or "").strip()
            missing = []
            for field, label in _required_fields:
                val = brow.get(field)
                if field == "vendor_data_json":
                    if not _json_dict(val).get("vendor_code"):
                        missing.append(label)
                elif field in ("positive_fields_json", "metadata_json", "spv_data_json"):
                    if not _json_dict(val):
                        missing.append(label)
                else:
                    if not str(val or "").strip():
                        missing.append(label)
            if missing:
                print(f"  ⚠️  [PREPARE-CHECK] Row {row_index} '{row_spv}' has EMPTY fields:")
                for m in missing:
                    print(f"       • {m}")
                print(f"       These will be retried during booking, but may cause failures if unresolved.")

        # If this email's SPV table has the trigger row-count (default 5),
        # register a debit/credit email to the finance team (sent + tracked
        # separately so it is retryable and replies can be matched back).
        _maybe_register_debit_credit_email(run_id, metadata_row, spv_dsm_rows, booking_folder)

        lm.clear_context()
        print(f"✅  Prepared run_id={run_id}  ({len(row_classifications)} row(s) seeded)\n")
        return True
    except Exception as exc:
        db.mark_run_failed(run_id, str(exc))
        fail_spv, fail_from, fail_to = _failure_context_from_run(run_id)
        _send_failure_email(run_id, fail_spv, fail_from, fail_to, "nis_prepare", str(exc))
        db.print_run_summary(run_id)
        lm.clear_context()
        print(f"❌  Prepare FAILED  run_id={run_id}\n{traceback.format_exc()}")
        return False


def _process_booking_row(row_id, login_barrier=None) -> tuple:
    """Book ONE independent SPV row end-to-end. Safe to run in parallel.

    Everything is reconstructed from the DB row, runs in a dedicated working
    directory and uses the thread-local logging context — so concurrent bookings
    never overlap. *login_barrier* (when given) makes this session wait until
    all parallel sessions have logged in before it starts filling the form.
    Returns (status, row_id, spv_name).
    """
    row = db.get_booking_row(row_id)
    if not row:
        return ("missing", row_id, "")
    if row.get("status") in (db.STATUS_DONE, db.STATUS_SKIPPED, db.STATUS_NIS_DONE):
        return ("skipped", row_id, str(row.get("spv_name") or ""))

    run_id = str(row.get("run_id") or "")
    row_index = int(row.get("row_index") or 0)
    spv_name = str(row.get("spv_name") or "").strip()
    metadata_row = _json_dict(row.get("metadata_json"))
    spv_row = _json_dict(row.get("spv_data_json"))
    if spv_name:
        spv_row["spv_name"] = spv_name
    positive = _json_dict(row.get("positive_fields_json"))
    negative = _json_dict(row.get("negative_fields_json"))
    booking_folder = str(row.get("booking_folder_path") or "").strip()
    if not booking_folder:
        run = db.get_run_by_id(run_id) or {}
        booking_folder = str(run.get("booking_folder_path") or "").strip()

    from_date = str(metadata_row.get("from_date","")).strip()
    to_date = str(metadata_row.get("to_date","")).strip()
    lm.set_context(spv_name, from_date, to_date)  # thread-local → per-worker

    # Older seeds may not have a stored classification — recompute if needed.
    if not positive and not negative:
        positive, negative = _classify_amounts(spv_row)

    work_dir = _row_tmp_dir(row_id)
    try:
        db.mark_booking_row_started(row_id)
        db.save_spv_name(run_id, spv_name)
        db.save_booking_row_classification(row_id, positive, negative)

        # ── Pre-download the invoice PDF into this row's work_dir ────────────
        # The work_dir is a fresh empty directory. _step_nis_booking needs a PDF
        # to upload. Download it from the SharePoint booking folder NOW so it's
        # available regardless of which path _step_nis_booking takes.
        #
        # RULE: Every row MUST have a PDF to upload in the NIS portal.
        #   Priority: 1) Invoice attachment PDF  2) email.pdf (email converted to PDF)
        #   For AWEK SPVs: always use email.pdf (invoice PDF is suppressed).
        if booking_folder and positive:
            try:
                _is_awek = is_email_pdf_spv(spv_name)
                _pdf_downloaded = False

                if _is_awek:
                    # AWEK exception: always use email.pdf
                    _email_pdf_dest = os.path.join(work_dir, "email.pdf")
                    if not os.path.isfile(_email_pdf_dest):
                        _dl = sp.download_file_from_sp(
                            f"{booking_folder}/email.pdf", _email_pdf_dest
                        )
                        if _dl and os.path.isfile(_dl):
                            _pdf_downloaded = True
                            print(f"  [PDF] Downloaded email.pdf for {spv_name} (AWEK rule)")
                    else:
                        _pdf_downloaded = True

                    # Fallback: if email.pdf doesn't exist on SharePoint,
                    # regenerate it from the email body HTML stored in the DB.
                    if not _pdf_downloaded:
                        print(f"  [PDF] email.pdf not on SharePoint for {spv_name} — regenerating from email body HTML")
                        _run_data = db.get_run_by_id(run_id) or {}
                        _body_html = str(_run_data.get("body_html") or "").strip()
                        if _body_html:
                            _created = _convert_html_to_pdf(_body_html, _email_pdf_dest)
                            if _created and os.path.isfile(_email_pdf_dest):
                                _pdf_downloaded = True
                                print(f"  [PDF] ✅ email.pdf regenerated locally for {spv_name}")
                                # Also upload to SharePoint so future retries don't need to regenerate
                                try:
                                    sp.upload_file_to_sp(_email_pdf_dest, booking_folder, "email.pdf")
                                    print(f"  [PDF] ✅ email.pdf uploaded to SharePoint: {booking_folder}/email.pdf")
                                except Exception as _upload_exc:
                                    print(f"  [PDF] ⚠ Could not upload regenerated email.pdf to SP: {_upload_exc}")
                            else:
                                print(f"  [PDF] ❌ email.pdf regeneration failed for {spv_name} — HTML-to-PDF conversion failed")
                        else:
                            print(f"  [PDF] ❌ No body_html in DB for run_id={run_id} — cannot regenerate email.pdf")
                else:
                    # Normal: try invoice attachment PDF first
                    _sp_items = sp.list_sp_folder(booking_folder)
                    _inv_pdfs = sorted(
                        [f["name"] for f in _sp_items
                         if not f["is_folder"] and _is_invoice_upload_pdf_name(f["name"])],
                        key=str.lower,
                    )
                    if _inv_pdfs:
                        _pdf_dest = os.path.join(work_dir, _inv_pdfs[0])
                        _dl = sp.download_file_from_sp(
                            f"{booking_folder}/{_inv_pdfs[0]}", _pdf_dest
                        )
                        if _dl and os.path.isfile(_dl):
                            _pdf_downloaded = True
                            print(f"  [PDF] Downloaded {_inv_pdfs[0]} for {spv_name}")

                    # Fallback: any PDF that's not a checklist
                    if not _pdf_downloaded:
                        _any_pdfs = sorted(
                            [f["name"] for f in _sp_items
                             if not f["is_folder"] and f["name"].lower().endswith(".pdf")
                             and not f["name"].lower().startswith(("sap_checklist_", "checklist_", "email"))],
                            key=str.lower,
                        )
                        if _any_pdfs:
                            _pdf_dest = os.path.join(work_dir, _any_pdfs[0])
                            _dl = sp.download_file_from_sp(
                                f"{booking_folder}/{_any_pdfs[0]}", _pdf_dest
                            )
                            if _dl and os.path.isfile(_dl):
                                _pdf_downloaded = True
                                print(f"  [PDF] Downloaded {_any_pdfs[0]} (fallback) for {spv_name}")

                    # Ultimate fallback: use email.pdf (always exists in the booking folder)
                    if not _pdf_downloaded:
                        _email_pdf_dest = os.path.join(work_dir, "email.pdf")
                        _dl = sp.download_file_from_sp(
                            f"{booking_folder}/email.pdf", _email_pdf_dest
                        )
                        if _dl and os.path.isfile(_dl):
                            _pdf_downloaded = True
                            print(f"  [PDF] Downloaded email.pdf as fallback for {spv_name} (no invoice PDF found)")

                if not _pdf_downloaded:
                    print(f"  [PDF] ⚠ WARNING: No PDF could be downloaded for {spv_name} — booking may proceed without upload")
            except Exception as _pdf_dl_err:
                print(f"  [PDF] Pre-download failed for {spv_name}: {_pdf_dl_err}")

        validation_step = _row_step_name(db.STEP_EMAIL_VALIDATION, row_index, spv_name)
        _step_validation(run_id, metadata_row, spv_row, positive, negative, step_name=validation_step)

        invoice_date = _format_invoice_date(metadata_row.get("dsm_statement_published_date") or metadata_row.get("from_date") or "10.02.2020")
        db.save_booking_row_invoice_date(row_id, invoice_date)

        if negative:
            print(f"  [ROW {row_index}] {spv_name}: SAP tasks queued for {list(negative.keys())} (handled by sap_automation).")
            # ── Create pending SAP steps so run_sap_automation finds them ─────
            # Each per-row SAP step is named:  sap_automation_<field>_row_<n>_<spv>
            # The SAP service queries run_steps for all sap_automation_* steps
            # that are not done/skipped.  Without this block those step rows are
            # never inserted, so the SAP runner always finds nothing.
            for fk, amt in negative.items():
                if abs(float(amt or 0)) <= 0:
                    continue
                sap_step = _row_step_name(db.step_sap(fk), row_index, spv_name)
                db.ensure_dynamic_step(run_id, sap_step)
                print(f"    [SAP-QUEUE] Created pending step: {sap_step}")
        else:
            # No negative fields — skip any pre-existing (legacy) run-level SAP steps.
            for fk, _ in SPV_AMOUNT_FIELDS:
                s = _row_step_name(db.step_sap(fk), row_index, spv_name)
                step_row = db.get_step(run_id, s)
                if step_row and not db.is_step_done(run_id, s):
                    db.step_skip(run_id, s, "no negative fields")

        # ── Load vendor data for ALL rows (positive AND negative) ────────────
        # Vendor data is needed by the SAP service for negative amounts and by
        # the NIS booking for positive amounts. Load it unconditionally so the
        # booking_row always has vendor_data_json populated.
        vendor_data = _step_vendor_master_for_row(run_id, spv_name, row_index)
        db.save_booking_row_vendor(row_id, vendor_data)

        if not vendor_data or not vendor_data.get("vendor_code"):
            print(f"  ⚠️  [VENDOR-EMPTY] Row {row_index} '{spv_name}': vendor_data_json is EMPTY!")
            print(f"  ⚠️  [VENDOR-EMPTY]   Reason: SPV name '{spv_name}' could not be matched "
                  f"in DSM-NIS-INPUT.xlsx (sheet '{getattr(config, 'INPUT_SHEET_NAME', 'main')}').")
            print(f"  ⚠️  [VENDOR-EMPTY]   Action needed: Add this SPV to the INPUT file or "
                  f"check for name spelling differences.")

        if positive:
            nis_step = _row_step_name(db.STEP_NIS_BOOKING, row_index, spv_name)
            _step_nis_booking(run_id, positive, spv_row, negative, vendor_data,
                invoice_date, "", metadata_row=metadata_row,
                booking_dir=booking_folder, step_name=nis_step, row_index=row_index,
                booking_row_id=row_id, work_dir=work_dir, login_barrier=login_barrier)
            tracker_step = _row_step_name(db.STEP_UPDATE_TRACKER, row_index, spv_name)
            _step_update_tracker(run_id, metadata_row, spv_row, spv_name, step_name=tracker_step, booking_row_id=row_id)
            if negative:
                # Row has BOTH +ve and -ve: NIS is done but SAP steps are pending.
                # Mark as nis_done; SAP automation will finalize to 'done' once
                # all SAP steps complete and checklist values are captured.
                db.mark_booking_row_nis_done(row_id)
                print(f"  ✅ [ROW {row_index}] {spv_name} NIS booked (awaiting SAP automation for {list(negative.keys())}).")
                return ("nis_done", row_id, spv_name)
            else:
                # Row has only +ve: fully done after NIS.
                db.mark_booking_row_done(row_id)
                print(f"  ✅ [ROW {row_index}] {spv_name} booked.")
                return ("done", row_id, spv_name)
        else:
            nis_step = _row_step_name(db.STEP_NIS_BOOKING, row_index, spv_name)
            db.ensure_dynamic_step(run_id, nis_step)
            if not db.is_step_done(run_id, nis_step) and not db.is_step_skipped(run_id, nis_step):
                db.step_skip(run_id, nis_step, "no positive fields")
            tracker_step = _row_step_name(db.STEP_UPDATE_TRACKER, row_index, spv_name)
            _step_update_tracker(run_id, metadata_row, spv_row, spv_name, step_name=tracker_step, booking_row_id=row_id)
            if negative:
                # Row has only -ve values: NIS skipped, SAP pending.
                # Mark as nis_done (NIS part complete); SAP automation will
                # finalize to 'done' once all SAP steps are complete.
                db.mark_booking_row_nis_done(row_id)
                print(f"  ⏭  [ROW {row_index}] {spv_name} NIS skipped (awaiting SAP automation for {list(negative.keys())}).")
                return ("nis_done", row_id, spv_name)
            else:
                # Row has neither +ve nor -ve (all zeros) — truly skipped.
                db.mark_booking_row_skipped(row_id, "no positive fields")
                print(f"  ⏭  [ROW {row_index}] {spv_name} skipped (no positive or negative fields).")
                return ("skipped", row_id, spv_name)
    except Exception as row_exc:
        db.increment_booking_row_retry(row_id, str(row_exc))
        db.mark_booking_row_failed(row_id, str(row_exc))
        print(f"  ❌ [ROW {row_index}] {spv_name} booking failed: {str(row_exc)[:160]}")
        return ("failed", row_id, spv_name)
    finally:
        _cleanup_tmp(work_dir)
        lm.clear_context()


def _process_booking_unit(unit, login_barrier=None) -> tuple:
    """Process one booking unit. A unit is either a single SPV row (unchanged
    behaviour) or rows from one email sharing a cost center that must be booked
    together as ONE NIS booking with their combined expense items.
    """
    if not unit:
        return ("missing", None, "")
    if len(unit) == 1:
        return _process_booking_row(unit[0].get("row_id"), login_barrier)
    return _process_booking_group(unit, login_barrier)


def _process_booking_group(rows, login_barrier=None) -> tuple:
    """Book rows from one email with the same cost center as ONE NIS booking.

    All rows share a single captured NIS number, while each row still gets its
    own vendor read, expense items, and tracker update.
    Returns (status, primary_row_id, group_label).
    """
    rows = sorted(rows, key=lambda r: int(r.get("row_index") or 0))
    primary = rows[0]
    primary_row_id = primary.get("row_id")
    run_id = str(primary.get("run_id") or "")
    group_label = ", ".join(str(r.get("spv_name") or "") for r in rows)

    booking_folder = str(primary.get("booking_folder_path") or "").strip()
    if not booking_folder:
        run = db.get_run_by_id(run_id) or {}
        booking_folder = str(run.get("booking_folder_path") or "").strip()

    base_meta = _json_dict(primary.get("metadata_json"))
    from_date = str(base_meta.get("from_date", "")).strip()
    to_date = str(base_meta.get("to_date", "")).strip()
    lm.set_context(group_label, from_date, to_date)

    work_dir = _row_tmp_dir(primary_row_id)
    try:
        # ── Pre-download the invoice PDF for the group booking ───────────────
        if booking_folder:
            try:
                _sp_items = sp.list_sp_folder(booking_folder)
                _pdf_downloaded = False
                _inv_pdfs = sorted(
                    [f["name"] for f in _sp_items
                     if not f["is_folder"] and _is_invoice_upload_pdf_name(f["name"])],
                    key=str.lower,
                )
                if _inv_pdfs:
                    _pdf_dest = os.path.join(work_dir, _inv_pdfs[0])
                    _dl = sp.download_file_from_sp(
                        f"{booking_folder}/{_inv_pdfs[0]}", _pdf_dest
                    )
                    if _dl and os.path.isfile(_dl):
                        _pdf_downloaded = True
                        print(f"  [PDF] Downloaded {_inv_pdfs[0]} for group {group_label}")

                if not _pdf_downloaded:
                    # Fallback: any PDF that's not a checklist or email
                    _any_pdfs = sorted(
                        [f["name"] for f in _sp_items
                         if not f["is_folder"] and f["name"].lower().endswith(".pdf")
                         and not f["name"].lower().startswith(("sap_checklist_", "checklist_", "email"))],
                        key=str.lower,
                    )
                    if _any_pdfs:
                        _pdf_dest = os.path.join(work_dir, _any_pdfs[0])
                        _dl = sp.download_file_from_sp(
                            f"{booking_folder}/{_any_pdfs[0]}", _pdf_dest
                        )
                        if _dl and os.path.isfile(_dl):
                            _pdf_downloaded = True
                            print(f"  [PDF] Downloaded {_any_pdfs[0]} (fallback) for group {group_label}")

                # Ultimate fallback: email.pdf
                if not _pdf_downloaded:
                    _email_pdf_dest = os.path.join(work_dir, "email.pdf")
                    _dl = sp.download_file_from_sp(
                        f"{booking_folder}/email.pdf", _email_pdf_dest
                    )
                    if _dl and os.path.isfile(_dl):
                        _pdf_downloaded = True
                        print(f"  [PDF] Downloaded email.pdf as fallback for group {group_label}")

                if not _pdf_downloaded:
                    print(f"  [PDF] ⚠ WARNING: No PDF could be downloaded for group {group_label}")
            except Exception as _pdf_dl_err:
                print(f"  [PDF] Pre-download failed for group {group_label}: {_pdf_dl_err}")

        # ── Reconstruct + validate every row in the group ────────────────────
        per_row = []
        for r in rows:
            rid = r.get("row_id")
            r_index = int(r.get("row_index") or 0)
            r_spv_name = str(r.get("spv_name") or "").strip()
            r_meta = _json_dict(r.get("metadata_json")) or base_meta
            r_spv = _json_dict(r.get("spv_data_json"))
            if r_spv_name:
                r_spv["spv_name"] = r_spv_name
            r_pos = _json_dict(r.get("positive_fields_json"))
            r_neg = _json_dict(r.get("negative_fields_json"))
            if not r_pos and not r_neg:
                r_pos, r_neg = _classify_amounts(r_spv)

            db.mark_booking_row_started(rid)
            db.save_spv_name(run_id, r_spv_name)
            db.save_booking_row_classification(rid, r_pos, r_neg)

            validation_step = _row_step_name(db.STEP_EMAIL_VALIDATION, r_index, r_spv_name)
            _step_validation(run_id, r_meta, r_spv, r_pos, r_neg, step_name=validation_step)

            r_invoice_date = _format_invoice_date(
                r_meta.get("dsm_statement_published_date") or r_meta.get("from_date") or "10.02.2020")
            db.save_booking_row_invoice_date(rid, r_invoice_date)

            if r_neg:
                print(f"  [GROUP ROW {r_index}] {r_spv_name}: SAP tasks queued for "
                      f"{list(r_neg.keys())} (handled by sap_automation).")
                # ── Create pending SAP steps so run_sap_automation finds them ─
                for fk, amt in r_neg.items():
                    if abs(float(amt or 0)) <= 0:
                        continue
                    sap_step = _row_step_name(db.step_sap(fk), r_index, r_spv_name)
                    db.ensure_dynamic_step(run_id, sap_step)
                    print(f"    [SAP-QUEUE] Created pending step: {sap_step}")
            else:
                for fk, _ in SPV_AMOUNT_FIELDS:
                    s = db.step_sap(fk)
                    step_row = db.get_step(run_id, s)
                    if step_row and not db.is_step_done(run_id, s):
                        db.step_skip(run_id, s, "no negative fields")

            per_row.append({
                "row": r, "rid": rid, "row_index": r_index, "spv_name": r_spv_name,
                "meta": r_meta, "spv": r_spv, "positive": r_pos, "negative": r_neg,
                "invoice_date": r_invoice_date,
            })

        positive_entries = [e for e in per_row if e["positive"]]

        # ── Whole group has nothing to book ──────────────────────────────────
        if not positive_entries:
            for e in per_row:
                nis_step = _row_step_name(db.STEP_NIS_BOOKING, e["row_index"], e["spv_name"])
                db.ensure_dynamic_step(run_id, nis_step)
                if not db.is_step_done(run_id, nis_step) and not db.is_step_skipped(run_id, nis_step):
                    db.step_skip(run_id, nis_step, "no positive fields")
                tracker_step = _row_step_name(db.STEP_UPDATE_TRACKER, e["row_index"], e["spv_name"])
                _step_update_tracker(run_id, e["meta"], e["spv"], e["spv_name"],
                                     step_name=tracker_step, booking_row_id=e["rid"],
                                     allow_duplicate_week=True)
                db.mark_booking_row_skipped(e["rid"], "no positive fields")
            print(f"  ⏭  [GROUP] {group_label} skipped (no positive fields).")
            return ("skipped", primary_row_id, group_label)

        # ── Read vendor data per positive variant + build combined items ─────
        purpose = f"{from_date} TO {to_date} DSM charges" if from_date and to_date else "NIS booking for DSM"
        expense_items = []
        for e in positive_entries:
            vendor_data = _step_vendor_master_for_row(run_id, e["spv_name"], e["row_index"])
            db.save_booking_row_vendor(e["rid"], vendor_data)
            e["vendor"] = vendor_data
            for fk, _lbl in SPV_AMOUNT_FIELDS:
                if fk not in e["positive"]:
                    continue
                expense_items.append({
                    "amount": e["positive"][fk],
                    "label": f"{e['spv_name']} - {NIS_ITEM_LABELS.get(fk, fk)}",
                    "cost_center": vendor_data.get("cost_center", ""),
                    "plant": vendor_data.get("plant", ""),
                    "purpose": purpose,
                    "additional_remarks": "PSS8",
                    "field_key": fk,
                })

        primary_positive = positive_entries[0]
        shared_vendor = primary_positive["vendor"]
        combined_positive = _sum_amount_maps([e["positive"] for e in positive_entries])
        extra_ids = [e["rid"] for e in positive_entries[1:]]

        print(f"\n  ══ GROUP BOOKING {group_label} — {len(expense_items)} expense item(s) "
              f"in ONE NIS booking ══")
        nis_step = _row_step_name(db.STEP_NIS_BOOKING, primary_positive["row_index"],
                                  primary_positive["spv_name"])
        _step_nis_booking(run_id, combined_positive, primary_positive["spv"], {}, shared_vendor,
            primary_positive["invoice_date"], "", metadata_row=primary_positive["meta"],
            booking_dir=booking_folder, step_name=nis_step,
            row_index=primary_positive["row_index"], booking_row_id=primary_positive["rid"],
            work_dir=work_dir, login_barrier=login_barrier,
            expense_items=expense_items, extra_booking_row_ids=extra_ids)

        # ── Propagate the NIS number, update each tracker, finalise rows ─────
        primary_row = db.get_booking_row(primary_positive["rid"]) or {}
        nis_number = str(primary_row.get("nis_checklist_id") or "").strip()
        for e in per_row:
            if nis_number and not str((db.get_booking_row(e["rid"]) or {}).get("nis_checklist_id") or "").strip():
                db.save_booking_row_nis_checklist_id(e["rid"], nis_number)
            nis_step_row = _row_step_name(db.STEP_NIS_BOOKING, e["row_index"], e["spv_name"])
            tracker_step = _row_step_name(db.STEP_UPDATE_TRACKER, e["row_index"], e["spv_name"])
            _step_update_tracker(run_id, e["meta"], e["spv"], e["spv_name"],
                                 step_name=tracker_step, booking_row_id=e["rid"],
                                 allow_duplicate_week=True)
            if e["positive"]:
                # The primary's NIS step is closed by _step_nis_booking; close the
                # secondary variants' per-row steps too (they were booked together).
                if not db.is_step_done(run_id, nis_step_row):
                    db.step_done(run_id, nis_step_row,
                                 detail=f"booked_with_group nis_number={nis_number or 'n/a'}")
                if e["negative"]:
                    # Row has BOTH +ve and -ve: NIS done, awaiting SAP.
                    db.mark_booking_row_nis_done(e["rid"])
                else:
                    # Row has only +ve: fully done.
                    db.mark_booking_row_done(e["rid"])
            else:
                db.ensure_dynamic_step(run_id, nis_step_row)
                if not db.is_step_done(run_id, nis_step_row) and not db.is_step_skipped(run_id, nis_step_row):
                    db.step_skip(run_id, nis_step_row, "no positive fields")
                if e["negative"]:
                    # Row has only -ve: NIS skipped, awaiting SAP.
                    db.mark_booking_row_nis_done(e["rid"])
                else:
                    # Row has neither — truly skipped.
                    db.mark_booking_row_skipped(e["rid"], "no positive fields")
        any_awaiting_sap = any(e["negative"] for e in per_row)
        if any_awaiting_sap:
            print(f"  ✅ [GROUP] {group_label} NIS booked (NIS={nis_number or 'n/a'}, some rows awaiting SAP).")
            return ("nis_done", primary_row_id, group_label)
        print(f"  ✅ [GROUP] {group_label} booked together (NIS={nis_number or 'n/a'}).")
        return ("done", primary_row_id, group_label)
    except Exception as grp_exc:
        for r in rows:
            db.increment_booking_row_retry(r.get("row_id"), str(grp_exc))
            db.mark_booking_row_failed(r.get("row_id"), str(grp_exc))
        print(f"  ❌ [GROUP] {group_label} booking failed: {str(grp_exc)[:160]}")
        return ("failed", primary_row_id, group_label)
    finally:
        _cleanup_tmp(work_dir)
        lm.clear_context()


def _finalize_run(run_id):
    """Roll up per-row outcomes into run-level status once all rows are terminal
    from the NIS service's perspective.

    A row is "NIS-terminal" when it has reached done/skipped/failed OR nis_done
    (NIS booking complete, awaiting SAP). The NIS service marks the run as done
    once all its rows hit one of these states. The SAP service finds pending
    tasks via run_steps directly and does NOT depend on email_runs.status.

    Marks the aggregate steps done/skipped, sets the run done or failed, and
    triggers completion emails. A no-op while any booking row is still pending
    or running (NIS is actively processing).
    """
    rows = db.get_booking_rows(run_id)
    if not rows:
        return
    # nis_done = NIS work complete (awaiting SAP). From the NIS service's
    # perspective this IS a terminal state — the booking is done.
    nis_terminal = (db.STATUS_DONE, db.STATUS_SKIPPED, db.STATUS_FAILED, db.STATUS_NIS_DONE)
    if not all(r.get("status") in nis_terminal for r in rows):
        return  # still rows in flight — finalize later

    positive_rows = [r for r in rows if _json_dict(r.get("positive_fields_json"))]
    failed_rows = [r for r in rows if r.get("status") == db.STATUS_FAILED]
    total = len(rows)

    if positive_rows:
        # A row in nis_done state has had its NIS booking completed successfully
        incomplete = [r for r in positive_rows
                      if r.get("status") not in (db.STATUS_DONE, db.STATUS_NIS_DONE)]
        if not incomplete and not db.is_step_done(run_id, db.STEP_NIS_BOOKING):
            db.step_done(run_id, db.STEP_NIS_BOOKING, detail=f"completed_positive_rows={len(positive_rows)}")
    else:
        step_row = db.get_step(run_id, db.STEP_NIS_BOOKING)
        if step_row and not db.is_step_done(run_id, db.STEP_NIS_BOOKING):
            db.step_skip(run_id, db.STEP_NIS_BOOKING, "no positive rows")

    if not db.is_step_done(run_id, db.STEP_EMAIL_VALIDATION):
        db.step_done(run_id, db.STEP_EMAIL_VALIDATION, detail=f"validated_rows={total}")
    if not db.is_step_done(run_id, db.STEP_READ_VENDOR):
        db.step_done(run_id, db.STEP_READ_VENDOR, detail="vendor data loaded per SPV row")
    if not db.is_step_done(run_id, db.STEP_UPDATE_TRACKER):
        db.step_done(run_id, db.STEP_UPDATE_TRACKER, detail=f"updated_rows={total}")

    if failed_rows:
        detail = "; ".join(
            f"row {r.get('row_index')} ({r.get('spv_name')}): {str(r.get('error_message') or '')[:100]}"
            for r in failed_rows
        )
        db.mark_run_failed(run_id, f"NIS booking failed for {len(failed_rows)} row(s): {detail}")
        fail_spv, fail_from, fail_to = _failure_context_from_run(run_id)
        _send_failure_email(run_id, fail_spv, fail_from, fail_to, "nis_booking_pipeline",
                            f"{len(failed_rows)}/{total} row(s) failed")
    else:
        db.mark_run_done(run_id)

    send_ready_nis_completion_emails()
    db.print_run_summary(run_id)
    print(f"{'✅' if not failed_rows else '⚠️ '}  Run finalized  run_id={run_id} "
          f"(done={total - len(failed_rows)}/{total})\n")


def _row_will_book(row) -> bool:
    """True when this booking row will actually drive an NIS booking (it has
    positive fields). Used to size the login barrier so only booking sessions
    are counted — skipped rows never reach the barrier.
    """
    positive = _json_dict(row.get("positive_fields_json"))
    negative = _json_dict(row.get("negative_fields_json"))
    if not positive and not negative:
        # Legacy/unseeded row → classify from the stored SPV snapshot.
        positive, _ = _classify_amounts(_json_dict(row.get("spv_data_json")))
    return bool(positive)


def _book_pending_rows_parallel(max_workers=None) -> int:
    """Book ALL pending booking rows across every email, up to *max_workers* at
    a time. A free worker slot is filled with work from another email when one
    email doesn't have enough rows. Returns the number of rows processed.

    Rows are grouped into *booking units* by non-empty cost center within each
    email. Every unit is booked in one NIS submission and occupies one worker
    slot; rows with unresolved cost centers remain separate.
    """
    if max_workers is None:
        max_workers = max(1, int(getattr(config, "NIS_PARALLEL_COUNT", 1)))
    max_row_retries = int(getattr(config, "NIS_MAX_ROW_RETRIES", 5))
    stagger = float(getattr(config, "NIS_PARALLEL_LAUNCH_STAGGER", 0) or 0)
    barrier_timeout = float(getattr(config, "NIS_LOGIN_BARRIER_TIMEOUT", 0) or 0)

    rows = db.get_all_pending_booking_rows(max_retries=max_row_retries)
    if not rows:
        return 0

    units = _build_booking_units(rows)

    print(f"\n{'═'*65}\n  BOOKING {len(rows)} pending row(s) in {len(units)} unit(s) — "
          f"up to {max_workers} in parallel\n{'═'*65}")
    touched_runs = set()

    def _unit_runs(unit):
        return {str(r.get("run_id")) for r in unit}

    if max_workers <= 1:
        for unit in units:
            touched_runs.update(_unit_runs(unit))
            _process_booking_unit(unit)
    else:
        # ── Build the login barrier (FIRST WAVE ONLY) ────────────────────────
        # The barrier exists so the initial simultaneous batch all logs in
        # before any of them starts filling. It must apply ONLY to the first
        # wave of booking units (= the units that start together when the pool
        # opens). Units that run later — as worker slots free up — must NOT wait
        # on it, otherwise a smaller final group would block at the barrier
        # until the timeout (the "later automations hang for ages" symptom).
        positive_count = sum(1 for u in units if _unit_will_book(u))
        parties = min(max_workers, positive_count) if positive_count >= 2 else 0
        login_barrier = None
        if barrier_timeout > 0 and parties >= 2:
            login_barrier = threading.Barrier(parties, timeout=barrier_timeout)
            print(f"  [SYNC] Login barrier armed for the first {parties} session(s) "
                  f"(timeout={barrier_timeout:.0f}s). Later units run without waiting.")

        from concurrent.futures import ThreadPoolExecutor, as_completed
        with ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix="nisbook") as ex:
            futures = {}
            barrier_assigned = 0
            for unit in units:
                touched_runs.update(_unit_runs(unit))
                # Hand the barrier only to the first `parties` units that will
                # actually book — they are the simultaneous opening wave.
                use_barrier = None
                if (login_barrier is not None and barrier_assigned < parties
                        and _unit_will_book(unit)):
                    use_barrier = login_barrier
                    barrier_assigned += 1
                futures[ex.submit(_process_booking_unit, unit, use_barrier)] = unit
                # Optional stagger between launches (0 by default → all open at
                # the same time so they log in together).
                if stagger and len(futures) < max_workers:
                    time.sleep(stagger)
            for fut in as_completed(futures):
                try:
                    fut.result()
                except Exception as exc:
                    unit = futures[fut]
                    labels = ", ".join(str(r.get("spv_name") or "") for r in unit)
                    print(f"  ❌ Worker crashed for unit [{labels}]: {exc}")

    # Roll up run-level status for every run we touched.
    for run_id in touched_runs:
        try:
            _finalize_run(run_id)
        except Exception as exc:
            print(f"  [WARN] Finalize failed for run {run_id}: {exc}")

    return len(rows)


def run_pipeline(run_id, email_result=None):
    """Sequential single-email pipeline: prepare → book its rows → finalize.

    Retained for direct/resume callers. The parallel, cross-email path used by
    run_immediate is _prepare_run + _book_pending_rows_parallel + _finalize_run.
    """
    try:
        if not _prepare_run(run_id, email_result=email_result):
            return
        rows = db.get_pending_booking_rows(run_id)
        for unit in _build_booking_units(rows):
            _process_booking_unit(unit)
        _finalize_run(run_id)
    finally:
        _cleanup_tmp(TEMP_BOOKING_ROOT)


# ── Scan & dispatch ──────────────────────────────────────────────────────────
def _process_one_message_batch(messages, headers, base_url):
    """
    Process at most one actionable email from the current fetch.

    Returning after one completed action keeps the service responsive to newly
    arrived checklist/invoice-print emails: the caller immediately fetches the
    mailbox again instead of working through a stale list.
    """
    skipped = 0
    checklist_seen = 0

    for msg in messages:
        message_id, subject = msg["id"], msg.get("subject","")

        checklist_status = _handle_checklist_email(msg, headers, base_url)
        if checklist_status != CHECKLIST_NOT_MATCHED:
            checklist_seen += 1
            if checklist_status == CHECKLIST_PROCESSED:
                return {
                    "did_work": True,
                    "kind": "checklist",
                    "processed": 0,
                    "checklist_processed": 1,
                    "skipped": skipped,
                    "checklist_seen": checklist_seen,
                }
            skipped += 1
            continue

        dc_status = _handle_debit_credit_reply(msg, headers, base_url)
        if dc_status != DC_NOT_MATCHED:
            checklist_seen += 1
            if dc_status == DC_PROCESSED:
                return {
                    "did_work": True,
                    "kind": "debit_credit_reply",
                    "processed": 0,
                    "checklist_processed": 1,
                    "skipped": skipped,
                    "checklist_seen": checklist_seen,
                }
            skipped += 1
            continue

        # ── HOD / FC Approval email detection ────────────────────────────────
        hod_status = _handle_hod_approval_email(msg, headers, base_url)
        if hod_status != APPROVAL_NOT_MATCHED:
            checklist_seen += 1
            if hod_status == APPROVAL_PROCESSED:
                return {
                    "did_work": True,
                    "kind": "hod_approval",
                    "processed": 0,
                    "checklist_processed": 1,
                    "skipped": skipped,
                    "checklist_seen": checklist_seen,
                }
            skipped += 1
            continue

        fc_status = _handle_fc_approval_email(msg, headers, base_url)
        if fc_status != APPROVAL_NOT_MATCHED:
            checklist_seen += 1
            if fc_status == APPROVAL_PROCESSED:
                return {
                    "did_work": True,
                    "kind": "fc_approval",
                    "processed": 0,
                    "checklist_processed": 1,
                    "skipped": skipped,
                    "checklist_seen": checklist_seen,
                }
            skipped += 1
            continue

def _handle_extraction_error(msg, email_result=None, val_result=None, exception_exc=None):
    """Handles cases where an email subject matches DSM-NIS-BOOKING, but extraction

    or mandatory column validation fails. Prints to console, sends email to all
    notification recipients, and records email in DB to prevent duplicate alerts.
    """
    subject = msg.get("subject", "")
    received_at = (email_result or {}).get("received_at") or msg.get("receivedDateTime", "")
    message_id = msg.get("id", "")
    existing_run = db.get_run_by_message_id(message_id)
    extraction_alert_already_sent = (
        str((existing_run or {}).get("status") or "") == "extraction_failed"
    )

    parsed_tables = (email_result or {}).get("parsed_tables", {})
    if not val_result and parsed_tables:
        val_result = validate_mandatory_columns(parsed_tables)
    elif not val_result:
        val_result = {
            "is_valid": False,
            "missing_tables": ["Email Body Tables / Attachment Parsing Failed"],
            "missing_metadata_cols": [],
            "missing_spv_cols": [],
            "error_summary": str(exception_exc) if exception_exc else "Failed to parse email contents",
        }

    # 1. Output to Console
    print("\n" + "=" * 70)
    print(f"❌ [ALERT - EXTRACTION FAILED] Subject: '{subject}'")
    print(f"   Received At : {received_at}")
    print(f"   Message ID  : {message_id}")
    if val_result.get("missing_tables"):
        print(f"   Missing Table(s)            : {', '.join(val_result['missing_tables'])}")
    if val_result.get("missing_metadata_cols"):
        print(f"   Missing Metadata Column(s)  : {', '.join(val_result['missing_metadata_cols'])}")
    if val_result.get("missing_spv_cols"):
        print(f"   Missing SPV DSM Column(s)   : {', '.join(val_result['missing_spv_cols'])}")
    if exception_exc:
        print(f"   Exception                   : {exception_exc}")
    print("=" * 70 + "\n")

    # 2. Build HTML Alert Email for Recipients
    missing_items = []
    if val_result.get("missing_tables"):
        missing_items.append("<li><b>Missing Table(s):</b> " + html.escape(", ".join(val_result["missing_tables"])) + "</li>")
    if val_result.get("missing_metadata_cols"):
        missing_items.append("<li><b>Missing Mandatory Metadata Column(s):</b> " + html.escape(", ".join(val_result["missing_metadata_cols"])) + "</li>")
    if val_result.get("missing_spv_cols"):
        missing_items.append("<li><b>Missing Mandatory SPV DSM Column(s):</b> " + html.escape(", ".join(val_result["missing_spv_cols"])) + "</li>")
    if exception_exc:
        missing_items.append("<li><b>Parsing Exception:</b> " + html.escape(str(exception_exc)) + "</li>")

    alert_html = f"""
<div style="font-family: Arial, sans-serif; font-size: 14px; color: #333; line-height: 1.5;">
  <h3 style="color: #d9534f; margin-bottom: 5px;">⚠️ DSM-NIS Booking Alert: Mandatory Column / Extraction Failure</h3>
  <p>An email matching the subject keyword <b>"DSM-NIS-BOOKING"</b> was received, but mandatory column(s) or table(s) could not be extracted from the email body.</p>
  
  <table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; margin: 15px 0; width: 100%; max-width: 650px;">
    <tr style="background-color: #f2f2f2;"><td style="width: 30%;"><b>Email Subject</b></td><td>{html.escape(subject)}</td></tr>
    <tr><td><b>Received At</b></td><td>{html.escape(str(received_at))}</td></tr>
    <tr style="background-color: #f2f2f2;"><td><b>Message ID</b></td><td>{html.escape(str(message_id))}</td></tr>
  </table>

  <h4 style="color: #c9302c; margin-bottom: 5px;">Extraction Issue Details:</h4>
  <ul style="margin-top: 5px; color: #333;">
    {''.join(missing_items)}
  </ul>

  <div style="background-color: #fcf8e3; border-left: 4px solid #f0ad4e; padding: 12px; margin-top: 15px;">
    <b>Action Required:</b> Please verify the HTML table structure in the email body and ensure all required column headers are present.
  </div>

  <br>
  <p>Thanks &amp; Regards,<br><b>AGEL DSM-NIS Booking Automation</b></p>
</div>
"""

    # 3. Send one alert, while allowing extraction itself to retry every scan.
    if extraction_alert_already_sent:
        print("  [ALERT SKIP] Extraction alert was already sent for this email; retry remains enabled.")
    else:
        try:
            recipients = db.get_nis_completion_recipients()
            if recipients:
                spv_names = []
                if parsed_tables and isinstance(parsed_tables.get("spv_dsm_table"), list):
                    for r in parsed_tables["spv_dsm_table"]:
                        if isinstance(r, dict):
                            s_name = str(r.get("spv_name", "")).strip()
                            if s_name and s_name.upper() != "TOTAL" and s_name not in spv_names:
                                spv_names.append(s_name)
                spv_str = ", ".join(spv_names) if spv_names else ""
                if spv_str:
                    alert_subject = f"[ALERT] Column Extraction Failed | SPV: {spv_str}"
                else:
                    alert_subject = "[ALERT] Column Extraction Failed"

                # Remove any DSM-NIS-BOOKING keyword from the alert subject to prevent triggering loops
                alert_subject = re.sub(r"DSM-NIS-BOOKING", "", alert_subject, flags=re.IGNORECASE).strip()
                alert_subject = re.sub(r"\s+", " ", alert_subject)

                send_html_email(alert_subject, alert_html, recipients)
                print(f"  [ALERT SENT] Extraction failure notification sent to: {', '.join(recipients)}")
            else:
                print("  [WARN] No completion/notification recipients configured to receive extraction error alert.")
        except Exception as mail_exc:
            print(f"  [ERROR] Failed to send extraction alert email: {mail_exc}")


    # 4. Preserve diagnostics as extraction_failed. db.is_email_registered()
    # treats this status as retryable, so the next mailbox scan extracts again.
    try:
        run_id = db.register_email(
            message_id=message_id,
            subject=subject,
            received_at=str(received_at or datetime.now().isoformat()),
            body_html=msg.get("body", {}).get("content", ""),
        )
        db._update_run(run_id, {"status": "extraction_failed"})
    except Exception as db_exc:
        print(f"  [WARN] Could not register failed email in DB: {db_exc}")


def _process_one_message_batch(messages, headers, base_url):
    skipped = 0
    checklist_seen = 0

    for msg in messages:
        message_id, subject = msg["id"], msg.get("subject", "")

        checklist_status = _handle_checklist_email(msg, headers, base_url)
        if checklist_status != CHECKLIST_NOT_MATCHED:
            checklist_seen += 1
            if checklist_status == CHECKLIST_PROCESSED:
                return {
                    "did_work": True,
                    "kind": "checklist",
                    "processed": 0,
                    "checklist_processed": 1,
                    "skipped": skipped,
                    "checklist_seen": checklist_seen,
                }
            skipped += 1
            continue

        dc_status = _handle_debit_credit_reply(msg, headers, base_url)
        if dc_status != DC_NOT_MATCHED:
            checklist_seen += 1
            if dc_status == DC_PROCESSED:
                return {
                    "did_work": True,
                    "kind": "debit_credit_reply",
                    "processed": 0,
                    "checklist_processed": 1,
                    "skipped": skipped,
                    "checklist_seen": checklist_seen,
                }
            skipped += 1
            continue

        # ── HOD / FC Approval email detection ────────────────────────────────
        hod_status = _handle_hod_approval_email(msg, headers, base_url)
        if hod_status != APPROVAL_NOT_MATCHED:
            checklist_seen += 1
            if hod_status == APPROVAL_PROCESSED:
                return {
                    "did_work": True,
                    "kind": "hod_approval",
                    "processed": 0,
                    "checklist_processed": 1,
                    "skipped": skipped,
                    "checklist_seen": checklist_seen,
                }
            skipped += 1
            continue

        fc_status = _handle_fc_approval_email(msg, headers, base_url)
        if fc_status != APPROVAL_NOT_MATCHED:
            checklist_seen += 1
            if fc_status == APPROVAL_PROCESSED:
                return {
                    "did_work": True,
                    "kind": "fc_approval",
                    "processed": 0,
                    "checklist_processed": 1,
                    "skipped": skipped,
                    "checklist_seen": checklist_seen,
                }
            skipped += 1
            continue

        if not is_dsm_email(subject):
            skipped += 1
            continue

        if db.is_email_registered(message_id):
            skipped += 1
            continue

        print(f"  [NEW] Processing: {subject[:60]}")
        email_result = {}
        try:
            email_result = process_email(msg, headers, base_url)
        except Exception as exc:
            _cleanup_tmp(email_result.get("attachment_dir", ""))
            _cleanup_tmp(_message_tmp_dir(message_id))
            _handle_extraction_error(msg, exception_exc=exc)
            skipped += 1
            continue

        val_result = validate_mandatory_columns(email_result.get("parsed_tables", {}))
        if not val_result.get("is_valid", False):
            _cleanup_tmp(email_result.get("attachment_dir", ""))
            _handle_extraction_error(msg, email_result=email_result, val_result=val_result)
            skipped += 1
            continue

        run_id = db.register_email(
            message_id=message_id,
            subject=subject,
            received_at=email_result["received_at"],
            body_html=email_result["body_html"],
        )
        try:
            run_pipeline(run_id, email_result=email_result)
            return {
                "did_work": True,
                "kind": "booking",
                "processed": 1,
                "checklist_processed": 0,
                "skipped": skipped,
                "checklist_seen": checklist_seen,
            }
        except Exception:
            return {
                "did_work": True,
                "kind": "booking_failed",
                "processed": 0,
                "checklist_processed": 0,
                "skipped": skipped,
                "checklist_seen": checklist_seen,
            }

    return {
        "did_work": False,
        "kind": "",
        "processed": 0,
        "checklist_processed": 0,
        "skipped": skipped,
        "checklist_seen": checklist_seen,
    }


def scan_and_dispatch(lookback_days=DEFAULT_LOOKBACK_DAYS):
    total_processed = total_checklist_processed = total_skipped = scans = 0

    while True:
        scans += 1
        print(f"\n[{_ts()}] Scanning last {lookback_days} day(s) of email …")
        try:
            messages, headers, base_url = fetch_emails_last_n_days(days=lookback_days)
        except Exception as exc:
            print(f"[ERROR] Could not fetch emails: {exc}")
            break

        if not messages:
            print("[INFO] No emails found.")
            break

        print(f"[INFO] {len(messages)} email(s) to evaluate.\n")
        result = _process_one_message_batch(messages, headers, base_url)
        total_processed += result["processed"]
        total_checklist_processed += result["checklist_processed"]
        total_skipped += result["skipped"]

        if not result["did_work"]:
            break

        print(f"  [RESCAN] Completed one {result['kind']} action — checking mailbox again.")

    print(
        f"\n[SCAN COMPLETE] scans={scans}  processed={total_processed}  "
        f"checklist_processed={total_checklist_processed}  skipped={total_skipped}\n"
    )
    return bool(total_processed or total_checklist_processed)


# ── Run modes ────────────────────────────────────────────────────────────────
def _scan_and_prepare(lookback_days=DEFAULT_LOOKBACK_DAYS):
    """Fetch the mailbox and PREPARE every actionable email (no booking).

    Each new DSM email is registered and run through ``_prepare_run`` so its
    booking rows are seeded into the global pool. Checklist emails are handled
    inline. Returns the number of emails prepared + checklists processed, so the
    caller knows whether new work appeared.
    """
    print(f"\n[{_ts()}] Scanning last {lookback_days} day(s) of email …")
    try:
        messages, headers, base_url = fetch_emails_last_n_days(days=lookback_days)
    except Exception as exc:
        print(f"[ERROR] Could not fetch emails: {exc}")
        return 0

    if not messages:
        print("[INFO] No emails found.")
        return 0

    print(f"[INFO] {len(messages)} email(s) to evaluate.\n")
    prepared = 0
    checklist_processed = 0

    for msg in messages:
        message_id, subject = msg["id"], msg.get("subject", "")

        checklist_status = _handle_checklist_email(msg, headers, base_url)
        if checklist_status != CHECKLIST_NOT_MATCHED:
            if checklist_status == CHECKLIST_PROCESSED:
                checklist_processed += 1
            continue

        dc_status = _handle_debit_credit_reply(msg, headers, base_url)
        if dc_status != DC_NOT_MATCHED:
            if dc_status == DC_PROCESSED:
                checklist_processed += 1
            continue

        # ── HOD / FC Approval email detection ────────────────────────────────
        hod_status = _handle_hod_approval_email(msg, headers, base_url)
        if hod_status != APPROVAL_NOT_MATCHED:
            if hod_status == APPROVAL_PROCESSED:
                checklist_processed += 1
            continue

        fc_status = _handle_fc_approval_email(msg, headers, base_url)
        if fc_status != APPROVAL_NOT_MATCHED:
            if fc_status == APPROVAL_PROCESSED:
                checklist_processed += 1
            continue

        if not is_dsm_email(subject):
            continue
        if db.is_email_registered(message_id):
            continue

        print(f"  [NEW] Preparing: {subject[:60]}")
        email_result = {}
        try:
            email_result = process_email(msg, headers, base_url)
        except Exception as exc:
            _cleanup_tmp(email_result.get("attachment_dir", ""))
            _cleanup_tmp(_message_tmp_dir(message_id))
            _handle_extraction_error(msg, exception_exc=exc)
            continue

        val_result = validate_mandatory_columns(email_result.get("parsed_tables", {}))
        if not val_result.get("is_valid", False):
            _cleanup_tmp(email_result.get("attachment_dir", ""))
            _handle_extraction_error(msg, email_result=email_result, val_result=val_result)
            continue

        run_id = db.register_email(
            message_id=message_id,
            subject=subject,
            received_at=email_result["received_at"],
            body_html=email_result["body_html"],
        )
        try:
            if _prepare_run(run_id, email_result=email_result):
                prepared += 1
        except Exception as exc:
            print(f"  [WARN] Prepare failed for {subject[:50]}: {exc}")

    print(
        f"\n[PREPARE COMPLETE] prepared={prepared}  "
        f"checklist_processed={checklist_processed}\n"
    )
    return prepared + checklist_processed


def _prepare_incomplete_runs(skip):
    """Seed booking rows for any incomplete run that hasn't been seeded yet
    (e.g. a run that crashed before its rows were created). Runs that already
    have booking rows are left for the parallel booking driver.
    """
    for r in db.get_incomplete_runs():
        run_id = str(r.get("id") or "")
        if not run_id or run_id in skip:
            continue
        if db.get_booking_rows(run_id):
            continue  # already seeded → booking driver will pick the rows up
        print(f"[RESUME] Seeding booking rows for incomplete run: {run_id}")
        try:
            if not _prepare_run(run_id):
                skip.add(run_id)
        except Exception:
            skip.add(run_id)


def install_dependencies():
    req_file = "requirements.txt"
    if not os.path.exists(req_file):
        parent_req = os.path.join("..", "requirements.txt")
        if os.path.exists(parent_req):
            req_file = parent_req
        else:
            print("No requirements.txt found.")
            return

    print(f"Installing dependencies from {req_file}...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", req_file])
        print("Dependencies installed successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Failed to install dependencies: {e}")
        sys.exit(1)

def run_immediate(lookback_days):
    install_dependencies()
    parallel = max(1, int(getattr(config, "NIS_PARALLEL_COUNT", 1)))
    # Exit immediately once a scan finds no new emails / no pending work.
    # (Set NIS_MAX_IDLE_SCANS > 1 in config to re-enable the old "keep polling"
    #  behaviour.)
    max_idle_scans = max(1, int(getattr(config, "NIS_MAX_IDLE_SCANS", 1)))
    print(f"\n{'═'*65}\n  NIS BOOKING — ONE-SHOT RUN  |  lookback={lookback_days}d"
          f"  |  parallel={parallel}  |  max_idle_scans={max_idle_scans}\n  {_ts()}\n{'═'*65}\n")
    db.init_db()
    send_ready_debit_credit_emails()

    prepare_failed = set()
    idle_scans = 0  # counts consecutive scans with no new work
    try:
        while True:
            # ── 1) PREPARE: seed booking rows from new emails + crashed runs ──
            # (sequential, no browser). This fills the global pool of pending
            # booking rows that the parallel booker draws from.
            prepared = _scan_and_prepare(lookback_days)
            _prepare_incomplete_runs(prepare_failed)

            # ── 2) BOOK: process pending rows across ALL emails, N at a time ──
            # A free worker slot is filled with a row from another email, so a
            # 2-row email plus rows from other emails still keep N workers busy.
            booked = _book_pending_rows_parallel(parallel)

            # ── 3) DEBIT/CREDIT EMAIL: once a run's NIS (+ve) bookings are done,
            # release its 5-row debit/credit request email to finance (gated).
            # NOTE: nis_booking NEVER runs SAP checklist automation. The SAP
            # side (— including finance-reply PDF checklists —) is handled
            # exclusively by the independent sap_automation service.
            send_ready_debit_credit_emails()

            # ── 4) HOD/FC APPROVAL + PDF MERGE + STAGE EMAILS ────────────────
            # Process approval stages: send HOD/FC stage notifications,
            # merge PDFs, and send final completion emails.
            process_approval_emails_and_stages()

            # Nothing new prepared and no rows left to book → count as idle scan.
            if not prepared and booked == 0:
                idle_scans += 1
                if idle_scans >= max_idle_scans:
                    print(f"\n[EXIT] No new booking emails / pending work found. Exiting immediately.")
                    break
                print(f"  [IDLE] No work found (scan {idle_scans}/{max_idle_scans}). "
                      f"Will exit after {max_idle_scans - idle_scans} more idle scan(s).")
                # Wait briefly before next scan to avoid hammering the mailbox
                time.sleep(30)
            else:
                # Reset idle counter when work was found
                idle_scans = 0
    finally:
        _cleanup_tmp(TEMP_BOOKING_ROOT)

    send_ready_debit_credit_emails()
    process_approval_emails_and_stages()
    print(f"\n{'═'*65}\n  ALL DONE — NIS pipeline complete.\n  {_ts()}\n{'═'*65}\n")
    sys.exit(0)


def nis_main():
    parser = argparse.ArgumentParser(description="NIS Booking Pipeline")
    parser.add_argument("--lookback", type=int, default=DEFAULT_LOOKBACK_DAYS)
    args = parser.parse_args()
    run_immediate(args.lookback)

if __name__ == "__main__":
    nis_main()
