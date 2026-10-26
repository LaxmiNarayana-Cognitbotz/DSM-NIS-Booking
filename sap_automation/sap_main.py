"""
sap_main.py  (sap_automation)
=============================
Standalone entry point for SAP automation.

This service polls the shared Sheet DB for pending SAP automation steps,
processes them sequentially (login → fill form → save → capture checklist → PDF),
and writes results back to the same Sheet DB and Excel tracker.

Usage
─────
  python sap_main.py                       # one-shot: process all pending, then exit
  python sap_main.py --poll --interval 60  # poll every 60s
"""

import argparse
import html
import json
import os
import re
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path

import db_manager as db
from sap_automation import run_sap_automation, cleanup_sap_windows
from read_vendor_master_data import read_vendor_data
from add_new_row_data_nis import add_incremental_week_row, add_sap_checklist_reply_row, update_sap_checklist_on_existing_row
import sharepoint_upload as sp
import sap_config as config
import log_manager as lm
import os
import subprocess
import sys

# ── Compatibility shim: ensure new attributes exist on db_manager ────────────
# Prevents crashes when the code is newer than the db_manager.py on the VM.
if not hasattr(db, "STATUS_NIS_DONE"):
    db.STATUS_NIS_DONE = "nis_done"
if not hasattr(db, "mark_booking_row_nis_done"):
    db.mark_booking_row_nis_done = lambda row_id: db.update_booking_row(
        row_id, {"status": "nis_done"}
    )
if not hasattr(db, "finalize_booking_row_if_ready"):
    db.finalize_booking_row_if_ready = lambda run_id, row_index, spv_name: False
if not hasattr(db, "are_all_sap_steps_done_for_row"):
    db.are_all_sap_steps_done_for_row = lambda run_id, row_index, spv_name: True
# ──────────────────────────────────────────────────────────────────────────────
sys.stdout.reconfigure(encoding='utf-8')
import requests


# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────
MAX_STEP_RETRIES = getattr(config, "MAX_STEP_RETRIES", 3)

SPV_AMOUNT_FIELDS = [
    ("total_dsm_charges_payable", "Total DSM Charges Payable"),
    ("drawl_charges_payable",     "Drawl Charges Payable"),
    ("revenue_diff",              "Revenue Diff"),
    ("revenue_loss",              "Revenue Loss"),
]

# Local temp root for SAP PDF files
TEMP_BOOKING_ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp-book")
POST_SAP_DONE_MARKER_DIR = os.path.join(TEMP_BOOKING_ROOT, "_completed_sap_tasks")


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────
def _ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _sanitize_folder_name(name: str) -> str:
    import re as _re
    name = _re.sub(r'[\\/:*?"<>|]', "_", name)
    name = _re.sub(r"\s+", " ", name).strip()
    return name


def _spv_tmp_dir(spv_name: str) -> str:
    safe = _sanitize_folder_name(spv_name) if spv_name else "unknown_spv"
    path = os.path.join(TEMP_BOOKING_ROOT, safe)
    os.makedirs(path, exist_ok=True)
    return path


def _cleanup_tmp(path: str):
    import shutil
    try:
        if path and os.path.isdir(path):
            shutil.rmtree(path, ignore_errors=True)
            print(f"  [TMP] Cleaned up local temp: {path}")
    except Exception as exc:
        print(f"  [TMP] Cleanup warning: {exc}")


def normalize_vendor_data(raw: dict) -> dict:
    out = {}
    for k, v in raw.items():
        key = k.strip().upper()
        if key in ("VENDOR", "VENDOR CODE"):
            out["vendor_code"]  = str(v)
        elif key in ("COMPANY", "COMPANY CODE"):
            out["company_code"] = str(v)
        elif key == "COST CENTER":
            out["cost_center"]  = str(v)
        elif key == "PLANT":
            out["plant"]        = str(v)
        elif key == "BANK KEY":
            out["bank_key"]     = str(v)
        elif key == "GL":
            out["gl_account"]   = str(v)
        elif key == "PURPOSE":
            out["purpose"]      = str(v)
    return out


def safe_float(val):
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    cleaned = str(val).replace(",", "").strip()
    try:
        return float(cleaned)
    except (ValueError, TypeError):
        return 0.0


def _post_sap_call(label: str, fn, *args, attempts: int = 3, **kwargs):
    last_exc = None
    for attempt in range(1, attempts + 1):
        try:
            return fn(*args, **kwargs)
        except Exception as exc:
            last_exc = exc
            print(f"  ⚠️  Post-SAP {label} failed ({attempt}/{attempts}): {str(exc)[:180]}")
            if attempt < attempts:
                time.sleep(5 * attempt)
    raise last_exc


def _marker_path(run_id: str, step_name: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", f"{run_id}_{step_name}")[:180]
    return os.path.join(POST_SAP_DONE_MARKER_DIR, f"{safe}.json")


def _write_post_sap_done_marker(run_id: str, step_name: str, detail: str):
    try:
        os.makedirs(POST_SAP_DONE_MARKER_DIR, exist_ok=True)
        with open(_marker_path(run_id, step_name), "w", encoding="utf-8") as fh:
            json.dump({"run_id": run_id, "step_name": step_name, "detail": detail}, fh)
        print("  [RECOVERY] Local SAP-completed marker written for later Sheet DB recovery.")
    except Exception as exc:
        print(f"  ⚠️  Could not write SAP-completed marker: {exc}")


def _consume_post_sap_done_marker(run_id: str, step_name: str) -> bool:
    path = _marker_path(run_id, step_name)
    if not os.path.isfile(path):
        return False
    try:
        with open(path, "r", encoding="utf-8") as fh:
            marker = json.load(fh)
        detail = str(marker.get("detail") or "SAP Save previously completed; recovered from local marker")
        _post_sap_call("recover SAP step done marker", db.step_done, run_id, step_name, detail=detail)
        os.remove(path)
        print(f"  [RECOVERY] Marked {step_name} done from local SAP-completed marker; skipping SAP rerun.")
        return True
    except Exception as exc:
        print(f"  ⚠️  Could not consume SAP-completed marker for {step_name}: {str(exc)[:180]}")
        return True


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


def _run_context(run_id: str):
    run = db.get_run_by_id(run_id) or {}
    metadata = _json_dict(run.get("metadata_json"))
    spv_rows = _json_list(run.get("spv_dsm_json"))
    spv_name = str(run.get("spv_name") or "").strip()
    if not spv_name:
        for row in spv_rows:
            spv_name = str((row or {}).get("spv_name") or "").strip()
            if spv_name:
                break
    return run, metadata, spv_name


def _graph_token():
    try:
        import msal
        app = msal.ConfidentialClientApplication(
            config.GRAPH_CLIENT_ID,
            authority=config.GRAPH_AUTHORITY,
            client_credential=config.GRAPH_CLIENT_SECRET,
        )
        result = app.acquire_token_for_client(scopes=config.GRAPH_SCOPE)
        return result.get("access_token") or ""
    except Exception as exc:
        print(f"  [EMAIL] Could not acquire Graph token: {exc}")
        return ""


def _send_mail_to_completion_recipients(subject: str, body: str, mark_sent=None) -> bool:
    recipients = db.get_nis_completion_recipients()
    if not recipients:
        print("  [EMAIL] Notification not sent: no recipients in nis_completion_recipients.emails")
        return False
    token = _graph_token()
    if not token:
        return False
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body},
            "toRecipients": [{"emailAddress": {"address": addr}} for addr in recipients],
        },
        "saveToSentItems": "true",
    }
    try:
        resp = requests.post(
            f"https://graph.microsoft.com/v1.0/users/{config.USER_EMAIL}/sendMail",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json=payload,
            timeout=60,
        )
        if resp.status_code in (200, 202):
            if mark_sent:
                mark_sent()
            print(f"  [EMAIL] Notification sent to {', '.join(recipients)}")
            return True
        print(f"  [EMAIL] Notification failed: HTTP {resp.status_code} {resp.text[:300]}")
    except Exception as exc:
        print(f"  [EMAIL] Could not send notification: {exc}")
    return False


def _send_failure_email(run_id: str, step_name: str, error_message: str) -> bool:
    if db.is_failure_email_sent(run_id):
        return False
    run, metadata, spv_name = _run_context(run_id)
    from_date = str(metadata.get("from_date") or "").strip()
    to_date = str(metadata.get("to_date") or "").strip()
    subject = f"[DSM-NIS Automation FAILED] {spv_name or 'unknown'} | {from_date} to {to_date}"

    def esc(value):
        return html.escape(str(value or ""))

    body = f"""
<h3>Automation Failure</h3>
<p>The automation could not be completed.</p>
<table border="1" cellspacing="0" cellpadding="6">
  <tr><td><b>Run ID</b></td><td>{esc(run_id)}</td></tr>
  <tr><td><b>SPV</b></td><td>{esc(spv_name)}</td></tr>
  <tr><td><b>Period</b></td><td>{esc(from_date)} to {esc(to_date)}</td></tr>
  <tr><td><b>Step</b></td><td>{esc(step_name)}</td></tr>
  <tr><td><b>Subject</b></td><td>{esc(run.get("subject"))}</td></tr>
</table>
<pre>{esc(str(error_message)[:1000])}</pre>
"""
    return _send_mail_to_completion_recipients(
        subject,
        body,
        mark_sent=lambda: db.mark_failure_email_sent(run_id),
    )


def _is_run_ready_for_completion_email(run_id: str) -> bool:
    run = db.get_run_by_id(run_id)
    if not run or db.is_nis_completion_email_sent(run_id):
        return False
    nis_step = db.get_step(run_id, db.STEP_NIS_BOOKING)
    if not nis_step or nis_step.get("status") != db.STATUS_DONE:
        return False
    return db.are_all_sap_steps_complete(run_id)


def _send_nis_completion_email(run_id: str) -> bool:
    # DISABLED: Completion emails are now gated behind HOD/FC approval stages.
    # The nis_booking service (main.py) sends completion emails only after ALL
    # stages are done (NIS + SAP + HOD + FC + PDF merge). This SAP service
    # should NOT send premature completion emails after SAP done.
    return False
    run, metadata, spv_name = _run_context(run_id)
    from_date = str(metadata.get("from_date") or "").strip()
    to_date = str(metadata.get("to_date") or "").strip()
    week_no = str(metadata.get("week_no") or "").strip()
    nis_number = str(run.get("nis_checklist_id") or "").strip()
    sap_number = str(run.get("checklist_value") or "").strip()
    booking_folder = str(run.get("booking_folder_path") or "").strip()
    sap_number_note = (
        sap_number
        if sap_number
        else "Not captured. SAP Save was submitted, but the checklist number could not be read from SAP."
    )

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
    return _send_mail_to_completion_recipients(
        subject,
        body,
        mark_sent=lambda: db.mark_nis_completion_email_sent(run_id),
    )


def send_ready_nis_completion_emails():
    sent = 0
    for run in db.get_runs_pending_nis_completion_email():
        try:
            if _send_nis_completion_email(run.get("id") or ""):
                sent += 1
        except Exception as exc:
            print(f"  [EMAIL] Completion notification check failed for run_id={run.get('id')}: {exc}")
    return sent


# ─────────────────────────────────────────────────────────────────────────────
# Core: find pending SAP tasks
# ─────────────────────────────────────────────────────────────────────────────
def _has_pending_sap_work(run: dict) -> bool:
    """
    Determine whether a run has SAP automation work that still needs doing.

    Logic:
      1. Parse negative_fields_json — if empty / no data → no SAP work.
      2. For each negative field, check if a corresponding sap_automation_*
         step exists AND is already done/skipped.  If even one is NOT
         done/skipped → SAP work is pending.
      3. Additionally scan run_steps for any step whose name contains "sap"
         with a "pending" or "running" status — that also counts.
    """
    # ── 1. Check negative_fields_json ────────────────────────────────────
    neg_raw = run.get("negative_fields_json") or "{}"
    try:
        negative = json.loads(neg_raw) if isinstance(neg_raw, str) else neg_raw
    except (json.JSONDecodeError, TypeError):
        negative = {}

    if not negative or not isinstance(negative, dict):
        return False          # nothing negative → no SAP work

    run_id = str(run.get("id") or "")
    if not run_id:
        return False

    steps = db.get_steps(run_id)
    step_map = {str(s.get("step_name") or ""): s for s in steps}

    # ── 2. Per-field check ───────────────────────────────────────────────
    for field_key, amount in negative.items():
        if safe_float(amount) <= 0:
            continue                          # zero / invalid → skip
        sap_step_name = db.step_sap(field_key)
        step = step_map.get(sap_step_name)
        if step is None:
            # Step was never created → SAP work is pending
            return True
        status = str(step.get("status") or "")
        if status not in (db.STATUS_DONE, db.STATUS_SKIPPED):
            # Step exists but not finished → pending
            return True

    # ── 3. Catch-all: any run_step with "sap" in name still pending ─────
    for s in steps:
        sname = str(s.get("step_name") or "").lower()
        sstatus = str(s.get("status") or "")
        if "sap" in sname and sstatus in (db.STATUS_PENDING, db.STATUS_RUNNING):
            return True

    return False


def _ensure_sap_steps_for_run(run: dict) -> list:
    """
    Inspect negative_fields_json for *run* and guarantee that a
    ``sap_automation_<field>`` step row exists in ``run_steps`` for every
    negative field with a non-zero amount.  Steps that are missing are
    created in 'pending' status via ``db.ensure_dynamic_step``.

    IMPORTANT: If per-row SAP steps already exist for this run (created by
    the NIS booking's _process_booking_row), do NOT create run-level steps —
    those per-row steps are the authoritative source. Creating run-level
    duplicates would cause the same amount to be processed twice.

    Returns a list of (step_name, field_key) tuples that are still pending
    (i.e. not done / skipped).
    """
    run_id = str(run.get("id") or "")
    if not run_id:
        return []

    # ── Check if per-row SAP steps already exist for this run ────────────────
    # If any step matching "sap_automation_*_row_*" exists, the NIS booking
    # service has already created the authoritative per-row steps. Do NOT
    # create additional run-level steps (they would be duplicates).
    existing_steps = db.get_steps(run_id)
    has_per_row_sap_steps = any(
        "_row_" in str(s.get("step_name") or "")
        for s in existing_steps
        if str(s.get("step_name") or "").startswith("sap_automation_")
    )
    if has_per_row_sap_steps:
        return []  # per-row steps are the source of truth — no run-level needed

    neg_raw = run.get("negative_fields_json") or "{}"
    try:
        negative = json.loads(neg_raw) if isinstance(neg_raw, str) else neg_raw
    except (json.JSONDecodeError, TypeError):
        negative = {}

    if not negative or not isinstance(negative, dict):
        return []

    pending_pairs = []
    for field_key, amount in negative.items():
        if safe_float(amount) <= 0:
            continue
        sap_step_name = db.step_sap(field_key)          # "sap_automation_<field>"
        db.ensure_dynamic_step(run_id, sap_step_name)   # no-op if already exists
        step = db.get_step(run_id, sap_step_name)
        status = str(step.get("status") or "") if step else ""
        if status not in (db.STATUS_DONE, db.STATUS_SKIPPED):
            pending_pairs.append((sap_step_name, field_key))

    return pending_pairs


def _parse_sap_step_name(step_name: str):
    """Decode a SAP step name into (field_key, row_index_or_None).

    Per-row steps (5-row table system) are named:
        sap_automation_<field>_row_<N>_<spv_safe_name>
    e.g.  sap_automation_revenue_diff_row_3_AWEK01L_50_MW

    Legacy / single-row steps are named:
        sap_automation_<field>
    """
    base = step_name
    if base.startswith("sap_automation_"):
        base = base[len("sap_automation_"):]

    # Look for the _row_<digits>_ sentinel
    import re as _re
    m = _re.search(r"_row_(\d+)_", base)
    if m:
        field_key = base[:m.start()]      # everything before _row_N_
        row_index = int(m.group(1))
        return field_key, row_index
    return base, None


def get_pending_sap_tasks() -> list:
    """
    Scan the Sheet DB for ALL runs that have pending SAP automation work.

    PRIMARY path - scan run_steps directly:
      db.get_all_pending_sap_steps() returns every sap_automation_* step row
      whose status is not done/skipped, regardless of the parent run status
      in email_runs.

      WHY: nis_booking calls mark_run_done() on the parent run as soon as its
      own pipeline finishes - even though SAP steps are still pending.
      This makes get_incomplete_runs() completely blind to those tasks.

    SECONDARY path - negative_fields_json sweep:
      For runs still marked incomplete where sap_automation_* step rows were
      never created, we auto-create them here.

    Returns a list of dicts:
      [
        {
          "run_id":      <str>,
          "step_name":   "sap_automation_<field>[_row_N_<spv>]",
          "field_key":   "<field>",
          "run":         { ... email_runs row ... },
          "step":        { ... run_steps row ... },
          "booking_row": { ... booking_rows row ... } or None,
        },
        ...
      ]
    """
    tasks = []
    seen_keys = set()  # (run_id, step_name) dedup

    # PRIMARY: scan run_steps directly (finds tasks even when run status=done)
    pending_steps = db.get_all_pending_sap_steps()
    print(f"  [DETECT] run_steps scan -> {len(pending_steps)} pending SAP step(s) found")

    for step in pending_steps:
        run_id    = str(step.get("run_id") or "")
        step_name = str(step.get("step_name") or "")
        if not run_id or not step_name:
            continue
        key = (run_id, step_name)
        if key in seen_keys:
            continue
        seen_keys.add(key)

        # Decode field_key and optional per-row index from the step name
        field_key, row_index = _parse_sap_step_name(step_name)

        run = db.get_run_by_id(run_id) or {}

        # Resolve the booking row for per-row step names (5-row table system)
        booking_row = None
        if row_index is not None:
            all_rows = db.get_booking_rows(run_id)
            for br in all_rows:
                if int(br.get("row_index") or 0) == row_index:
                    booking_row = br
                    break

        print(f"    -> run_id={run_id}  step={step_name}  "
              f"field={field_key}  row={row_index}  "
              f"booking_row={'yes' if booking_row else 'no'}  "
              f"status={step.get('status')}  "
              f"subject={str(run.get('subject', ''))[:50]}")
        tasks.append({
            "run_id":      run_id,
            "step_name":   step_name,
            "field_key":   field_key,
            "run":         run,
            "step":        step,
            "booking_row": booking_row,
        })

    # SECONDARY: sweep incomplete runs for orphaned negative fields
    for run in db.get_incomplete_runs():
        run_id = str(run.get("id") or "")
        if not run_id:
            continue
        for step_name, field_key in _ensure_sap_steps_for_run(run):
            key = (run_id, step_name)
            if key in seen_keys:
                continue
            seen_keys.add(key)
            step = db.get_step(run_id, step_name)
            print(f"  [DETECT-SECONDARY] run_id={run_id}  step={step_name} "
                  f"(auto-created from negative_fields_json)")
            tasks.append({
                "run_id":      run_id,
                "step_name":   step_name,
                "field_key":   field_key,
                "run":         run,
                "step":        step or {},
                "booking_row": None,
            })

    return tasks


# ─────────────────────────────────────────────────────────────────────────────
# Core: process a single SAP task
# ─────────────────────────────────────────────────────────────────────────────
def process_sap_task(task: dict):
    """
    Run SAP automation for one pending task.

    Steps:
      1. Load run data (metadata, negative_fields, vendor_data) from Sheet DB
      2. Read vendor master data if not already loaded
      3. Run SAP GUI automation (login → form → save → PDF → capture checklist)
      4. Upload SAP checklist PDF to SharePoint booking folder
      5. Update Sheet DB with checklist number and mark step done
      6. Update Excel tracker (DSM-NIS-OUTPUT) with checklist info
    """
    run_id    = task["run_id"]
    step_name = task["step_name"]
    field_key = task["field_key"]
    run       = task["run"]

    label = field_key
    for fk, lbl in SPV_AMOUNT_FIELDS:
        if fk == field_key:
            label = lbl
            break

    print(f"  SAP TASK  run_id={run_id}  step={step_name}")
    print(f"  Subject: {str(run.get('subject', ''))[:60]}")
    print(f"{'─' * 65}")

    if _consume_post_sap_done_marker(run_id, step_name):
        _send_nis_completion_email(run_id)
        return

    # ── Load run data ─────────────────────────────────────────────────────
    metadata_row   = json.loads(run.get("metadata_json")        or "{}")
    spv_dsm_rows   = json.loads(run.get("spv_dsm_json")         or "[]")
    negative       = json.loads(run.get("negative_fields_json")  or "{}")
    vendor_json    = run.get("vendor_data_json") or "{}"
    vendor_data_raw = json.loads(vendor_json) if vendor_json else {}
    spv_name       = str(run.get("spv_name", "")).strip()
    booking_folder = str(run.get("booking_folder_path", "")).strip()

    # The run-level vendor_data_json can be either:
    #   a) A flat dict: {"vendor_code": "...", "company_code": "...", ...}
    #   b) A nested map: {"SPV-Name": {"vendor_code": "...", ...}, ...}
    # Detect which format it is and resolve to a flat vendor dict.
    _vendor_keys = {"vendor_code", "company_code", "cost_center", "plant", "bank_key", "gl_account"}
    if vendor_data_raw and any(k in vendor_data_raw for k in _vendor_keys):
        # Already flat
        vendor_data = vendor_data_raw
    elif vendor_data_raw and spv_name:
        # Nested map — look up by SPV name (exact or partial match)
        vendor_data = vendor_data_raw.get(spv_name, {})
        if not vendor_data:
            # Try case-insensitive / partial key match
            target_upper = spv_name.upper()
            for key, val in vendor_data_raw.items():
                if isinstance(val, dict) and key.upper() == target_upper:
                    vendor_data = val
                    break
            if not vendor_data:
                # Try matching by the SPV code prefix
                import re as _re_vendor
                target_code = _re_vendor.split(r"[\s\-_]+", spv_name)[0].upper()
                for key, val in vendor_data_raw.items():
                    if isinstance(val, dict) and key.upper().startswith(target_code):
                        vendor_data = val
                        print(f"  [VENDOR] Resolved from nested map by code prefix: "
                              f"'{spv_name}' → key='{key}'")
                        break
    else:
        vendor_data = vendor_data_raw if isinstance(vendor_data_raw, dict) else {}

    # ── Per-row override (5-row parallel system) ─────────────────────────
    # When this task was created by the per-row NIS booking path, the task
    # dict carries the matching booking_rows record.  Its per-row negative,
    # vendor, and metadata values take precedence over the run-level ones
    # stored on email_runs (which may be empty or represent a DIFFERENT row).
    booking_row = task.get("booking_row")
    if booking_row:
        _row_neg  = _json_dict(booking_row.get("negative_fields_json"))
        _row_vend = _json_dict(booking_row.get("vendor_data_json"))
        _row_meta = _json_dict(booking_row.get("metadata_json"))
        _row_spv  = _json_dict(booking_row.get("spv_data_json"))
        _row_name = str(booking_row.get("spv_name") or "").strip()
        _row_fold = str(booking_row.get("booking_folder_path") or "").strip()
        if _row_neg:
            negative = _row_neg
        if _row_vend:
            vendor_data = _row_vend
        if _row_meta:
            metadata_row = _row_meta
        if _row_spv:
            spv_row_override = _row_spv
        else:
            spv_row_override = None
        if _row_name:
            spv_name = _row_name
        if _row_fold:
            booking_folder = _row_fold
        print(f"  [PER-ROW] Using booking_row data: spv={spv_name}  "
              f"negative_keys={list(negative.keys())}  "
              f"vendor={'yes' if vendor_data else 'no'}")
    else:
        spv_row_override = None
    # ─────────────────────────────────────────────────────────────────────

    # Find the SPV DSM row
    spv_row = {}
    if spv_row_override:
        spv_row = spv_row_override
    else:
        for r in spv_dsm_rows:
            if str(r.get("spv_name", "")).strip():
                spv_row = r
                break


    if not spv_name and spv_row:
        spv_name = str(spv_row.get("spv_name", "")).strip()

    # ── Get the absolute amount for this field ───────────────────────────
    abs_amount = safe_float(negative.get(field_key, 0))
    if abs_amount <= 0:
        # Try to compute from the SPV DSM row
        raw_val = safe_float(spv_row.get(field_key, 0))
        abs_amount = abs(raw_val) if raw_val < 0 else 0

    if abs_amount <= 0:
        print(f"  [SKIP] No negative amount found for {field_key} – marking as skipped")
        db.step_skip(run_id, step_name, f"No negative amount for {field_key}")
        return

    # ── Load vendor data if not already present ──────────────────────────
    if not vendor_data and spv_name:
        try:
            raw_vendor  = read_vendor_data(spv_name)
            vendor_data = normalize_vendor_data(raw_vendor)
            db.save_vendor_data(run_id, vendor_data)
            print(f"  [VENDOR] Loaded vendor data for {spv_name}: {vendor_data}")
        except Exception as exc:
            print(f"  ❌ Could not load vendor data for {spv_name}: {exc}")
            db.step_failed(run_id, step_name, f"Vendor data load failed: {exc}")
            _send_failure_email(run_id, step_name, f"Vendor data load failed: {exc}")
            return

    if not vendor_data:
        msg = f"No vendor data available for {spv_name}"
        print(f"  ❌ {msg}")
        db.step_failed(run_id, step_name, msg)
        _send_failure_email(run_id, step_name, msg)
        return

    # ── Final guard: ensure vendor fields are not empty ──────────────────
    company_code = vendor_data.get("company_code", "")
    vendor_code  = vendor_data.get("vendor_code",  "")
    cost_center  = vendor_data.get("cost_center",  "")
    plant        = vendor_data.get("plant",        "")

    if not company_code or not vendor_code:
        msg = (f"Vendor data resolved but critical fields are empty for '{spv_name}': "
               f"company_code='{company_code}' vendor_code='{vendor_code}' "
               f"cost_center='{cost_center}' plant='{plant}'. "
               f"Raw vendor_data={json.dumps(vendor_data)[:300]}")
        print(f"  ❌ {msg}")
        db.step_failed(run_id, step_name, msg)
        _send_failure_email(run_id, step_name, msg)
        return

    # ── Set up log context ───────────────────────────────────────────────
    from_date = str(metadata_row.get("from_date", "")).strip()
    to_date   = str(metadata_row.get("to_date", "")).strip()
    lm.set_context(spv_name, from_date, to_date)
    lm.log(f"SAP_TASK_START run_id={run_id} step={step_name} amount={abs_amount}")

    # ── SAP date + output dir ────────────────────────────────────────────
    deb_not_date = datetime.today().strftime("%d.%m.%Y")

    # SAP PDF output dir
    pdf_output_dir = _spv_tmp_dir(spv_name) if spv_name else TEMP_BOOKING_ROOT

    # ── Guard: check retry count ─────────────────────────────────────────
    retry_count = db.get_step_retry_count(run_id, step_name)
    if retry_count >= MAX_STEP_RETRIES:
        msg = (f"SAP [{label}] exceeded maximum {MAX_STEP_RETRIES} "
               f"retries – marking permanently failed")
        print(f"  ❌ {msg}")
        db.step_failed(run_id, step_name, msg)
        lm.log_error(msg)
        _send_failure_email(run_id, step_name, msg)
        return

    # ── Retry loop ────────────────────────────────────────────────────────
    while True:
        sap_save_submitted = False
        step_done_detail = ""
        current_attempt = db.get_step_retry_count(run_id, step_name) + 1
        db.step_start(
            run_id, step_name,
            detail=f"attempt={current_attempt}/{MAX_STEP_RETRIES}  "
                   f"amount={abs_amount}  date={deb_not_date}",
        )
        try:
            print(f"\n  ▶ SAP automation  [{label}]  amount={abs_amount}"
                  f"  (attempt {current_attempt}/{MAX_STEP_RETRIES})")
            try:
                cleanup_sap_windows()
            except Exception as cleanup_exc:
                print(f"  ⚠️  Pre-run SAP cleanup error (non-fatal): {cleanup_exc}")

            sap_result = run_sap_automation(
                company_code   = company_code,
                vendor_code    = vendor_code,
                amount         = abs_amount,
                cost_center    = cost_center,
                plant          = plant,
                deb_not_date   = deb_not_date,
                pdf_output_dir = pdf_output_dir,
            )
            sap_save_submitted = True

            # ── Parse result ─────────────────────────────────────────────
            sap_checklist_number = ""
            sap_pdf_full_path    = ""
            sap_pdf_name         = ""
            if isinstance(sap_result, tuple) and len(sap_result) >= 2:
                sap_checklist_number = str(sap_result[0] or "")
                sap_pdf_full_path    = str(sap_result[1] or "")
                sap_pdf_name         = str(Path(sap_pdf_full_path).name) if sap_pdf_full_path else ""
            elif isinstance(sap_result, str):
                sap_checklist_number = sap_result or ""

            # ── Persist checklist number ─────────────────────────────────
            if sap_checklist_number:
                _post_sap_call("save checklist number", db.save_checklist_value, run_id, sap_checklist_number)
                print(f"  [DB] SAP checklist number saved: {sap_checklist_number}")
                # Also persist on the per-row booking_row record
                if booking_row and booking_row.get("row_id"):
                    try:
                        db.save_booking_row_sap_checklist_number(
                            booking_row["row_id"], sap_checklist_number
                        )
                        print(f"  [DB] SAP checklist saved to booking_row #{booking_row['row_id']}")
                    except Exception as br_exc:
                        print(f"  ⚠️  Could not save SAP checklist to booking_row: {br_exc}")
            else:
                # SAP automation ran but checklist number was NOT captured.
                # This means the save likely went through but the screen-read
                # failed. Do NOT mark done — retry to capture the number.
                print(f"  ⚠️  SAP ran but checklist number NOT captured — will retry to read it.")
                try:
                    cleanup_sap_windows()
                except Exception:
                    pass
                new_count = db.increment_step_retry(run_id, step_name,
                    "SAP save submitted but checklist number not captured")
                if new_count >= MAX_STEP_RETRIES:
                    # Exhausted retries — mark done anyway with a note so the
                    # pipeline doesn't get permanently stuck.
                    print(f"  ⚠️  Max retries reached without capturing checklist number — marking done with note.")
                    step_done_detail = (
                        f"amount={abs_amount}  attempts={new_count}"
                        f"  checklist=NOT_CAPTURED"
                        f"  pdf={sap_pdf_name or 'n/a'}"
                    )
                    _post_sap_call("mark SAP step done (no checklist)",
                                   db.step_done, run_id, step_name,
                                   detail=step_done_detail)
                    # Still finalize the row
                    try:
                        _, row_index_parsed = _parse_sap_step_name(step_name)
                        if row_index_parsed is not None:
                            db.finalize_booking_row_if_ready(run_id, row_index_parsed, spv_name)
                        elif booking_row and booking_row.get("row_id"):
                            br_index = int(booking_row.get("row_index") or 0)
                            db.finalize_booking_row_if_ready(run_id, br_index, spv_name)
                    except Exception:
                        pass
                    if pdf_output_dir != TEMP_BOOKING_ROOT:
                        _cleanup_tmp(pdf_output_dir)
                    lm.clear_context()
                    return
                wait_sec = 10
                print(f"  🔁 Retrying in {wait_sec}s to capture checklist number "
                      f"({MAX_STEP_RETRIES - new_count} attempt(s) remaining) …")
                time.sleep(wait_sec)
                if pdf_output_dir != TEMP_BOOKING_ROOT:
                    _cleanup_tmp(pdf_output_dir)
                lm.clear_context()
                continue  # retry the while loop


            # ── Upload SAP checklist PDF to SharePoint ───────────────────
            if booking_folder and sap_pdf_full_path and os.path.isfile(sap_pdf_full_path):
                orig_basename = os.path.basename(sap_pdf_full_path)
                dst_basename = f"sap_checklist_{orig_basename}"
                sp_path = sp.upload_file_to_sp(sap_pdf_full_path, booking_folder, dst_basename)
                if sp_path:
                    sap_pdf_name = dst_basename
                    _post_sap_call("mark SAP PDF uploaded", db.mark_sap_checklist_pdf_uploaded, run_id, True)
                    _post_sap_call("save SAP PDF name", db.save_downloaded_pdf_name, run_id, "sap_checklist_pdf", sap_pdf_name)
                    print(f"  [BOOKING] Uploaded SAP checklist PDF → SP:{sp_path}")
                    try:
                        os.remove(sap_pdf_full_path)
                    except Exception:
                        pass

            # ── Mark step done ───────────────────────────────────────────
            step_done_detail = (
                f"amount={abs_amount}  attempts={current_attempt}"
                f"  checklist={sap_checklist_number or 'n/a'}"
                f"  pdf={sap_pdf_name or 'n/a'}"
            )
            _post_sap_call(
                "mark SAP step done",
                db.step_done,
                run_id, step_name,
                detail=step_done_detail,
            )
            print(f"  ✅ SAP done  [{label}]  checklist={sap_checklist_number or '(n/a)'}")

            lm.log(f"SAP_TASK_DONE run_id={run_id} step={step_name} "
                   f"checklist={sap_checklist_number}")

            # ── Update Excel tracker column K immediately ─────────────────
            # Write the SAP checklist number to column K right away (merged
            # with any existing numbers if multiple negatives produce multiple
            # checklists). Don't wait for all SAP steps to finish.
            if sap_checklist_number:
                try:
                    update_sap_checklist_on_existing_row(
                        metadata_row=metadata_row,
                        sheet_name=spv_name,
                        sap_checklist_number=sap_checklist_number,
                    )
                except Exception as post_exc:
                    print(f"  ⚠️  Tracker K-column update failed: {str(post_exc)[:180]}")

            # ── Finalize booking row if all SAP steps for this row are done ──
            try:
                _, row_index_parsed = _parse_sap_step_name(step_name)
                if row_index_parsed is not None:
                    finalized = db.finalize_booking_row_if_ready(
                        run_id, row_index_parsed, spv_name
                    )
                    if finalized:
                        print(f"  ✅ [ROW {row_index_parsed}] {spv_name} fully done "
                              f"(NIS + all SAP steps complete).")
                elif booking_row and booking_row.get("row_id"):
                    # Legacy single-row: check via booking_row's row_index
                    br_index = int(booking_row.get("row_index") or 0)
                    finalized = db.finalize_booking_row_if_ready(
                        run_id, br_index, spv_name
                    )
                    if finalized:
                        print(f"  ✅ [ROW {br_index}] {spv_name} fully done "
                              f"(NIS + all SAP steps complete).")
            except Exception as fin_exc:
                print(f"  ⚠️  Row finalization check failed: {str(fin_exc)[:180]}")

            try:
                _send_nis_completion_email(run_id)
            except Exception as post_exc:
                print(f"  ⚠️  Completion email check failed after SAP completion: {str(post_exc)[:180]}")

            # ── Cleanup ──────────────────────────────────────────────────
            if pdf_output_dir != TEMP_BOOKING_ROOT:
                _cleanup_tmp(pdf_output_dir)

            # Close SAP windows after successful task completion
            try:
                cleanup_sap_windows()
            except Exception as cleanup_exc:
                print(f"  ⚠️  SAP cleanup after success (non-fatal): {cleanup_exc}")

            lm.clear_context()
            return  # ← success

        except Exception as exc:
            err_str = str(exc)
            if sap_save_submitted:
                print(
                    "  ⚠️  SAP Save already completed, but post-SAP bookkeeping failed. "
                    "Not retrying SAP to avoid duplicate checklist creation."
                )
                print(f"     Last post-SAP error: {err_str[:220]}")
                _write_post_sap_done_marker(
                    run_id,
                    step_name,
                    step_done_detail or f"amount={abs_amount}  checklist=n/a  post_sap_error={err_str[:120]}",
                )
                try:
                    cleanup_sap_windows()
                except Exception as cleanup_exc:
                    print(f"  ⚠️  SAP cleanup error (non-fatal): {cleanup_exc}")
                lm.clear_context()
                return
            new_count = db.increment_step_retry(run_id, step_name, err_str)
            print(f"  ⚠️  SAP attempt {current_attempt} failed: {err_str[:120]}")

            # Close SAP windows before retry
            try:
                cleanup_sap_windows()
            except Exception as cleanup_exc:
                print(f"  ⚠️  SAP cleanup error (non-fatal): {cleanup_exc}")

            if new_count >= MAX_STEP_RETRIES:
                final_msg = (
                    f"SAP [{label}] failed after {MAX_STEP_RETRIES} "
                    f"attempt(s). Last error: {err_str[:200]}"
                )
                db.step_failed(run_id, step_name, final_msg)
                print(f"  ❌ {final_msg}")
                lm.log_error(final_msg)
                _send_failure_email(run_id, step_name, final_msg)
                lm.clear_context()
                return

            wait_sec = 5 * current_attempt
            print(f"  🔁 Retrying in {wait_sec}s "
                  f"({MAX_STEP_RETRIES - new_count} attempt(s) remaining) …")
            time.sleep(wait_sec)


# ─────────────────────────────────────────────────────────────────────────────
# Update Excel tracker after SAP completion
# ─────────────────────────────────────────────────────────────────────────────
def _update_tracker_after_sap(
    run_id: int,
    metadata_row: dict,
    spv_row: dict,
    spv_name: str,
    sap_checklist_number: str,
):
    """
    Update the DSM-NIS-OUTPUT Excel tracker with SAP checklist info.
    Only runs if all SAP steps for this run are complete.
    """
    if not db.are_all_sap_steps_complete(run_id):
        print("  [TRACKER] Not all SAP steps complete yet – deferring tracker update")
        return

    # Check if tracker was already updated
    step = db.STEP_UPDATE_TRACKER
    if db.is_step_done(run_id, step):
        print("  [TRACKER] Tracker already updated for this run")
        return

    print(f"  [TRACKER] All SAP steps complete – updating Excel tracker")
    db.step_start(run_id, step, detail=f"sheet={spv_name}")

    try:
        run_row              = db.get_run_by_id(run_id)
        checklist_id         = run_row.get("nis_checklist_id") or ""
        sap_checklist_value  = sap_checklist_number or run_row.get("checklist_value") or ""

        # Email received date
        email_received_date = ""
        raw_received = run_row.get("received_at") or ""
        if raw_received:
            try:
                dt = datetime.strptime(raw_received.replace("Z", "+00:00")[:19],
                                       "%Y-%m-%dT%H:%M:%S")
                email_received_date = dt.strftime("%d-%b-%Y")
            except Exception:
                email_received_date = raw_received[:10]

        nis_booking_date = datetime.utcnow().strftime("%d-%b-%Y")

        # Update column K on the existing NIS row (merge SAP checklist numbers,
        # comma-separated if multiple negatives produced multiple checklists).
        update_sap_checklist_on_existing_row(
            metadata_row         = metadata_row,
            sheet_name           = spv_name,
            sap_checklist_number = sap_checklist_value,
        )

        db.step_done(run_id, step,
                     detail=f"week={metadata_row.get('week_no')}  spv={spv_name}"
                            f"  sap_checklist={sap_checklist_value or 'n/a'}")
        print(f"  ✅ Tracker updated for run_id={run_id}")

    except Exception as exc:
        db.step_failed(run_id, step, str(exc))
        print(f"  ❌ Tracker update failed: {exc}")


# ─────────────────────────────────────────────────────────────────────────────
# Debit/Credit finance-reply PDF → SAP checklist
# ─────────────────────────────────────────────────────────────────────────────
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


def _write_debit_credit_output_row(pdf_row: dict, checklist_number: str):
    """Write a DSM-NIS-OUTPUT reply row for a finance-reply Debit/Credit SAP
    checklist (SPV sheet, NO NIS number, SAP number in col K). Best-effort:
    a failure here never fails the SAP checklist (already recorded in the DB)."""
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
        add_sap_checklist_reply_row(
            metadata_row=metadata_row,
            sheet_name=sheet_name,
            sap_checklist_number=checklist_number,
            cn_dn_date=cn_dn_date,
        )
    except Exception as exc:
        print(f"  [DC-SAP] Could not write OUTPUT reply row (checklist {checklist_number}): {exc}")


def _resolve_vendor_from_filename(row: dict) -> dict:
    """Resolve vendor/company/cost-center/plant for a Debit/Credit reply PDF
    from its filename (SPV name) → closest match in DSM-NIS-INPUT.

    Used as a fallback when vendor details were not resolved at reply-time.
    Returns a dict of debit_credit_pdfs columns to update, or {} on failure.
    """
    pdf_name = str(row.get("pdf_name") or "").strip()
    if not pdf_name:
        return {}
    try:
        from read_vendor_master_data import read_vendor_data_by_pdf_filename
        vendor = read_vendor_data_by_pdf_filename(pdf_name)
    except Exception as exc:
        print(f"  [DC-SAP] Vendor resolution failed for {pdf_name}: {str(exc)[:160]}")
        return {}
    if not vendor:
        print(f"  [DC-SAP] No SPV match in DSM-NIS-INPUT for {pdf_name}")
        return {}
    out = {
        "matched_spv_name": vendor.get("Matched SPV", ""),
        "vendor_code":      vendor.get("Vendor", ""),
        "company_code":     vendor.get("Company", ""),
        "cost_center":      vendor.get("Cost Center", ""),
        "plant":            vendor.get("Plant", ""),
    }
    print(f"  [DC-SAP] Resolved vendor for {pdf_name} → "
          f"spv={out['matched_spv_name']} vendor={out['vendor_code']}")
    return out


def get_pending_debit_credit_pdfs() -> list:
    """Finance-reply Debit/Credit PDFs still awaiting a SAP checklist."""
    if not getattr(config, "DEBIT_CREDIT_ENABLED", True):
        return []
    if not getattr(config, "DEBIT_CREDIT_SAP_ENABLED", True):
        return []
    try:
        return db.get_pending_sap_debit_credit_pdfs()
    except Exception as exc:
        print(f"  [DC-SAP] Could not read pending Debit/Credit PDFs: {exc}")
        return []


def process_debit_credit_sap_checklists() -> int:
    """Create a SAP checklist for each finance-reply Debit/Credit PDF that has
    its extracted values (vendor/company/cost-center/plant resolved from
    DSM-NIS-INPUT by the filename SPV, plus ref/amount/date from the PDF), store
    the checklist number, and upload the generated PDF to the shared
    booking/<from_date>_<to_date>/ SharePoint folder.

    Returns the number of checklists completed this pass. Rows missing data or
    unable to run SAP stay pending/failed for the next loop (retryable).
    """
    pending = get_pending_debit_credit_pdfs()
    if not pending:
        return 0

    print(f"\n  [DC-SAP] {len(pending)} finance-reply Debit/Credit PDF(s) pending SAP checklist")
    done = 0
    for row in pending:
        pdf_id = row.get("id")

        # If vendor details were not resolved at reply time, resolve them now
        # from the PDF filename → closest SPV match in DSM-NIS-INPUT.
        if not str(row.get("vendor_code") or "").strip():
            resolved = _resolve_vendor_from_filename(row)
            if resolved:
                db.update_debit_credit_pdf(pdf_id, resolved)
                row.update(resolved)

        # Validate the values SAP needs before attempting a checklist.
        missing = [f for f in ("vendor_code", "company_code", "ref_number", "doc_date")
                   if not str(row.get(f) or "").strip()]
        if safe_float(row.get("amount")) <= 0:
            missing.append("amount")
        if missing:
            db.mark_debit_credit_pdf_sap_failed(pdf_id, "missing values: " + ", ".join(missing))
            print(f"  [DC-SAP] PDF id={pdf_id} ({row.get('pdf_name')}) not ready — missing {missing}")
            continue

        try:
            cleanup_sap_windows()
        except Exception as cleanup_exc:
            print(f"  ⚠️  Pre-run SAP cleanup error (non-fatal): {cleanup_exc}")

        try:
            folder = sp.build_sap_checklist_folder_path(row.get("from_date"), row.get("to_date"))
            try:
                sp.ensure_sp_folder(folder)
            except Exception as exc:
                print(f"  [DC-SAP] Could not ensure folder {folder}: {exc}")

            out_dir = os.path.join(TEMP_BOOKING_ROOT, f"dc_sap_{pdf_id}")
            os.makedirs(out_dir, exist_ok=True)

            print(f"  [DC-SAP] Checklist for {row.get('pdf_name')} "
                  f"(spv={row.get('matched_spv_name') or row.get('spv_from_filename')} "
                  f"vendor={row.get('vendor_code')} company={row.get('company_code')} "
                  f"ref={row.get('ref_number')} amount={row.get('amount')} date={row.get('doc_date')})")

            # Strip trailing "-DSM" from ref_number — SAP expects just the
            # note reference (e.g. "CN/26-27/308") without the DSM suffix that
            # appears in the finance-reply PDF text.
            _raw_ref = str(row.get("ref_number") or "").strip()
            if _raw_ref.upper().endswith("-DSM"):
                _raw_ref = _raw_ref[:-4].strip()

            sap_result = run_sap_automation(
                company_code=str(row.get("company_code") or ""),
                vendor_code=str(row.get("vendor_code") or ""),
                amount=safe_float(row.get("amount")),
                cost_center=str(row.get("cost_center") or ""),
                plant=str(row.get("plant") or ""),
                deb_not_date=str(row.get("doc_date") or ""),
                deb_not_ref=_raw_ref,
                pdf_output_dir=out_dir,
            )

            checklist_number, sap_pdf_path = "", ""
            if isinstance(sap_result, tuple) and len(sap_result) >= 2:
                checklist_number = str(sap_result[0] or "")
                sap_pdf_path = str(sap_result[1] or "")
            elif isinstance(sap_result, str):
                checklist_number = sap_result or ""

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
            print(f"  [DC-SAP] ❌ id={pdf_id} failed: {str(exc)[:200]}")
        finally:
            try:
                cleanup_sap_windows()
            except Exception:
                pass

    return done


def install_dependencies():
    # Look for requirements.txt in the current directory or parent directory
    req_file = "requiredment.txt"
    if not os.path.exists(req_file):
        parent_req = os.path.join("..", "requirements.txt")
        if os.path.exists(parent_req):
            req_file = parent_req
        else:
            print("No requirements.txt or requiredment.txt found.")
            return

    print(f"Installing dependencies from {req_file}...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", req_file])
        print("Dependencies installed successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Failed to install dependencies: {e}")
        sys.exit(1)


def _is_nis_actively_processing(run: dict) -> bool:
    """True only when NIS is genuinely still booking rows for this run.

    A run in 'running' status may simply be awaiting SAP (all its booking rows
    are in nis_done/done/skipped/failed). In that case NIS is NOT actively
    processing — it already finished its part and SAP should proceed.
    """
    run_id = str(run.get("id") or "")
    if not run_id:
        return False
    rows = db.get_booking_rows(run_id)
    if not rows:
        # No booking rows seeded yet → NIS is still in early extraction stages
        return True
    # NIS is done with its part when every row has reached a post-NIS state
    nis_terminal = (db.STATUS_DONE, db.STATUS_SKIPPED, db.STATUS_FAILED, db.STATUS_NIS_DONE)
    return not all(r.get("status") in nis_terminal for r in rows)


def wait_for_nis_to_finish(interval_sec: int = 60, max_waits: int = 0):
    """Wait until no NIS email run is actively being processed.

    Runs whose booking rows are ALL in nis_done/done/skipped/failed are NOT
    considered active — their NIS work is complete and SAP should proceed.

    In one-shot mode ``max_waits`` is 0, so this returns immediately without
    blocking: if NIS happens to be mid-run the pending SAP tasks will simply be
    picked up on the next SAP run instead of holding this terminal open.
    Polling mode passes a positive ``max_waits`` (or -1 to wait indefinitely).
    """
    waits = 0
    while True:
        running = [
            run for run in db.get_pending_runs()
            if str(run.get("status") or "").lower() == db.STATUS_RUNNING
            and _is_nis_actively_processing(run)
        ]
        if not running:
            return

        # Bounded wait: don't hold the terminal open indefinitely.
        if max_waits == 0 or (max_waits > 0 and waits >= max_waits):
            print(
                f"  [WAIT] NIS still running for {len(running)} run(s); not waiting. "
                f"Pending SAP work will be handled on the next run."
            )
            return

        names = ", ".join(
            str(run.get("spv_name") or run.get("subject") or run.get("id") or "unknown")
            for run in running
        )
        print(
            f"  [WAIT] NIS is still running for {len(running)} run(s): {names}. "
            f"Checking again in {interval_sec} seconds..."
        )
        waits += 1
        time.sleep(interval_sec)

# ─────────────────────────────────────────────────────────────────────────────
# Main loop
# ─────────────────────────────────────────────────────────────────────────────
def run_once():
    """
    Single-shot: find all pending SAP tasks and process them sequentially.
    """
    install_dependencies()
    print(f"  SAP AUTOMATION  |  ONE-SHOT RUN")
    print(f"  {_ts()}")
    print(f"{'═' * 65}\n")

    db.init_db()
    wait_for_nis_to_finish()
    send_ready_nis_completion_emails()

    # Reset any crashed/stuck SAP steps first.
    # Check incomplete runs AND any run that has pending SAP steps (the NIS
    # service marks runs 'done' once its part finishes, but SAP steps may still
    # be stuck in 'running'/'failed' from a prior crash).
    _reset_targets = set()
    for run in db.get_incomplete_runs():
        _reset_targets.add(str(run.get("id") or ""))
    for step in db.get_all_pending_sap_steps():
        _reset_targets.add(str(step.get("run_id") or ""))
    for run_id in _reset_targets:
        if run_id:
            reset = db.reset_crashed_sap_steps(run_id, MAX_STEP_RETRIES)
            if reset:
                print(f"  [RESET] Reset {reset} crashed SAP step(s) for run_id={run_id}")

    # Find and process pending tasks
    tasks = get_pending_sap_tasks()
    dc_pending = get_pending_debit_credit_pdfs()

    if not tasks and not dc_pending:
        print("[INFO] No pending SAP automation tasks found.\n")
        return

    print(f"[INFO] Found {len(tasks)} typed SAP task(s) and "
          f"{len(dc_pending)} finance-reply Debit/Credit PDF(s):\n")
    for t in tasks:
        print(f"  → run_id={t['run_id']}  step={t['step_name']}")
    for p in dc_pending:
        print(f"  → debit/credit pdf id={p.get('id')}  {p.get('pdf_name')}  "
              f"spv={p.get('matched_spv_name') or p.get('spv_from_filename')}")
    print()

    success = 0
    failed  = 0
    for task in tasks:
        try:
            process_sap_task(task)
            success += 1
        except Exception as exc:
            print(f"  ❌ Task failed: run_id={task['run_id']} step={task['step_name']}")
            print(f"     Error: {str(exc)[:200]}")
            failed += 1

    # ── Debit/Credit finance-reply PDFs → SAP checklists ──────────────
    dc_done = 0
    try:
        dc_done = process_debit_credit_sap_checklists()
    except Exception as exc:
        print(f"  ❌ Debit/Credit SAP processing error: {str(exc)[:200]}")

    # ── Final SAP cleanup: close all SAP windows entirely ─────────────
    print("\n  🔒 Final cleanup: closing all SAP windows entirely …")
    try:
        cleanup_sap_windows()
    except Exception as cleanup_exc:
        print(f"  ⚠️  Final SAP cleanup error (non-fatal): {cleanup_exc}")

    print(f"\n{'═' * 65}")
    print(f"  SAP RUN COMPLETE  |  typed_success={success}  typed_failed={failed}"
          f"  debit_credit_done={dc_done}")
    print(f"  {_ts()}")
    send_ready_nis_completion_emails()
    print(f"{'═' * 65}\n")


def run_poll(interval_sec: int = 60):
    """
    Polling mode: continuously check for pending SAP tasks at fixed intervals.
    """
    print(f"\n{'═' * 65}")
    print(f"  SAP AUTOMATION  |  POLLING MODE  interval={interval_sec}s")
    print(f"  {_ts()}")
    print(f"{'═' * 65}\n")

    db.init_db()

    while True:
        # Reset crashed steps — include runs marked 'done' by NIS that still
        # have pending/stuck SAP steps.
        _reset_targets = set()
        for run in db.get_incomplete_runs():
            _reset_targets.add(str(run.get("id") or ""))
        for step in db.get_all_pending_sap_steps():
            _reset_targets.add(str(step.get("run_id") or ""))
        for run_id in _reset_targets:
            if run_id:
                db.reset_crashed_sap_steps(run_id, MAX_STEP_RETRIES)

        tasks = get_pending_sap_tasks()
        dc_pending = get_pending_debit_credit_pdfs()

        if tasks or dc_pending:
            print(f"\n[{_ts()}] Found {len(tasks)} typed SAP task(s) and "
                  f"{len(dc_pending)} Debit/Credit PDF(s)")
            for task in tasks:
                try:
                    process_sap_task(task)
                except Exception as exc:
                    print(f"  ❌ Task failed: {str(exc)[:200]}")

            # Debit/Credit finance-reply PDFs → SAP checklists
            try:
                process_debit_credit_sap_checklists()
            except Exception as exc:
                print(f"  ❌ Debit/Credit SAP processing error: {str(exc)[:200]}")

            # Close all SAP windows after processing the batch
            print(f"\n  🔒 Batch cleanup: closing all SAP windows entirely …")
            try:
                cleanup_sap_windows()
            except Exception as cleanup_exc:
                print(f"  ⚠️  Batch SAP cleanup error (non-fatal): {cleanup_exc}")
        else:
            print(f"[{_ts()}] No pending SAP tasks")

        print(f"[{_ts()}] Next poll in {interval_sec}s …\n")
        time.sleep(interval_sec)


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────
def main():
    install_dependencies()
    parser = argparse.ArgumentParser(
        description="SAP Automation Service – processes pending SAP tasks from Sheet DB",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument(
        "--poll",
        action="store_true",
        help="Run in polling mode (continuously check for pending tasks)",
    )
    parser.add_argument(
        "--interval",
        type=int,
        default=60,
        metavar="SECONDS",
        help="Poll interval in seconds (default: 60, only used with --poll)",
    )
    args = parser.parse_args()

    if args.poll:
        run_poll(args.interval)
    else:
        run_once()


if __name__ == "__main__":
    main()
