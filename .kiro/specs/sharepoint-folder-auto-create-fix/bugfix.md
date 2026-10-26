# Bugfix Requirements Document

## Introduction

When the NIS booking or SAP automation pipeline uploads PDFs (email.pdf, SAP checklist, NIS checklist) to SharePoint, the target booking folder must exist. Currently, two code paths skip the `ensure_sp_folder()` call before uploading, causing silent upload failures when the folder does not yet exist on SharePoint.

The folder path structure is: `booking/<SPV name>/<from_date to to_date>/`

Affected upload sites:
- `sap_automation/sap_main.py` → `process_sap_task()` uploads SAP checklist PDF without ensuring the folder exists first.
- `nis_booking/main.py` → `_handle_checklist_email()` uploads NIS/SAP checklist PDFs from email without ensuring the folder exists first.

The correct pattern already exists in `nis_booking/main.py` → `_save_email_to_booking_folder()`, which calls `sp.ensure_sp_folder(sp_folder)` before any uploads.

## Bug Analysis

### Current Behavior (Defect)

1.1 WHEN `sap_automation/sap_main.py` `process_sap_task()` uploads the SAP checklist PDF to the booking folder AND the booking folder does not exist on SharePoint THEN the system silently fails the upload (HTTP error returned by Graph API, empty string returned by `upload_file_to_sp`)

1.2 WHEN `nis_booking/main.py` `_handle_checklist_email()` uploads NIS or SAP checklist PDFs to the booking folder AND the booking folder does not exist on SharePoint THEN the system silently fails the upload (HTTP error returned by Graph API, empty string returned by `upload_file_to_sp`)

1.3 WHEN the `upload_file_to_sp()` function is called with a folder path that does not exist on SharePoint THEN the system does not auto-create the folder and returns an empty string indicating failure

### Expected Behavior (Correct)

2.1 WHEN `sap_automation/sap_main.py` `process_sap_task()` uploads the SAP checklist PDF to the booking folder THEN the system SHALL call `ensure_sp_folder(booking_folder)` before the upload so the folder is created if it does not exist

2.2 WHEN `nis_booking/main.py` `_handle_checklist_email()` uploads NIS or SAP checklist PDFs to the booking folder THEN the system SHALL call `ensure_sp_folder(booking_folder)` before the upload so the folder is created if it does not exist

2.3 WHEN any PDF upload targets a SharePoint folder that does not exist THEN the system SHALL ensure the target folder exists (by creating it) before attempting the upload, preventing silent failures

### Unchanged Behavior (Regression Prevention)

3.1 WHEN `nis_booking/main.py` `_save_email_to_booking_folder()` uploads email.pdf and invoice PDFs THEN the system SHALL CONTINUE TO call `ensure_sp_folder()` before uploads (existing correct behavior)

3.2 WHEN the booking folder already exists on SharePoint THEN the system SHALL CONTINUE TO upload PDFs successfully without errors (ensure_sp_folder is idempotent — it skips creation if the folder exists)

3.3 WHEN debit/credit PDF uploads occur in both `nis_booking/main.py` and `sap_main.py` THEN the system SHALL CONTINUE TO call `ensure_sp_folder()` before those uploads (existing correct behavior at lines 443 and 1127 respectively)

3.4 WHEN the `upload_file_to_sp()` function is called with a valid existing folder path THEN the system SHALL CONTINUE TO upload the file and return the SharePoint item path on success
