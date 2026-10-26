# SharePoint Folder Auto-Create Fix — Bugfix Design

## Overview

PDF uploads to SharePoint fail silently when the target booking folder does not yet exist, because `ensure_sp_folder()` is not called before the upload in two specific code paths:

1. `sap_automation/sap_main.py` → `process_sap_task()` — uploads the SAP checklist PDF directly without ensuring the folder exists.
2. `nis_booking/main.py` → `_handle_checklist_email()` — uploads NIS/SAP checklist PDFs from email without ensuring the folder exists.

The fix adds `ensure_sp_folder(booking_folder)` calls before the affected uploads, consistent with the pattern already used in `_save_email_to_booking_folder()` and the debit/credit flows. Optionally, `upload_file_to_sp()` itself can be hardened to defensively ensure the folder exists before uploading.

## Glossary

- **Bug_Condition (C)**: An upload to SharePoint is attempted on a booking folder path that does not yet exist on the drive, AND no prior `ensure_sp_folder()` call has been made in that code path.
- **Property (P)**: The target folder is guaranteed to exist before any `upload_file_to_sp()` call, so the upload succeeds (returns a non-empty SharePoint item path).
- **Preservation**: Existing code paths that already call `ensure_sp_folder()` before uploads must continue to work identically. Uploads to already-existing folders must remain unaffected (idempotent).
- **`ensure_sp_folder(path)`**: Checks if the SharePoint folder exists via Graph API; if not, creates it segment-by-segment. Returns the folder path. Idempotent (safe to call if folder already exists).
- **`upload_file_to_sp(local, sp_folder, filename)`**: Uploads a local file to the specified SharePoint folder. Returns the item path on success, or `""` on failure.
- **`booking_folder`**: The SharePoint path `<SP_BOOKING_BASE>/booking/<SPV>/<from_date to to_date>/` where all PDFs for a booking run are stored.

## Bug Details

### Bug Condition

The bug manifests when the SAP automation or the NIS checklist-email handler attempts to upload a PDF to the booking folder on SharePoint, but that folder has not yet been created. The Graph API returns an HTTP error (typically 404 for the parent path), `upload_file_to_sp()` returns `""`, and the upload is silently lost.

**Formal Specification:**
```
FUNCTION isBugCondition(input)
  INPUT: input of type UploadRequest { local_path: str, booking_folder: str, caller: str }
  OUTPUT: boolean

  RETURN input.caller IN ['process_sap_task', '_handle_checklist_email']
         AND input.booking_folder != ""
         AND NOT folder_exists_on_sharepoint(input.booking_folder)
END FUNCTION
```

### Examples

- **SAP checklist upload**: `process_sap_task()` completes SAP automation, captures the checklist PDF locally as `/tmp/sap_checklist_12345.pdf`, attempts `sp.upload_file_to_sp(path, "DSM/booking/AHEJ4L/23-Jan-2026 to 28-Jan-2026", "sap_checklist_12345.pdf")`. The folder does not exist → upload returns `""` → PDF is never stored on SharePoint.
- **NIS checklist email**: An email arrives with subject "NIS Checklist 98765". `_handle_checklist_email()` extracts the PDF and calls `sp.upload_file_to_sp(pdf, "DSM/booking/AHEJ4L/23-Jan-2026 to 28-Jan-2026", "nis_checklist_98765.pdf")`. Same folder-missing scenario → silent failure.
- **Existing folder (no bug)**: A second upload to the same booking folder after `_save_email_to_booking_folder()` has already created it succeeds because the folder exists.
- **Debit/credit path (no bug)**: The debit/credit flow already calls `ensure_sp_folder()` at lines 443 and 1127 in the respective files, so those uploads always succeed.

## Expected Behavior

### Preservation Requirements

**Unchanged Behaviors:**
- `_save_email_to_booking_folder()` in `nis_booking/main.py` continues to call `ensure_sp_folder()` before uploading email.pdf and invoice PDFs (existing correct behavior at line 1111).
- Debit/credit upload flows in both `nis_booking/main.py` (line 443, 643) and `sap_main.py` (line 1127) continue to call `ensure_sp_folder()` before their uploads.
- When the booking folder already exists, `ensure_sp_folder()` returns immediately after a single GET (idempotent) — no duplicate folder creation attempts, no errors.
- `upload_file_to_sp()` return value semantics remain unchanged: non-empty string on success, `""` on failure.
- All other SharePoint operations (download, list, file_exists) remain unaffected.

**Scope:**
All inputs where the booking folder already exists on SharePoint should be completely unaffected by this fix. The only observable change is that previously-failing uploads now succeed because the folder is auto-created.

## Hypothesized Root Cause

Based on the bug description and code analysis, the root cause is straightforward:

1. **Missing `ensure_sp_folder()` in `process_sap_task()`** (sap_automation/sap_main.py, ~line 798): The SAP checklist PDF upload block checks `if booking_folder and sap_pdf_full_path and os.path.isfile(sap_pdf_full_path)` then calls `sp.upload_file_to_sp()` directly. There is no preceding `ensure_sp_folder()` call. If this code path executes before the NIS booking flow has created the folder (or if the SAP automation runs independently), the folder may not exist.

2. **Missing `ensure_sp_folder()` in `_handle_checklist_email()`** (nis_booking/main.py, ~line 230): The checklist email handler retrieves `booking_folder` from the DB, collects PDF paths, and immediately loops over them calling `sp.upload_file_to_sp()`. No folder-existence check is made. If the checklist email arrives before the main booking email (which triggers `_save_email_to_booking_folder()` and its `ensure_sp_folder()` call), the folder won't exist.

3. **No defensive layer in `upload_file_to_sp()`**: The upload function assumes the target folder exists. It does not call `ensure_sp_folder()` internally. This is by design (separation of concerns) but leaves callers responsible for folder creation — a responsibility the two identified callers fail to fulfil.

## Correctness Properties

Property 1: Bug Condition - Folder Auto-Created Before Upload

_For any_ upload request where the caller is `process_sap_task` or `_handle_checklist_email`, the booking folder is non-empty, and the folder does not yet exist on SharePoint, the fixed code SHALL call `ensure_sp_folder(booking_folder)` before invoking `upload_file_to_sp()`, guaranteeing the folder exists so the upload can succeed.

**Validates: Requirements 2.1, 2.2, 2.3**

Property 2: Preservation - Existing Upload Paths Unchanged

_For any_ upload request where the booking folder already exists on SharePoint (or where `ensure_sp_folder()` is already called by the existing code path), the fixed code SHALL produce exactly the same behavior as the original code, preserving successful uploads, return values, logging, and DB state updates.

**Validates: Requirements 3.1, 3.2, 3.3, 3.4**

## Fix Implementation

### Changes Required

Assuming our root cause analysis is correct:

**File**: `sap_automation/sap_main.py`

**Function**: `process_sap_task()`

**Specific Changes**:
1. **Add `ensure_sp_folder()` before SAP checklist upload** (around line 798): Insert `sp.ensure_sp_folder(booking_folder)` immediately before `sp.upload_file_to_sp(sap_pdf_full_path, booking_folder, dst_basename)`, inside the existing `if booking_folder and sap_pdf_full_path and os.path.isfile(sap_pdf_full_path):` guard.

---

**File**: `nis_booking/main.py`

**Function**: `_handle_checklist_email()`

**Specific Changes**:
2. **Add `ensure_sp_folder()` before checklist PDF upload loop** (before `for pdf_path in pdf_paths:`): Insert `sp.ensure_sp_folder(booking_folder)` after the `if not pdf_paths:` early-return check and before the upload loop begins.

---

**File**: `sap_automation/sharepoint_upload.py` and `nis_booking/sharepoint_upload.py`

**Function**: `upload_file_to_sp()`

**Specific Changes (optional defensive layer)**:
3. **Add defensive `ensure_sp_folder()` inside `upload_file_to_sp()`**: At the start of the function, after the `if not os.path.isfile(local_path):` check, add `ensure_sp_folder(sp_folder_path)`. This provides a safety net for any future callers that forget to ensure the folder. Since `ensure_sp_folder()` is idempotent (single GET when folder exists), the performance impact is one extra HTTP call per upload that can be accepted for reliability.

---

**Summary of changes**:
| # | File | Location | Change |
|---|------|----------|--------|
| 1 | `sap_automation/sap_main.py` | Line ~798, inside the upload guard | Add `sp.ensure_sp_folder(booking_folder)` |
| 2 | `nis_booking/main.py` | Line ~230, before the `for pdf_path` loop | Add `sp.ensure_sp_folder(booking_folder)` |
| 3 | Both `sharepoint_upload.py` | Inside `upload_file_to_sp()`, after file-exists check | Add `ensure_sp_folder(sp_folder_path)` |

## Testing Strategy

### Validation Approach

The testing strategy follows a two-phase approach: first, surface counterexamples that demonstrate the bug on unfixed code, then verify the fix works correctly and preserves existing behavior.

### Exploratory Bug Condition Checking

**Goal**: Surface counterexamples that demonstrate the bug BEFORE implementing the fix. Confirm or refute the root cause analysis. If we refute, we will need to re-hypothesize.

**Test Plan**: Mock the SharePoint Graph API responses to simulate a non-existent folder (return 404 on GET, then verify whether `ensure_sp_folder` is called). Run these tests against the UNFIXED code to observe failures.

**Test Cases**:
1. **SAP Checklist Upload — Missing Folder**: Call `process_sap_task()` with a valid `booking_folder` and `sap_pdf_full_path` but mock the Graph API to return 404 on the folder path. Assert that `upload_file_to_sp()` returns `""` (will fail on unfixed code because no `ensure_sp_folder` is called).
2. **NIS Checklist Email — Missing Folder**: Invoke `_handle_checklist_email()` with a matching NIS checklist email, a DB run that has a `booking_folder_path` set, but mock SharePoint to return 404 on that folder. Assert uploads return `""` (will fail on unfixed code).
3. **Upload Function Direct Call**: Call `upload_file_to_sp(valid_local_file, non_existent_folder, name)` and confirm it returns `""` without creating the folder (demonstrates the lack of defensive behavior in unfixed code).

**Expected Counterexamples**:
- `upload_file_to_sp()` returns `""` because the Graph API returns a 404 or error for the non-existent folder path
- `ensure_sp_folder()` is never called in the affected code paths

### Fix Checking

**Goal**: Verify that for all inputs where the bug condition holds, the fixed function produces the expected behavior.

**Pseudocode:**
```
FOR ALL input WHERE isBugCondition(input) DO
  result := fixed_upload_path(input)
  ASSERT ensure_sp_folder_was_called(input.booking_folder)
  ASSERT result.upload_return_value != ""  // upload succeeds
END FOR
```

### Preservation Checking

**Goal**: Verify that for all inputs where the bug condition does NOT hold, the fixed function produces the same result as the original function.

**Pseudocode:**
```
FOR ALL input WHERE NOT isBugCondition(input) DO
  ASSERT original_upload_path(input) = fixed_upload_path(input)
END FOR
```

**Testing Approach**: Property-based testing is recommended for preservation checking because:
- It generates many combinations of existing/non-existing folders, valid/invalid paths, and different callers
- It catches edge cases where `ensure_sp_folder()` might alter behavior (e.g., empty folder strings, special characters)
- It provides strong guarantees that behavior is unchanged for all non-buggy inputs

**Test Plan**: Observe behavior on UNFIXED code first for uploads where the folder already exists, then write property-based tests capturing that behavior to confirm it stays the same after the fix.

**Test Cases**:
1. **Existing Folder Upload Preservation**: Mock the folder as existing (GET returns 200), run upload in fixed code, verify result matches unfixed code behavior exactly.
2. **Empty Booking Folder Path**: Both unfixed and fixed code should skip the upload and return early. Verify no `ensure_sp_folder("")` call is made.
3. **Missing Local File**: Both unfixed and fixed code should return `""` immediately. Verify `ensure_sp_folder` is not called when the local file doesn't exist.
4. **Debit/Credit Flow Unchanged**: Run the debit/credit upload flow end-to-end, verify `ensure_sp_folder` is still called at its existing locations and the upload succeeds as before.

### Unit Tests

- Test `ensure_sp_folder()` idempotency: calling it on an already-existing folder does not error and returns the path.
- Test `ensure_sp_folder()` creation: calling it on a non-existent multi-level path creates each segment.
- Test `upload_file_to_sp()` with defensive `ensure_sp_folder()`: verify that even without a prior explicit call, the upload succeeds.
- Test `_handle_checklist_email()` with mocked SharePoint: verify `ensure_sp_folder` is called before uploads.
- Test `process_sap_task()` upload block: verify `ensure_sp_folder` is called before the SAP checklist upload.

### Property-Based Tests

- Generate random booking folder paths (with various SPV names, date formats, special characters) and verify that `ensure_sp_folder()` + `upload_file_to_sp()` always results in a successful upload when the local file exists and the Graph API is reachable.
- Generate random inputs where `booking_folder` is empty or `sap_pdf_full_path` does not exist, and verify the code skips both `ensure_sp_folder` and `upload_file_to_sp` (preservation of guard conditions).
- Generate various email subjects (matching and non-matching checklist patterns) and verify that only matching emails trigger the upload flow.

### Integration Tests

- End-to-end test: trigger SAP automation on a run whose booking folder does not exist → verify folder is created and PDF is uploaded.
- End-to-end test: send a checklist email for a run whose booking folder does not exist → verify folder is created and PDF is uploaded.
- End-to-end test: send a checklist email for a run whose booking folder already exists → verify upload succeeds without errors (idempotent).
- Verify that the completion email includes the correct `booking_folder` path after successful upload.
