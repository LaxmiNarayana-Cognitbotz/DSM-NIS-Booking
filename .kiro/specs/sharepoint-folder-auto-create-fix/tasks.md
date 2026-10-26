# Implementation Plan

- [ ] 1. Write bug condition exploration test
  - **Property 1: Bug Condition** - Missing Folder Causes Silent Upload Failure
  - **CRITICAL**: This test MUST FAIL on unfixed code - failure confirms the bug exists
  - **DO NOT attempt to fix the test or the code when it fails**
  - **NOTE**: This test encodes the expected behavior - it will validate the fix when it passes after implementation
  - **GOAL**: Surface counterexamples that demonstrate the bug exists in `process_sap_task()` and `_handle_checklist_email()`
  - **Scoped PBT Approach**: Scope the property to concrete failing cases: upload requests where `booking_folder` is non-empty, the local PDF file exists, but the SharePoint folder does not exist (Graph API returns 404)
  - Test that `upload_file_to_sp(valid_local_file, non_existent_folder, filename)` returns `""` when called from `process_sap_task()` or `_handle_checklist_email()` WITHOUT a preceding `ensure_sp_folder()` call
  - Mock Graph API: folder GET returns 404, upload PUT returns error (404 or 409)
  - Bug Condition from design: `isBugCondition(input)` where `input.caller IN ['process_sap_task', '_handle_checklist_email'] AND input.booking_folder != "" AND NOT folder_exists_on_sharepoint(input.booking_folder)`
  - Expected behavior assertions: `ensure_sp_folder(booking_folder)` is called before `upload_file_to_sp()`, and the upload returns a non-empty path (success)
  - Run test on UNFIXED code
  - **EXPECTED OUTCOME**: Test FAILS (this is correct - it proves the bug exists because `ensure_sp_folder` is never called in those code paths)
  - Document counterexamples: e.g., "upload_file_to_sp('/tmp/sap_checklist.pdf', 'DSM/booking/AHEJ4L/23-Jan-2026 to 28-Jan-2026', 'sap_checklist.pdf') returns '' because folder does not exist and ensure_sp_folder was not called"
  - Mark task complete when test is written, run, and failure is documented
  - _Requirements: 1.1, 1.2, 1.3_

- [ ] 2. Write preservation property tests (BEFORE implementing fix)
  - **Property 2: Preservation** - Existing Upload Paths Remain Unchanged
  - **IMPORTANT**: Follow observation-first methodology
  - Observe on UNFIXED code: `upload_file_to_sp(valid_file, existing_folder, name)` returns non-empty path when folder exists (Graph API GET returns 200, PUT returns 201)
  - Observe on UNFIXED code: `upload_file_to_sp(non_existent_file, any_folder, name)` returns `""` immediately (file-not-found guard)
  - Observe on UNFIXED code: `ensure_sp_folder(existing_folder)` returns the folder path after a single GET (idempotent, no creation)
  - Observe on UNFIXED code: `_save_email_to_booking_folder()` calls `ensure_sp_folder()` before uploads (existing correct behavior)
  - Write property-based tests: for all inputs where the booking folder already exists on SharePoint, `upload_file_to_sp()` returns a non-empty path (success)
  - Write property-based tests: for all inputs where `local_path` does not exist as a file, `upload_file_to_sp()` returns `""` and `ensure_sp_folder` is NOT called
  - Write property-based tests: for all inputs where `booking_folder` is empty string, the upload is skipped (guard conditions preserved)
  - Write property-based tests: `ensure_sp_folder()` on already-existing folder does single GET and returns path (no creation POST)
  - Run tests on UNFIXED code
  - **EXPECTED OUTCOME**: Tests PASS (this confirms baseline behavior to preserve)
  - Mark task complete when tests are written, run, and passing on unfixed code
  - _Requirements: 3.1, 3.2, 3.3, 3.4_

- [ ] 3. Fix for SharePoint folder auto-creation before PDF uploads

  - [ ] 3.1 Add `ensure_sp_folder()` in `sap_automation/sap_main.py` → `process_sap_task()`
    - Inside the `if booking_folder and sap_pdf_full_path and os.path.isfile(sap_pdf_full_path):` guard (around line 798)
    - Add `sp.ensure_sp_folder(booking_folder)` immediately before the `sp.upload_file_to_sp(...)` call
    - This ensures the booking folder exists before uploading the SAP checklist PDF
    - _Bug_Condition: isBugCondition(input) where input.caller = 'process_sap_task' AND input.booking_folder != "" AND NOT folder_exists(input.booking_folder)_
    - _Expected_Behavior: ensure_sp_folder(booking_folder) is called → folder created → upload_file_to_sp() returns non-empty path_
    - _Preservation: Existing debit/credit upload paths in sap_main.py (line 1127) already call ensure_sp_folder and remain unchanged_
    - _Requirements: 2.1_

  - [ ] 3.2 Add `ensure_sp_folder()` in `nis_booking/main.py` → `_handle_checklist_email()`
    - After the `if not pdf_paths:` early-return check (around line 230)
    - Add `sp.ensure_sp_folder(booking_folder)` before the `for pdf_path in pdf_paths:` loop
    - This ensures the booking folder exists before uploading any checklist PDFs from email
    - _Bug_Condition: isBugCondition(input) where input.caller = '_handle_checklist_email' AND input.booking_folder != "" AND NOT folder_exists(input.booking_folder)_
    - _Expected_Behavior: ensure_sp_folder(booking_folder) is called → folder created → upload_file_to_sp() returns non-empty path for each PDF_
    - _Preservation: Existing _save_email_to_booking_folder() already calls ensure_sp_folder and remains unchanged_
    - _Requirements: 2.2_

  - [ ] 3.3 Add defensive `ensure_sp_folder()` in `nis_booking/sharepoint_upload.py` → `upload_file_to_sp()`
    - After the `if not os.path.isfile(local_path):` check at the start of the function
    - Add `ensure_sp_folder(sp_folder_path)` as a defensive safety net
    - Since `ensure_sp_folder()` is idempotent (single GET when folder exists), performance impact is one extra HTTP call per upload
    - _Bug_Condition: Any caller that forgets to call ensure_sp_folder before upload_file_to_sp_
    - _Expected_Behavior: Folder guaranteed to exist before upload attempt, preventing silent failures_
    - _Preservation: When folder already exists, ensure_sp_folder does a single GET and returns immediately — no behavior change for existing callers_
    - _Requirements: 2.3_

  - [ ] 3.4 Add defensive `ensure_sp_folder()` in `sap_automation/sharepoint_upload.py` → `upload_file_to_sp()`
    - After the `if not os.path.isfile(local_path):` check at the start of the function
    - Add `ensure_sp_folder(sp_folder_path)` as a defensive safety net (same pattern as 3.3)
    - _Bug_Condition: Any caller that forgets to call ensure_sp_folder before upload_file_to_sp_
    - _Expected_Behavior: Folder guaranteed to exist before upload attempt_
    - _Preservation: Idempotent — no behavior change when folder already exists_
    - _Requirements: 2.3_

  - [ ] 3.5 Verify bug condition exploration test now passes
    - **Property 1: Expected Behavior** - Missing Folder Causes Silent Upload Failure (now fixed)
    - **IMPORTANT**: Re-run the SAME test from task 1 - do NOT write a new test
    - The test from task 1 encodes the expected behavior (ensure_sp_folder is called before upload)
    - When this test passes, it confirms the expected behavior is satisfied: folder is auto-created before upload
    - Run bug condition exploration test from step 1
    - **EXPECTED OUTCOME**: Test PASSES (confirms bug is fixed — ensure_sp_folder is now called in all affected paths)
    - _Requirements: 2.1, 2.2, 2.3_

  - [ ] 3.6 Verify preservation tests still pass
    - **Property 2: Preservation** - Existing Upload Paths Remain Unchanged
    - **IMPORTANT**: Re-run the SAME tests from task 2 - do NOT write new tests
    - Run preservation property tests from step 2
    - **EXPECTED OUTCOME**: Tests PASS (confirms no regressions — existing folder uploads, file-not-found guards, and idempotent ensure_sp_folder behavior all unchanged)
    - Confirm all tests still pass after fix (no regressions)

- [ ] 4. Checkpoint - Ensure all tests pass
  - Run the full test suite to confirm no regressions
  - Verify exploration test (Property 1) passes — bug is fixed
  - Verify preservation tests (Property 2) pass — no side effects
  - Ensure all other existing tests in the project still pass
  - Ask the user if questions arise
