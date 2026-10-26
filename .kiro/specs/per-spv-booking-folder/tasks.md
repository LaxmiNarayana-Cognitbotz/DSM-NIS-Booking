# Implementation Plan

- [x] 1. Write bug condition exploration test
  - **Property 1: Bug Condition** - Multi-SPV Email Creates Only One Folder
  - **CRITICAL**: This test MUST FAIL on unfixed code - failure confirms the bug exists
  - **DO NOT attempt to fix the test or the code when it fails**
  - **NOTE**: This test encodes the expected behavior - it will validate the fix when it passes after implementation
  - **GOAL**: Surface counterexamples that demonstrate only one folder is created for multi-SPV emails
  - **Scoped PBT Approach**: Scope the property to concrete failing cases: `spv_dsm_rows` with 3 distinct SPVs (e.g. `AHEJ4L`, `AWEMP1PL`, `AWEK06L`) and verify folder creation count
  - Test that `_save_email_to_booking_folder(metadata_row, spv_dsm_rows, ...)` with `|unique_labels| > 1` returns a dict with one entry per unique label (from Bug Condition in design: `isBugCondition` holds when `|unique_labels| > 1`)
  - Assert: `type(folder_map) == dict` and `set(folder_map.keys()) == unique_labels`
  - Assert: `email.pdf` uploaded to ALL folders in `folder_map`
  - Assert: invoice PDFs uploaded to ALL folders in `folder_map`
  - Mock SharePoint calls (`sp.ensure_sp_folder`, `sp.upload_file_to_sp`) and `_convert_html_to_pdf` to isolate logic
  - Run test on UNFIXED code
  - **EXPECTED OUTCOME**: Test FAILS (unfixed code returns a single string, not a dict; only 1 folder is created)
  - Document counterexamples: e.g. "3-SPV input produces only 1 `ensure_sp_folder` call and returns string instead of dict"
  - Mark task complete when test is written, run, and failure is documented
  - _Requirements: 1.1, 1.2, 1.3, 2.1, 2.3_

- [x] 2. Write preservation property tests (BEFORE implementing fix)
  - **Property 2: Preservation** - Single-SPV Email Behavior Unchanged
  - **IMPORTANT**: Follow observation-first methodology
  - **Step 1 - Observe on UNFIXED code**:
    - Observe: single-row email with `spv_dsm_rows=[{"spv_name": "AWEMP1PL"}]` creates exactly 1 folder and returns its path as a string
    - Observe: multi-row email where all rows collapse to same label (e.g. `AHEJ4L-Solar`, `AHEJ4L-Wind`) creates exactly 1 folder
    - Observe: `build_booking_folder_path()` returns format `<base>/booking/<SPV>/<from_date to to_date>/`
  - **Step 2 - Write property-based tests capturing observed behavior**:
    - Property: for all `spv_dsm_rows` where `|{_folder_spv_name(row["spv_name"]) for row in rows}| == 1`, exactly 1 SharePoint folder is created
    - Property: for all Wind/Solar variant combinations that collapse to same base code, only 1 folder is created with that base code label
    - Property: for single-SPV emails, the returned value (after fix: single-entry dict) is equivalent to the original string return
  - Generate random single-SPV inputs and Wind/Solar variant combos using hypothesis strategies
  - Run tests on UNFIXED code
  - **EXPECTED OUTCOME**: Tests PASS (confirms baseline behavior to preserve)
  - Mark task complete when tests are written, run, and passing on unfixed code
  - _Requirements: 3.1, 3.2, 3.5_

- [x] 3. Fix for multi-SPV emails creating only one booking folder

  - [x] 3.1 Modify `_save_email_to_booking_folder()` to create per-SPV folders
    - Replace early-exit loop (`for row ... if spv_name: break`) with set comprehension collecting all unique labels via `_folder_spv_name()`
    - Iterate over unique labels: call `sp.build_booking_folder_path()` and `sp.ensure_sp_folder()` for each
    - Convert `email.pdf` ONCE via `_convert_html_to_pdf()`, then upload to ALL folders
    - Upload invoice PDFs to ALL folders (loop over `folder_map` entries)
    - Return `dict {folder_label: sp_folder_path}` instead of single string
    - _Bug_Condition: isBugCondition(input) where |unique_labels| > 1 after Wind/Solar collapsing_
    - _Expected_Behavior: folder_map has one entry per unique label; shared files in all folders_
    - _Preservation: Single-SPV inputs produce single-entry dict equivalent to original behavior_
    - _Requirements: 2.1, 2.3_

  - [x] 3.2 Modify `_prepare_run()` to handle dict return value
    - Accept dict `folder_map` from `_save_email_to_booking_folder()`
    - Store `next(iter(folder_map.values()))` as run-level `booking_folder_path` (backward compat)
    - When seeding booking rows, resolve per-row folder: `folder_map.get(_folder_spv_name(row_spv_name), first_folder)`
    - Add backward-compat guard: if `folder_map` is already a string (from DB reload), wrap as `{"default": folder_map}`
    - _Bug_Condition: isBugCondition(input) where |unique_labels| > 1_
    - _Expected_Behavior: each booking row's booking_folder_path points to its own SPV folder_
    - _Preservation: run-level booking_folder_path remains set for debit/credit compat_
    - _Requirements: 2.2, 2.4, 3.3, 3.4_

  - [x] 3.3 Verify bug condition exploration test now passes
    - **Property 1: Expected Behavior** - Multi-SPV Email Creates Per-SPV Folders
    - **IMPORTANT**: Re-run the SAME test from task 1 - do NOT write a new test
    - The test from task 1 encodes the expected behavior (dict return, multiple folders, shared files in all)
    - When this test passes, it confirms the expected behavior is satisfied
    - Run bug condition exploration test from step 1
    - **EXPECTED OUTCOME**: Test PASSES (confirms bug is fixed)
    - _Requirements: 2.1, 2.2, 2.3, 2.4_

  - [x] 3.4 Verify preservation tests still pass
    - **Property 2: Preservation** - Single-SPV Email Behavior Unchanged
    - **IMPORTANT**: Re-run the SAME tests from task 2 - do NOT write new tests
    - Run preservation property tests from step 2
    - **EXPECTED OUTCOME**: Tests PASS (confirms no regressions for single-SPV emails)
    - Confirm all tests still pass after fix (no regressions)
    - _Requirements: 3.1, 3.2, 3.3, 3.4, 3.5_

- [x] 4. Checkpoint - Ensure all tests pass
  - Run full test suite to ensure no regressions
  - Verify downstream consumers are unaffected:
    - `_process_booking_row()` reads `booking_folder_path` from its booking_row - no change needed
    - SAP automation reads `booking_folder_path` from booking_row - no change needed
    - `_handle_checklist_email()` reads from DB - no change needed
    - Debit/credit registration uses first folder (run-level path) - no change needed
  - Ensure all tests pass, ask the user if questions arise
