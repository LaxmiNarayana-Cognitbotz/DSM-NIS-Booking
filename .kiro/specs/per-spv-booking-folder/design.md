# Per-SPV Booking Folder Bugfix Design

## Overview

The `_save_email_to_booking_folder()` function in `nis_booking/main.py` creates only one SharePoint booking folder per email (using the first SPV name found), even when the email contains a multi-row SPV table with several distinct SPVs. This causes all rows to share the same folder, co-mingling files and misdirecting downstream checklist uploads (NIS and SAP). The fix changes the function to iterate over all unique SPVs, create a folder for each, upload shared files (email.pdf, invoice PDFs) to every folder, and return a mapping so each booking row can be pointed to the correct folder.

## Glossary

- **Bug_Condition (C)**: A multi-row DSM email whose SPV table contains more than one unique folder label (after Wind/Solar collapsing via `_folder_spv_name()`)
- **Property (P)**: Each unique SPV gets its own SharePoint folder containing all shared files, and each booking row's `booking_folder_path` points to the correct folder
- **Preservation**: Single-SPV emails, Wind/Solar collapsing logic, downstream consumers of `booking_folder_path`, and debit/credit folder assignment must remain unchanged
- **`_save_email_to_booking_folder()`**: The function in `nis_booking/main.py` (~line 1120) that creates booking folders on SharePoint and uploads email.pdf + invoice PDFs
- **`_prepare_run()`**: The pipeline function in `nis_booking/main.py` (~line 1570) that calls `_save_email_to_booking_folder()` and seeds booking rows with their folder paths
- **`_folder_spv_name()`**: Helper that collapses Wind/Solar variant names (e.g. `AHEJ4L-Solar` → `AHEJ4L`) to produce a canonical folder label
- **`folder_map`**: The new return value from `_save_email_to_booking_folder()` — a dict mapping each unique folder label to its SharePoint folder path

## Bug Details

### Bug Condition

The bug manifests when a DSM email's SPV table contains multiple rows with distinct SPV names (after Wind/Solar collapsing). The `_save_email_to_booking_folder()` function breaks out of its loop after finding the first non-empty SPV name and creates only one folder, ignoring subsequent unique SPVs.

**Formal Specification:**
```
FUNCTION isBugCondition(input)
  INPUT: input of type {spv_dsm_rows: List[Dict], metadata_row: Dict}
  OUTPUT: boolean

  unique_labels := SET()
  FOR row IN input.spv_dsm_rows:
    name := row.get("spv_name", "").strip()
    IF name != "":
      unique_labels.add(_folder_spv_name(name))
  
  RETURN |unique_labels| > 1
END FUNCTION
```

### Examples

- **3-row email with 3 different SPVs** (e.g. `AHEJ4L`, `AWEMP1PL`, `AWEK06L`):
  - Expected: 3 folders created, email.pdf in all 3, each row points to its own folder
  - Actual: 1 folder (`AHEJ4L`), all 3 rows point to it

- **5-row email with Wind/Solar variants** (e.g. `AHEJ4L-Solar`, `AHEJ4L-Wind`, `AWEMP1PL`):
  - Expected: 2 folders (`AHEJ4L`, `AWEMP1PL`), email.pdf in both, Wind/Solar rows share `AHEJ4L` folder
  - Actual: 1 folder (`AHEJ4L`), all 5 rows point to it

- **Single-SPV email** (e.g. `AWEMP1PL` only):
  - Expected: 1 folder (`AWEMP1PL`) — no change from current behavior
  - Actual: 1 folder (`AWEMP1PL`) — correct (no bug triggered)

- **Multi-row email where all rows collapse to same label** (e.g. `AHEJ4L-Solar`, `AHEJ4L-Wind`):
  - Expected: 1 folder (`AHEJ4L`) — Wind/Solar collapsing produces one unique label
  - Actual: 1 folder (`AHEJ4L`) — correct (no bug triggered)

## Expected Behavior

### Preservation Requirements

**Unchanged Behaviors:**
- Single-SPV emails continue to produce exactly one folder and one stored path
- Wind/Solar collapsing via `_folder_spv_name()` continues to merge variant names into one base code
- `_process_booking_row()` continues to read `booking_folder_path` from its booking_row DB record without modification
- `sap_automation/sap_main.py` continues to read `booking_folder_path` from the booking_row record without modification
- `_handle_checklist_email()` continues to read `booking_folder_path` from the run/row record
- `build_booking_folder_path()` signature and return format remain unchanged
- Debit/credit email registration continues to use the first SPV's folder path (debit/credit is per-email)
- The `booking_folder_path` stored at the run level continues to be the first folder (backward compat)

**Scope:**
All inputs where `|unique_labels| == 1` (single unique SPV after collapsing) should produce behavior identical to the current code. This includes:
- Single-row DSM emails
- Multi-row emails where all SPVs collapse to the same folder label
- Any downstream function that reads `booking_folder_path` from a booking row

## Hypothesized Root Cause

Based on the bug description, the root cause is straightforward:

1. **Early-exit loop in `_save_email_to_booking_folder()`**: The function uses a `for row in (spv_dsm_rows or []): ... if spv_name: break` pattern that stops at the first non-empty SPV. It then creates exactly one folder from that single name.

2. **Single return value**: The function returns a single string (`sp_folder`), making it structurally impossible to communicate multiple folder paths back to the caller.

3. **Caller assigns one path to all rows**: In `_prepare_run()`, the single `booking_folder` string is stored on every booking row regardless of that row's SPV, because no per-SPV mapping exists.

4. **No iteration over unique SPVs for file upload**: The email.pdf conversion and invoice PDF upload happen once, targeting only the single folder.

## Correctness Properties

Property 1: Bug Condition - All Unique SPVs Get Dedicated Folders

_For any_ multi-row DSM email where the SPV table contains more than one unique folder label (after Wind/Solar collapsing), the fixed `_save_email_to_booking_folder` function SHALL create a separate SharePoint folder for each unique label, upload `email.pdf` and all invoice PDFs to every folder, and return a mapping from each label to its folder path.

**Validates: Requirements 2.1, 2.3**

Property 2: Bug Condition - Each Booking Row Gets Correct Folder Path

_For any_ multi-row DSM email where the SPV table contains more than one unique folder label, the fixed `_prepare_run` function SHALL assign each booking row's `booking_folder_path` to the folder corresponding to that row's SPV (resolved via `_folder_spv_name()`).

**Validates: Requirements 2.2, 2.4**

Property 3: Preservation - Single-SPV Behavior Unchanged

_For any_ email where all SPV rows collapse to a single unique folder label (including single-row emails), the fixed function SHALL produce the same single folder and the same stored paths as the original function.

**Validates: Requirements 3.1, 3.2, 3.5**

Property 4: Preservation - Downstream Consumer Interface Unchanged

_For any_ booking row processed by `_process_booking_row()` or SAP automation, the downstream functions SHALL continue to read `booking_folder_path` from the booking_row DB record without any API or schema change.

**Validates: Requirements 3.3, 3.4**

## Fix Implementation

### Changes Required

Assuming our root cause analysis is correct:

**File**: `nis_booking/main.py`

**Function**: `_save_email_to_booking_folder()`

**Specific Changes**:

1. **Collect all unique folder labels**: Replace the early-exit loop with a set comprehension that gathers all unique labels from `spv_dsm_rows` via `_folder_spv_name()`.

2. **Iterate and create folders for each label**: For each unique label, call `sp.build_booking_folder_path()` and `sp.ensure_sp_folder()`.

3. **Convert email.pdf once, upload to all folders**: Call `_convert_html_to_pdf()` once, then loop over all folders to upload.

4. **Upload invoice PDFs to all folders**: For each PDF in the temp dir, upload to every folder (using `shutil.copy` to preserve the local file for subsequent uploads).

5. **Return a dict instead of a string**: Return `{folder_label: sp_folder_path}` mapping.

**Function**: `_prepare_run()`

**Specific Changes**:

6. **Accept dict return value**: Change `booking_folder = _save_email_to_booking_folder(...)` to `folder_map = _save_email_to_booking_folder(...)`.

7. **Store first folder at run level**: For backward compatibility, save `next(iter(folder_map.values()))` as the run-level `booking_folder_path`.

8. **Assign per-row folder paths**: When iterating over `row_classifications` to persist folder paths on booking rows, resolve each row's folder using `folder_map.get(_folder_spv_name(spv_name), first_folder)`.

9. **Backward-compat string handling**: If `folder_map` is a string (from a previously stored path loaded from DB), treat it as a single-entry dict for the existing code path.

## Testing Strategy

### Validation Approach

The testing strategy follows a two-phase approach: first, surface counterexamples that demonstrate the bug on unfixed code, then verify the fix works correctly and preserves existing behavior.

### Exploratory Bug Condition Checking

**Goal**: Surface counterexamples that demonstrate the bug BEFORE implementing the fix. Confirm or refute the root cause analysis. If we refute, we will need to re-hypothesize.

**Test Plan**: Write tests that call `_save_email_to_booking_folder()` with multi-SPV `spv_dsm_rows` and verify how many SharePoint folders are created and what path is returned. Run these tests on the UNFIXED code to observe failures.

**Test Cases**:
1. **Multi-SPV 3-row email**: Pass `spv_dsm_rows` with 3 distinct SPVs (will create only 1 folder on unfixed code)
2. **Multi-SPV 5-row email with Wind/Solar**: Pass rows including `AHEJ4L-Solar`, `AHEJ4L-Wind`, `AWEMP1PL` (will create only 1 folder on unfixed code)
3. **Per-row folder assignment**: After calling `_prepare_run()`, check each booking row's `booking_folder_path` (all will point to same folder on unfixed code)
4. **Shared file presence**: Check if email.pdf exists in all SPV folders (will only exist in first folder on unfixed code)

**Expected Counterexamples**:
- `_save_email_to_booking_folder()` returns a single string path instead of a dict
- Only one `sp.ensure_sp_folder()` call is made regardless of SPV count
- All booking rows share the same `booking_folder_path` value

### Fix Checking

**Goal**: Verify that for all inputs where the bug condition holds, the fixed function produces the expected behavior.

**Pseudocode:**
```
FOR ALL input WHERE isBugCondition(input) DO
  folder_map := _save_email_to_booking_folder_fixed(input.metadata_row, input.spv_dsm_rows, ...)
  unique_labels := {_folder_spv_name(row["spv_name"]) FOR row IN input.spv_dsm_rows IF row["spv_name"]}
  
  ASSERT type(folder_map) == dict
  ASSERT set(folder_map.keys()) == unique_labels
  FOR label IN unique_labels:
    ASSERT "email.pdf" uploaded to folder_map[label]
    ASSERT all invoice PDFs uploaded to folder_map[label]
END FOR
```

### Preservation Checking

**Goal**: Verify that for all inputs where the bug condition does NOT hold, the fixed function produces the same result as the original function.

**Pseudocode:**
```
FOR ALL input WHERE NOT isBugCondition(input) DO
  folder_map := _save_email_to_booking_folder_fixed(input.metadata_row, input.spv_dsm_rows, ...)
  original_path := _save_email_to_booking_folder_original(input.metadata_row, input.spv_dsm_rows, ...)
  
  ASSERT len(folder_map) == 1
  ASSERT next(iter(folder_map.values())) == original_path
END FOR
```

**Testing Approach**: Property-based testing is recommended for preservation checking because:
- It generates many single-SPV configurations automatically
- It verifies Wind/Solar collapsing produces the same single folder
- It catches edge cases (empty SPV names, whitespace-only names)

**Test Plan**: Observe behavior on UNFIXED code first for single-SPV emails, then write property-based tests capturing that behavior.

**Test Cases**:
1. **Single-SPV preservation**: Verify single-row emails produce identical folder path and structure
2. **Wind/Solar collapsing preservation**: Verify `AHEJ4L-Solar` + `AHEJ4L-Wind` still produce one shared folder
3. **Run-level path preservation**: Verify `booking_folder_path` at the run level is still set (first folder)
4. **Debit/credit path preservation**: Verify `_maybe_register_debit_credit_email` still receives a valid folder path

### Unit Tests

- Test `_save_email_to_booking_folder()` with 1-SPV input returns single-entry dict
- Test `_save_email_to_booking_folder()` with 3-distinct-SPV input returns 3-entry dict
- Test `_save_email_to_booking_folder()` with Wind/Solar variants returns collapsed dict
- Test that `email.pdf` upload is called once per unique folder
- Test that invoice PDFs are uploaded to every unique folder
- Test edge case: all SPV names empty (should create one folder with empty/default label)

### Property-Based Tests

- Generate random `spv_dsm_rows` lists with varying SPV names and verify folder count equals unique label count
- Generate random mixtures of Wind/Solar variants and non-variant SPVs to verify correct collapsing
- Generate single-SPV inputs and verify output matches original function behavior exactly

### Integration Tests

- Test full `_prepare_run()` pipeline with a multi-SPV email and verify each booking row's `booking_folder_path` is distinct per SPV
- Test that downstream `_process_booking_row()` still functions correctly when given a per-row folder path
- Test debit/credit email registration receives the first folder from the map
