# Bugfix Requirements Document

## Introduction

When a DSM email contains a multi-row SPV table (3-row or 5-row), the system creates only ONE SharePoint booking folder using the first SPV name encountered. All subsequent SPV rows in that email incorrectly share the same folder path, causing files from different SPVs to be co-mingled and per-row `booking_folder_path` values to point to the wrong location. This breaks downstream processes (NIS checklist upload, SAP checklist upload) that rely on each row's `booking_folder_path` to upload files to the correct SPV-specific folder.

## Bug Analysis

### Current Behavior (Defect)

1.1 WHEN a multi-row DSM email (3-row or 5-row table) contains multiple unique SPVs THEN the system creates only ONE SharePoint folder using the first SPV name found in the table

1.2 WHEN a multi-row DSM email is processed THEN the system stores the same single `booking_folder_path` on ALL booking rows in the database regardless of their SPV name

1.3 WHEN `email.pdf` and invoice PDF attachments are saved to SharePoint THEN the system uploads them to only the first SPV's folder, leaving other SPV folders without these shared files

1.4 WHEN NIS checklist or SAP checklist PDFs are later uploaded for a non-first SPV row THEN the system uploads them to the wrong folder (first SPV's folder) because the row's `booking_folder_path` points there

### Expected Behavior (Correct)

2.1 WHEN a multi-row DSM email contains multiple unique SPVs THEN the system SHALL create a separate SharePoint folder for EACH unique SPV (using `_folder_spv_name()` to collapse Wind/Solar variants to their base code)

2.2 WHEN a multi-row DSM email is processed THEN the system SHALL store a per-SPV `booking_folder_path` on each booking row pointing to that row's own SPV folder (format: `<base>/booking/<SPV_NAME>/<from_date to to_date>/`)

2.3 WHEN `email.pdf` and invoice PDF attachments are saved to SharePoint THEN the system SHALL upload these shared files to ALL unique SPV folders created for that email

2.4 WHEN NIS checklist or SAP checklist PDFs are later uploaded for any SPV row THEN the system SHALL upload them to that row's own SPV folder (as resolved from the row's `booking_folder_path`)

### Unchanged Behavior (Regression Prevention)

3.1 WHEN a single-row DSM email (1 SPV) is processed THEN the system SHALL CONTINUE TO create exactly one folder and store its path on the single booking row

3.2 WHEN Wind/Solar variant SPVs (e.g. `AHEJ4L-Solar` and `AHEJ4L-Wind`) appear in the same email THEN the system SHALL CONTINUE TO collapse them into one shared folder using the base plant code (e.g. `booking/AHEJ4L/<dates>/`)

3.3 WHEN `_process_booking_row` or SAP automation reads a row's `booking_folder_path` to upload checklist PDFs THEN the system SHALL CONTINUE TO use the per-row `booking_folder_path` value without any change to those downstream functions

3.4 WHEN the debit/credit email registration stores `booking_folder_path` THEN the system SHALL CONTINUE TO use the first SPV's folder path for the debit/credit record (debit/credit is per-email, not per-SPV)

3.5 WHEN `build_booking_folder_path()` is called with an SPV name and date range THEN the system SHALL CONTINUE TO return the same path format (`<base>/booking/<SPV>/<from_date to to_date>`)
