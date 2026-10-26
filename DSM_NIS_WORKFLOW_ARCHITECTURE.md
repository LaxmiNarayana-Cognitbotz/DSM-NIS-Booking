# DSM-NIS Booking Automation -- Architecture & Workflow Document

## 1. System Overview

The DSM-NIS Booking Automation is a **two-service** system that processes DSM (Deviation Settlement Mechanism) emails for AGEL (Adani Green Energy Limited). It automates the entire workflow from email ingestion through NIS portal booking, SAP checklist creation, HOD/FC approval tracking, PDF merging, and completion notifications.

```
+=====================================================================+
|                    DSM-NIS BOOKING AUTOMATION                         |
|=====================================================================|
|                                                                       |
|  +-----------------+         +------------------+                    |
|  |  NIS BOOKING    |         |  SAP AUTOMATION  |                    |
|  |  SERVICE        |         |  SERVICE         |                    |
|  |  (run_nis_      |         |  (run_sap_       |                    |
|  |   booking.py)   |         |   automation.py) |                    |
|  +-----------------+         +------------------+                    |
|         |                            |                                |
|         +----------+  +--------------+                                |
|                    |  |                                                |
|              +-----v--v------+                                        |
|              |   SHARED DB   |                                        |
|              | (PostgreSQL / |                                        |
|              |   SQLite)     |                                        |
|              +---------------+                                        |
|                                                                       |
+=====================================================================+
```

### Two Independent Services

| Service | Entry Point | Runs On | Responsibility |
|---------|-------------|---------|----------------|
| **NIS Booking** | `run_nis_booking.py` | Any host with Playwright | Email scan, table extraction, NIS portal booking (+ve), tracker update, debit/credit email, approval tracking |
| **SAP Automation** | `run_sap_automation.py` | Windows VM with SAP GUI | SAP checklist creation (-ve), debit/credit reply PDF checklists |

Both services share ONE database (PostgreSQL on Azure or local SQLite). They never call each other directly -- they communicate exclusively through the shared DB state.

---

## 2. Email Types & Subject Triggers

The system monitors `agel.nis.automation@adani.com` via Microsoft Graph API (Inbox only). Each incoming email is classified by its **subject line** into one of six categories:


```
+=================================================================+
|           EMAIL CLASSIFICATION & ROUTING                         |
+=================================================================+
|                                                                   |
|  INCOMING EMAIL                                                   |
|       |                                                           |
|       v                                                           |
|  +--- Subject contains "NIS Checklist Print - <number>" ? ---+   |
|  | YES -> _handle_checklist_email (FC Approval + PDF upload)  |   |
|  +------------------------------------------------------------+   |
|       |                                                           |
|       v  NO                                                       |
|  +--- Subject is a reply with "Debit/Credit Notes" ? ---------+  |
|  | YES -> _handle_debit_credit_reply (finance reply PDFs)     |   |
|  +------------------------------------------------------------+   |
|       |                                                           |
|       v  NO                                                       |
|  +--- Subject contains "NIS Approval - <number>" ? -----------+  |
|  | YES -> _handle_hod_approval_email (HOD approval)           |   |
|  +------------------------------------------------------------+   |
|       |                                                           |
|       v  NO                                                       |
|  +--- Subject contains "DSM-NIS-BOOKING" keyword ? -----------+  |
|  | YES -> Main DSM booking flow (prepare + book)              |   |
|  +------------------------------------------------------------+   |
|       |                                                           |
|       v  NO                                                       |
|  [IGNORED - not a relevant email]                                 |
|                                                                   |
+=================================================================+
```

### 2.1 Email Subject Patterns (Complete Reference)

| # | Email Type | Subject Pattern (Regex) | Direction | Handler |
|---|-----------|------------------------|-----------|---------|
| 1 | **DSM Booking Trigger** | Contains `DSM-NIS-BOOKING` (case-insensitive) | Incoming | `_scan_and_prepare` -> `_prepare_run` |
| 2 | **NIS Checklist Print (FC Approval)** | `NIS Checklist Print - <digits>` | Incoming | `_handle_checklist_email` / `_handle_fc_approval_email` |
| 3 | **HOD Approval** | `NIS Approval - <digits>` | Incoming | `_handle_hod_approval_email` |
| 4 | **Debit/Credit Reply** | Contains `Debit/Credit Notes` + reply prefix (Re:/Fw:) | Incoming | `_handle_debit_credit_reply` |
| 5 | **Failure Notification** | `[DSM-NIS Automation FAILED] <SPV> \| <from> to <to>` | Outgoing | `_send_failure_email` |
| 6 | **Completion Notification** | `DSM-NIS [All Stages Complete] \| <SPV> \| <from> to <to>` | Outgoing | `send_final_completion_emails` |
| 7 | **Debit/Credit Request** | `Debit/Credit Notes \| Week <N> \| <from> to <to>` | Outgoing | `_send_one_debit_credit_email` |
| 8 | **Stage Notification (HOD)** | `DSM-NIS [HOD Approval Received] \| <SPV> \| <from> to <to>` | Outgoing | `send_hod_stage_emails` |
| 9 | **Stage Notification (FC)** | `DSM-NIS [FC Approval Received & PDF Merged] \| <SPV> \| ...` | Outgoing | `send_fc_stage_emails` |

### 2.2 Subject Filtering Rules

```python
# Primary filter: DSM_SUBJECT_KEYWORD = "DSM-NIS-BOOKING"
# The keyword must appear ANYWHERE in the subject (case-insensitive)

# MATCHES:
"AUTO-DSM-NIS-Booking"            # embedded keyword
"dsm-nis-booking week 34"         # lowercase
"Re: DSM-NIS-BOOKING summary"     # reply prefix
"FWD: Dsm-Nis-Booking"            # forwarded

# DOES NOT MATCH:
"DSM report"                      # missing "NIS-BOOKING" part
"NIS weekly update"               # missing "DSM-" prefix
"[ALERT] Column Extraction..."    # alert emails are excluded

# EXCLUSIONS:
# Subjects starting with "[ALERT]" or containing "alert:" are ALWAYS skipped
# (prevents re-processing the system's own alert/notification emails)
```

---

## 3. Email Ingestion & Table Extraction


```
+====================+       +=========================+
|   OUTLOOK MAILBOX  |       |  EMAIL PROCESSING       |
|   (Graph API)      |------>|                         |
|                    |       |  1. Fetch last N days   |
| agel.nis.automati  |       |  2. Filter DSM emails   |
| on@adani.com       |       |  3. Skip duplicates     |
+====================+       |     (message_id check)  |
                             |  4. Download attachments |
                             +=========================+
                                        |
                                        v
                    +======================================+
                    |       HTML TABLE PARSING              |
                    | (classified_html_table_parser.py)     |
                    |                                       |
                    |  Extracts TWO tables from email body: |
                    |                                       |
                    |  TABLE 1: METADATA                    |
                    |  +------+----------+---------+-----+  |
                    |  |WeekNo| From Date| To Date |Due  |  |
                    |  |      |          |         |Date |  |
                    |  +------+----------+---------+-----+  |
                    |                                       |
                    |  TABLE 2: SPV DSM DATA                |
                    |  +--------+------+------+-----+----+ |
                    |  |SPV Name|DSM   |Drawl |Rev  |Rev | |
                    |  |        |Chrgs |Chrgs |Diff |Loss| |
                    |  +--------+------+------+-----+----+ |
                    |  |AWEK01L | 1016 |  0   |4517 | 0  | |
                    |  |AWEK03L | -500 |  0   | 200 | 0  | |
                    |  |...     | ...  | ...  | ... |... | |
                    |  +--------+------+------+-----+----+ |
                    +======================================+
```

### Extracted Data Fields:
- **Metadata:** week_no, from_date, to_date, due_date, dsm_statement_published_date
- **SPV Row:** spv_name, total_dsm_charges_payable, drawl_charges_payable, revenue_diff, revenue_loss, net_payable_receivable

### Date Normalization:
All date fields are normalized to `DD-MM-YYYY` format regardless of input format. Supported input formats include:
- `14 Jul 2025`, `14-Jul-25`, `14/Jul/2025` (abbreviated month)
- `14 July 2025`, `July 14 2025` (full month)
- `14/07/2025`, `14-07-2025`, `14.07.2025` (numeric)
- `2025-07-14` (ISO)

---

## 4. Table Types & Row Classification

```
+=================================================================+
|              AMOUNT CLASSIFICATION (per SPV row)                  |
+=================================================================+
|                                                                   |
|  For each of 4 amount fields:                                     |
|    - total_dsm_charges_payable                                    |
|    - drawl_charges_payable                                        |
|    - revenue_diff                                                 |
|    - revenue_loss                                                 |
|                                                                   |
|  +----------+     +---> POSITIVE (+ve) ---> NIS BOOKING           |
|  | Amount   |-----|                                               |
|  | Value    |-----|---> NEGATIVE (-ve) ---> SAP AUTOMATION        |
|  +----------+     +---> ZERO (0)       ---> SKIPPED              |
|                                                                   |
+=================================================================+
```

### Table Types by Row Count:

| Type | Rows | Behaviour | Special Logic |
|------|------|-----------|---------------|
| **1-Row** | 1 SPV | Standard single booking | Normal vendor lookup + booking |
| **2-Row** | 2 SPVs | Same cost center rows combine into ONE NIS | Shared NIS number, combined expense items |
| **3-Row** | 3 SPVs | Per-row invoice pre-extraction, then cost-center grouping | INOX multi-page bill detection |
| **5-Row** | 5 SPVs | Cost-center grouping plus Debit/Credit email to Finance | Gated on ALL rows done before sending |

---

## 5. Table Type Handling Details

### 5.1 -- 1-Row Table (Standard)

```
[1 SPV Row] --> Classify --> NIS Booking (+ve) + SAP Steps (-ve)
                                  |
                                  v
                         One booking, one tracker row
```

- **PDF Upload:** Invoice attachment PDF from email
- **Invoice Number:** Extracted from PDF via LLM (or auto-generated for special vendors)
- **Vendor Data:** Read from DSM-NIS-INPUT.xlsx -> "main" sheet (exact SPV name match)

### 5.2 -- 2-Row Table (Solar/Wind Pair)

```
[AHEJ4L-Solar] --+
                  +--> GROUPED into ONE NIS booking
[AHEJ4L-Wind]  --+        |
                           v
              Combined expense items (each with own cost_center/plant)
              ONE shared NIS checklist number
              ONE shared SharePoint folder: booking/AHEJ4L/
              TWO separate tracker rows (allow_duplicate_week=True)
```

- **Detection:** SPV name matches regex `^\s*([A-Za-z0-9]+)\s*-\s*(Solar|Wind)\b`
- **PDF Upload:** Invoice attachment PDF (shared by both variants)
- **Invoice Number:** From primary (first) row's vendor logic
- **Expense Items:** Each variant's positive amounts become separate line items with own cost_center and plant
- **Folder:** Both variants share one folder named after the base plant code (e.g. `booking/AHEJ4L/`)

### 5.3 -- 3-Row Table (Invoice Pre-Extraction)

```
[SPV-1] --+                    +========================+
[SPV-2] --+-- 3 rows found --> | INVOICE PRE-EXTRACTION |
[SPV-3] --+                    | (invoice_extractor.py) |
                                +========================+
                                         |
                    +--------------------+--------------------+
                    |                                         |
          [INOX bill detected]                    [Non-INOX bill]
          (PDF contains "INOX")                   (single invoice)
                    |                                         |
                    v                                         v
         Per-page extraction:                    One shared invoice
         Match SPV to PDF page                   number for all 3 rows
         -> each row gets unique
           invoice number
```

- **PDF Upload:** Invoice attachment PDF (each row uses same file)
- **Invoice Number:** Pre-resolved before booking; stored in `booking_rows.invoice_number`
- **Booking unit:** Rows are combined by cost center after invoice pre-extraction

### 5.4 -- 5-Row Table (Debit/Credit Trigger)

```
[SPV-1] --+
[SPV-2] --+
[SPV-3] --+-- 5 rows --> NIS + SAP per row (normal booking)
[SPV-4] --+                    |
[SPV-5] --+                    v
                    ALL rows done? --YES--> SEND DEBIT/CREDIT EMAIL
                          |                        |
                          NO                       v
                          |              Finance team reply
                          v              with DN/CN PDFs
                     (wait...)                     |
                                                   v
                                          SAP checklist per reply PDF
                                                   |
                                                   v
                                          Tracker row update (col K)
```

- **Debit/Credit condition:** `len(valid_spv_rows) == config.DEBIT_CREDIT_TRIGGER_ROW_COUNT` (default 5)
- **Gated on:** ALL 5 booking rows must be `done` (NIS+SAP complete) before the D/C email is sent
- **Email content:** Transposed HTML table with DSM Charges, Revenue Diff, Total per SPV + grand total
- **Recipients:** Finance team (from SharePoint sheet `finance_team`, column `emails`)

---

## 6. SPV Exceptions & PDF Rules


```
+=================================================================+
|                  PDF SELECTION RULES                              |
+=================================================================+
|                                                                   |
|  SPV matches NIS_EMAIL_PDF_SPV_KEYWORDS?                          |
|  (AWEK06L, AWEK03L, AWEK05L)                                     |
|         |                                                         |
|    YES  |              NO                                         |
|    |    |              |                                          |
|    v    |              v                                          |
| Use email.pdf    Invoice attachment PDF from email                |
| (email body      available?                                       |
|  converted       |          |                                     |
|  to PDF)         YES        NO                                    |
|    |             |          |                                     |
|    |             v          v                                     |
|    |       Use invoice   Use email.pdf                            |
|    |       attachment    (ultimate fallback)                       |
|    |             |          |                                     |
|    +------+------+----------+                                     |
|           |                                                       |
|           v                                                       |
|   UPLOAD TO NIS PORTAL (header PDF)                               |
|                                                                   |
+=================================================================+
```

### Invoice Number Rules:

| Condition | Invoice Number Source |
|-----------|---------------------|
| AWEK SPVs (NIS_EMAIL_PDF_SPV_KEYWORDS) | Always auto-generated: `cost_center + random_digits` (max 15 chars) |
| Vendor in NIS_AUTO_INVOICE_VENDOR_CODES (215943, 215944, 168856) | Always auto-generated |
| NIS_SKIP_INVOICE_EXTRACTION = True | Always auto-generated |
| 3-row table (pre-extracted) | Pre-resolved from PDF before booking |
| Caller provides invoice_number | Used as-is (highest priority) |
| Normal vendor | Extracted from PDF via LLM (pdfplumber + GPT) |
| LLM extraction fails | Falls back to auto-generated |

### Summary Table:

| SPV / Vendor | PDF Uploaded to NIS | Invoice Number | Source of PDF |
|---|---|---|---|
| AWEK06L, AWEK03L, AWEK05L | `email.pdf` | Auto-generated | Email body -> converted to PDF |
| Vendor 215943, 215944, 168856 | Invoice attachment | Auto-generated | Email attachment PDF |
| Normal vendors (all others) | Invoice attachment | Extracted from PDF (LLM) | Email attachment PDF |
| Any SPV (no PDF found) | `email.pdf` (fallback) | Auto-generated | Email body -> converted to PDF |
| 3-row table (any vendor) | Invoice attachment | Pre-extracted per-row | Email attachment PDF |

---

## 7. NIS Booking Flow (Positive Amounts)

```
+=================================================================+
|                   NIS BOOKING AUTOMATION                          |
|              (nis_booking.py -> book_nis)                          |
+=================================================================+
|                                                                   |
|  INPUT DATA:                                                      |
|  +-- pdf_path         (invoice PDF or email.pdf)                  |
|  +-- company_code     (from vendor master)                        |
|  +-- vendor_code      (from vendor master)                        |
|  +-- bank_key         (from vendor master)                        |
|  +-- cost_center      (from vendor master)                        |
|  +-- plant            (from vendor master)                        |
|  +-- invoice_date     (dsm_statement_published_date or from_date) |
|  +-- invoice_number   (extracted/auto-generated/pre-resolved)     |
|  +-- purpose          "<from_date> TO <to_date> DSM charges"      |
|  +-- expense_items    (for combined cost-center booking)          |
|                                                                   |
|  BROWSER FLOW (Playwright + Edge):                                |
|  1. Launch Edge (headless or headed)                               |
|  2. Login: email -> password -> RSA OTP -> "Stay signed in"        |
|  3. Navigate to NIS Manage URL (direct URL, no tile click)         |
|  4. Wait for app iframe to load                                    |
|  5. Click "Create NIS"                                             |
|  6. Upload header PDF (file chooser)                               |
|  7. Fill Sub Document Type: "Non PO Vendor Payment"               |
|  8. Fill Company Code -> Click Process                             |
|  9. Fill Vendor Code, Bank Key, Invoice Number, Invoice Date      |
|  10. Enable "Dig. Sign Document" switch                            |
|  11. Click Next                                                    |
|  12. Add Expense Items (one per positive amount field):            |
|      +---------------------------------------------+              |
|      | For each positive amount:                   |              |
|      |   - Click "Add Expense Item"                |              |
|      |   - Select Category 6 (Finance/Legal)       |              |
|      |   - Select Sub-category                     |              |
|      |   - Fill Amount                             |              |
|      |   - Fill Cost Center, Plant, WBS            |              |
|      |   - Fill Purpose/Remarks                    |              |
|      |   - Click OK                                |              |
|      +---------------------------------------------+              |
|  13. Click Next -> Approval screen                                 |
|  14. Select Approver radio button -> Submit                        |
|  15. Capture NIS Checklist ID from success popup                  |
|  16. Click OK to dismiss popup                                     |
|                                                                   |
|  OUTPUT:                                                          |
|  +-- NIS Checklist Number (e.g. "7000088592")                    |
|                                                                   |
+=================================================================+
```

### Parallel NIS Booking:
- **NIS_PARALLEL_COUNT** (default: 3) controls how many bookings run simultaneously
- Each worker gets its own browser session, login, and temp directory
- Workers pull from a global pool of pending booking rows across ALL emails
- A login barrier syncs all workers before form filling begins (NIS_LOGIN_BARRIER_TIMEOUT: 180s)
- **NIS_MAX_ROW_RETRIES** (default: 3) limits retries per row before marking as failed

---

## 8. SAP Automation Flow (Negative Amounts)

```
+=================================================================+
|                SAP CHECKLIST AUTOMATION                           |
|           (sap_automation/ -> run_sap_automation)                  |
+=================================================================+
|                                                                   |
|  INPUT DATA (per negative field per row):                         |
|  +-- company_code     (from vendor master via DB)                 |
|  +-- vendor_code      (from vendor master via DB)                 |
|  +-- amount           (absolute value of the negative amount)     |
|  +-- cost_center      (from vendor master via DB)                 |
|  +-- plant            (from vendor master via DB)                 |
|  +-- deb_not_date     (today's date: DD.MM.YYYY)                 |
|  +-- deb_not_ref      (config: "DSM192501")                      |
|                                                                   |
|  SAP GUI FLOW (COM scripting on Windows):                         |
|  1. Launch SAP Logon -> Connect to system                          |
|  2. Open transaction ZDCC                                          |
|  3. Select Non-PO Based -> Create checklist                        |
|  4. Fill company code, vendor, year                                |
|  5. Fill amount, cost center, plant                                |
|  6. Fill debit note date, debit note reference                     |
|  7. Fill document sub-category, tick documents checkbox            |
|  8. Preview -> Save -> confirm popups                              |
|  9. Capture Checklist Number (from status bar)                     |
|  10. Trigger PDF print -> save checklist PDF                       |
|  11. Upload checklist PDF to SharePoint booking folder             |
|  12. Close all SAP windows                                         |
|                                                                   |
|  OUTPUT:                                                          |
|  +-- SAP Checklist Number (e.g. "5001482493")                    |
|  +-- Checklist PDF (uploaded to SP: sap_checklist_*.pdf)         |
|                                                                   |
+=================================================================+
```

### SAP Task Detection:
The SAP service polls the shared DB for pending SAP work using two strategies:
1. **Primary:** Scan `run_steps` table for any `sap_automation_*` step with status `pending`/`running`
2. **Secondary:** Scan `email_runs` with incomplete status and auto-create missing SAP steps from `negative_fields_json`

### Per-Row SAP Steps (5-Row Table):
Step names follow the pattern: `sap_automation_<field>_row_<N>_<spv_safe_name>`
e.g. `sap_automation_revenue_diff_row_3_AWEK01L_50_MW`

---

## 9. HOD / FC Approval Workflow (Post-Booking Stages)


```
+=================================================================+
|          POST-BOOKING APPROVAL & COMPLETION STAGES               |
+=================================================================+
|                                                                   |
|  STAGE 1: NIS Booking Complete                                    |
|    -> NIS checklist number captured                               |
|    -> booking_row.nis_checklist_id set                            |
|                                                                   |
|  STAGE 2: SAP Automation Complete (if row has negatives)          |
|    -> SAP checklist number captured                               |
|    -> booking_row.checklist_value set                             |
|    -> SAP checklist PDF uploaded to SharePoint                    |
|                                                                   |
|  STAGE 3: HOD Approval (auto-approved)                            |
|    -> Triggered: when NIS + SAP both captured                     |
|    -> Sets: hod_approval_status = "done"                          |
|    -> Note: currently auto-approved on row completion             |
|                                                                   |
|  STAGE 4: FC Approval (via email)                                 |
|    -> Triggered: incoming "NIS Checklist Print - <NIS#>" email    |
|    -> The email IS the FC approval event                          |
|    -> Downloads NIS checklist PDF from the email attachment        |
|    -> Uploads PDF to booking folder                               |
|    -> Sets: fc_approval_status = "done"                           |
|    -> Sets: fc_checklist_pdf_uploaded = true                      |
|                                                                   |
|  STAGE 5: PDF Merge                                               |
|    -> Triggered: after FC approval + all checklist PDFs present   |
|    -> Downloads ALL PDFs from booking folder                      |
|    -> Merges into: merged_<SPV>_<to_date>.pdf                    |
|    -> Upload order: NIS checklist > SAP checklist > invoice > email|
|    -> Uploads merged PDF back to booking folder                   |
|                                                                   |
|  STAGE 6: Final Completion Email                                  |
|    -> Triggered: PDF merge done for the row                       |
|    -> Sends email with merged PDF attached                        |
|    -> Recipients: nis_completion_recipients                       |
|                                                                   |
+=================================================================+
```

### Stage State Machine (per booking row):

```
                  NIS Booking
                      |
                      v
              [nis_checklist_id captured]
                      |
         +------------+------------+
         |                         |
    Has negatives?            No negatives
         |                         |
         v                         v
    SAP Automation          Row = "done"
         |                         |
         v                         |
    [checklist_value]              |
         |                         |
         +----------+--------------+
                    |
                    v
         HOD Auto-Approval
         (hod_approval_status = done)
                    |
                    v
         [WAIT for FC email]
         "NIS Checklist Print - <NIS#>"
                    |
                    v
         FC Approval + PDF Upload
         (fc_approval_status = done)
                    |
                    v
         PDF Merge
         (all PDFs in folder -> merged_*.pdf)
                    |
                    v
         Final Completion Email
         (merged PDF attached)
```

---

## 10. Debit/Credit Email Flow (5-Row Table)

```
+=================================================================+
|           DEBIT/CREDIT NOTE WORKFLOW (5-ROW TRIGGER)              |
+=================================================================+
|                                                                   |
|  TRIGGER: All 5 booking rows = "done" AND all SAP steps done     |
|                                                                   |
|  SENT TO: finance_team recipients (from SharePoint sheet)         |
|                                                                   |
|  EMAIL CONTENT:                                                   |
|  +-- Subject: "Debit/Credit Notes | Week 04 | 20-Apr to 26-Apr"  |
|  +-- Body: "Dear Team, Please provide Debit/Credit Notes"         |
|  +-- DSM Amount: total net payable (Indian format: 21,63,738.00)  |
|  +-- Company Code (from config)                                   |
|  +-- Due Date CN/DN                                               |
|  +-- Transposed table:                                            |
|      +----------+-------+-------+-------+-------+-------+------+ |
|      |          | SPV-1 | SPV-2 | SPV-3 | SPV-4 | SPV-5 |Total | |
|      +----------+-------+-------+-------+-------+-------+------+ |
|      |DSM Chrgs | 1,016 |-5,000 | 200   | 3,000 | -100  | ...  | |
|      |Rev Diff  | 4,517 | 0     | 1,200 | -800  | 500   | ...  | |
|      |Total     | 5,533 |-5,000 | 1,400 | 2,200 | 400   | ...  | |
|      +----------+-------+-------+-------+-------+-------+------+ |
|                                                                   |
+=================================================================+
                              |
                              v
+=================================================================+
|           FINANCE TEAM REPLY HANDLING                             |
+=================================================================+
|                                                                   |
|  DETECTION:                                                       |
|    - Subject contains "Debit/Credit Notes"                        |
|    - AND has reply prefix (Re:/Fw:/Fwd:/Aw:/Wg:)                 |
|    - Matched to sent record by normalized subject tokens          |
|                                                                   |
|  REPLY CONTAINS: Debit/Credit Note PDFs attached                  |
|  (e.g. "1.AWEK06L.pdf", "2.AWEK3L.pdf", "3.AWEK5L.pdf")         |
|                                                                   |
|  PER-PDF EXTRACTION (debit_credit_pdf_extractor.py):              |
|  +-- SPV name (from filename: "1.AWEK06L.pdf" -> "AWEK06L")      |
|  +-- vendor_code (matched from DSM-NIS-INPUT.xlsx by SPV)         |
|  +-- company_code, cost_center, plant                             |
|  +-- ref_number (Debit/Credit Note reference from PDF text)       |
|  +-- amount (extracted via LLM or keyword parser)                 |
|  +-- doc_date (document date)                                     |
|  +-- note_type (Debit Note / Credit Note)                         |
|                                                                   |
|  STORAGE: Uploaded to booking/<from>_<to>/Debit-Credit-Notes/     |
|                                                                   |
|  THEN: SAP automation creates a checklist for EACH reply PDF      |
|  (same SAP flow: fill form -> save -> print -> capture number)    |
|                                                                   |
|  TRACKER UPDATE: add_sap_checklist_reply_row()                    |
|  -> New row on SPV sheet with NIS column (I) BLANK                |
|  -> SAP checklist number in column K                              |
|  -> CN/DN Date in column L                                        |
|                                                                   |
+=================================================================+
```

---

## 11. Failure Handling & Notifications

```
+=================================================================+
|           FAILURE EMAIL NOTIFICATION                              |
+=================================================================+
|                                                                   |
|  TRIGGER: A booking row fails after MAX retries (3)               |
|                                                                   |
|  SENT TO: nis_completion_recipients (SharePoint sheet)            |
|                                                                   |
|  Subject: "[DSM-NIS Automation FAILED] <SPV> | <from> to <to>"   |
|                                                                   |
|  BODY CONTAINS:                                                   |
|  1. LLM-generated plain-language explanation of failure           |
|     (which value is missing or mismatched)                        |
|  2. Per-row status table:                                         |
|     +-----+----------+--------+-------------------+----------+   |
|     | Row | SPV      | Status | Missing value(s)  | Details  |   |
|     +-----+----------+--------+-------------------+----------+   |
|     | 0   | AWEK01L  | DONE   |                   |          |   |
|     | 1   | AWEK03L  | FAILED | Vendor master data| no match |   |
|     +-----+----------+--------+-------------------+----------+   |
|  3. Run ID for reference                                          |
|                                                                   |
|  LIMIT: ONE failure email per email/run (idempotent)              |
|                                                                   |
+=================================================================+
```

### Required Fields (checked before booking):
- `vendor_data_json` (vendor code, company code, cost center, plant)
- `invoice_date`
- `booking_folder_path`
- `positive_fields_json` (charge classification)
- `metadata_json` (week, dates, due date)
- `spv_data_json` (SPV charge values)

---

## 12. SharePoint Data Sources & Folder Structure


```
+=================================================================+
|              SHAREPOINT FILE USAGE                                |
+=================================================================+
|                                                                   |
|  DSM-NIS-INPUT.xlsx (read_vendor_master_data.py)                  |
|  +-- Sheet: "main"                                                |
|  |   Columns: SPV Name | Vendor | Company | Cost Center |         |
|  |            Plant | Bank Key | GL                                |
|  |                                                                |
|  |   USED FOR: Vendor master data lookup per SPV name             |
|  |   MATCHED BY: SPV name from email table -> exact/fuzzy match   |
|  |                                                                |
|  DSM-NIS-OUTPUT.xlsx (add_new_row_data_nis.py)                    |
|  +-- Sheet: <SPV Name> (one sheet per SPV, auto-created)          |
|  |   Columns A-L:                                                 |
|  |     A: Week No                                                 |
|  |     B: Week (from_date to to_date)                             |
|  |     C: Statement date                                          |
|  |     D: Working Received Date                                   |
|  |     E: DSM Charges                                             |
|  |     F: Revenue Diff amount                                     |
|  |     G: Drawl Charges Payable                                   |
|  |     H: Total Amount payable/Receivable (in Rs)                 |
|  |     I: NIS No (checklist number)                               |
|  |     J: Due Date                                                |
|  |     K: CN/DN Checklist No (SAP number)                         |
|  |     L: CN/DN Date                                              |
|  |                                                                |
|  |   WRITE MODES:                                                 |
|  |     - New row: after NIS booking (col I filled)                |
|  |     - Merge onto existing row: SAP adds col K to existing      |
|  |     - Reply row: debit/credit SAP (col I blank, col K filled)  |
|  |                                                                |
|  nis_completion_recipients.xlsx                                    |
|  +-- Sheet: "nis_completion_recipients"                           |
|  |   Column: "emails"                                             |
|  |   USED FOR: Failure/completion/stage email recipients          |
|  |                                                                |
|  automation_sheet_db.xlsx                                          |
|  +-- Sheet: "finance_team"                                        |
|  |   Column: "emails"                                             |
|  |   USED FOR: Debit/Credit email recipients                     |
|  |                                                                |
+=================================================================+
```

### Booking Folder Structure:
```
Bots/AGEL142 DSM NIS Booking Automation/booking/
    <SPV Name>/
        <from_date>_<to_date>/
            email.pdf                    (email body converted to PDF)
            DSM_<attachment>.pdf         (invoice PDF from email)
            sap_checklist_<number>.pdf   (SAP automation output)
            nis_checklist_<number>.pdf   (from FC approval email)
            merged_<SPV>_<to_date>.pdf   (final merged document)
            Debit-Credit-Notes/          (sub-folder)
                1.AWEK06L.pdf            (finance reply PDFs)
                2.AWEK3L.pdf
                3.AWEK5L.pdf
```

---

## 13. Row Completion Logic

```
+=================================================================+
|                ROW COMPLETION STATE MACHINE                       |
+=================================================================+
|                                                                   |
|  Row has ONLY +ve amounts (no -ve):                               |
|    NIS done --> row = "done" immediately                          |
|                                                                   |
|  Row has BOTH +ve AND -ve amounts:                                |
|    NIS done --> row = "nis_done" (partial)                        |
|              --> SAP automation runs independently                 |
|              --> ALL SAP steps done --> row = "done"              |
|                                                                   |
|  Row has ONLY -ve amounts (no +ve):                               |
|    NIS skipped --> row = "nis_done" (partial)                     |
|                 --> SAP automation runs independently              |
|                 --> ALL SAP steps done --> row = "done"           |
|                                                                   |
|  Row has NEITHER +ve nor -ve (all zero):                          |
|    row = "skipped"                                                |
|                                                                   |
+=================================================================+
```

### Status Values:
| Status | Meaning |
|--------|---------|
| `pending` | Waiting to be processed |
| `running` | Currently being processed |
| `nis_done` | NIS booking complete, awaiting SAP automation |
| `done` | Fully complete (NIS + SAP both finished) |
| `failed` | Error occurred (after retries exhausted) |
| `skipped` | No work needed (all amounts zero) |

---

## 14. End-to-End Workflow (Complete Pipeline)

```
+-----------------+
|  DSM EMAIL      |
|  RECEIVED       |
+--------+--------+
         |
         v
+-----------------+     +--------------------------+
|  PARSE HTML     |---->|  metadata_table           |
|  TABLES         |     |  spv_dsm_table (N rows)   |
+--------+--------+     +--------------------------+
         |
         v
+-----------------+
|  CLASSIFY       |  Per row: +ve -> NIS, -ve -> SAP, 0 -> skip
|  AMOUNTS        |
+--------+--------+
         |
         v
+-----------------+     +--------------------------+
|  SAVE TO        |---->|  SharePoint booking folder |
|  SHAREPOINT     |     |  +-- email.pdf             |
|                 |     |  +-- DSM_<attachment>.pdf   |
+--------+--------+     +--------------------------+
         |
         v
+-----------------+     +--------------------------+
|  READ VENDOR    |<----|  DSM-NIS-INPUT.xlsx        |
|  MASTER DATA    |     |  Sheet: "main"             |
|  (per SPV)      |     |  Match: SPV name -> vendor |
+--------+--------+     +--------------------------+
         |
         +------------------------------+
         v                              v
+-----------------+          +---------------------+
|  NIS BOOKING    |          |  SAP AUTOMATION      |
|  (+ve amounts)  |          |  (-ve amounts)       |
|  [Parallel x3]  |          |  [Independent svc]   |
|                 |          |                      |
|  OUT: NIS #     |          |  OUT: SAP # + PDF    |
+--------+--------+          +----------+-----------+
         |                              |
         v                              v
+-------------------------------------------------+
|          ROW MARKED "DONE"                       |
|  (only when BOTH NIS + SAP complete for row)     |
+------------------------+------------------------+
                         |
                         v
+-------------------------------------------------+
|  UPDATE DSM-NIS-OUTPUT.xlsx TRACKER              |
|  (NIS No in col I, SAP No in col K)             |
+------------------------+------------------------+
                         |
                         v
              +--- 5-row table? ---+
              |                    |
             YES                  NO
              |                    |
              v                    |
+---------------------+            |
| SEND DEBIT/CREDIT   |            |
| EMAIL TO FINANCE    |            |
+----------+----------+            |
           |                       |
           v                       |
+---------------------+            |
| FINANCE REPLIES     |            |
| WITH DN/CN PDFs     |            |
+----------+----------+            |
           |                       |
           v                       |
+---------------------+            |
| SAP AUTOMATION      |            |
| FOR EACH REPLY PDF  |            |
+----------+----------+            |
           |                       |
           +----------+------------+
                      |
                      v
+-------------------------------------------------+
|  HOD AUTO-APPROVAL                               |
|  (triggered when NIS+SAP both captured)          |
+------------------------+------------------------+
                         |
                         v
+-------------------------------------------------+
|  FC APPROVAL (via incoming email)                |
|  "NIS Checklist Print - <NIS#>"                  |
|  -> Download NIS checklist PDF                   |
|  -> Upload to booking folder                     |
+------------------------+------------------------+
                         |
                         v
+-------------------------------------------------+
|  PDF MERGE                                       |
|  All PDFs in folder -> merged_<SPV>_<date>.pdf   |
+------------------------+------------------------+
                         |
                         v
+-------------------------------------------------+
|  FINAL COMPLETION EMAIL                          |
|  (merged PDF attached, sent to recipients)       |
+-------------------------------------------------+
```

---

## 15. Database Schema


| Table | Purpose | Key Columns |
|-------|---------|-------------|
| `email_runs` | One record per email (keyed by message_id) | id, message_id, subject, status, metadata_json, spv_dsm_json, vendor_data_json, positive/negative_fields_json, nis_checklist_id, checklist_value, booking_folder_path |
| `booking_rows` | One record per SPV row (UNIQUE: message_id + row_index) | row_id, run_id, row_index, spv_name, status, vendor_data_json, invoice_date, invoice_number, nis_checklist_id, checklist_value, hod/fc_approval_status, pdf_merge_status, merged_pdf_name |
| `run_steps` | Step tracking: extraction, classification, NIS, SAP per field per row | id, run_id, step_name, status, detail, error, retry_count |
| `debit_credit_emails` | Tracks sent debit/credit emails and reply matching | id, run_id, message_id, spv_name, week_no, from/to_date, email_status, sent_subject, reply_message_id, reply_pdfs_uploaded |
| `debit_credit_pdfs` | One record per finance-reply PDF -> SAP checklist | id, dc_email_id, pdf_name, spv_from_filename, vendor_code, ref_number, amount, doc_date, sap_status, sap_checklist_number |

### Backend Support:
- **PostgreSQL** (default, production): `DB_BACKEND=postgres` -- Azure hosted at `az10psqldmrcbtp01.postgres.database.azure.com`
- **SQLite** (local development): `DB_BACKEND=sqlite` -- file at `nis_booking/nis_booking.db`

---

## 16. Test Cases & Email Scenarios

### 16.1 DSM Booking Email Test Cases

| Test Case | Email Subject | SPV Rows | Expected Behaviour |
|-----------|--------------|----------|-------------------|
| TC-01: Standard 1-row | `Fw: DSM-NIS-BOOKING` | 1 SPV (e.g. AWEMP1PL) | Single NIS booking + SAP for negatives |
| TC-02: Solar/Wind pair | `DSM-NIS-BOOKING Week 08` | 2 SPVs: AHEJ4L-Solar, AHEJ4L-Wind | Combined into ONE NIS booking, shared folder |
| TC-03: 3-row table | `AUTO-DSM-NIS-Booking` | 3 SPVs (with invoice PDF) | Pre-extract invoice per page, book each independently |
| TC-04: 3-row INOX | `Re: DSM-NIS-BOOKING` | 3 SPVs (INOX keyword in PDF) | Per-page invoice extraction matching SPV to page |
| TC-05: 5-row table | `dsm-nis-booking week 34` | 5 SPVs | Book all rows, then send Debit/Credit email |
| TC-06: All positive | `FWD: DSM-NIS-Booking` | Any (all amounts > 0) | NIS only, row done immediately, no SAP |
| TC-07: All negative | `DSM-NIS-BOOKING` | Any (all amounts < 0) | NIS skipped, SAP only, row done after SAP |
| TC-08: Mixed amounts | `DSM-NIS-BOOKING` | Any (some +ve, some -ve) | NIS for +ve, SAP for -ve, done when both finish |
| TC-09: All zero | `DSM-NIS-BOOKING` | Any (all amounts = 0) | Row skipped immediately |
| TC-10: AWEK SPV | `DSM-NIS-BOOKING` | AWEK06L, AWEK03L, or AWEK05L | Uses email.pdf, auto-generates invoice |
| TC-11: Duplicate email | `DSM-NIS-BOOKING` | Same message_id as previous | Skipped (already registered) |
| TC-12: Missing tables | `DSM-NIS-BOOKING` | No valid HTML tables | Extraction error -> failure email |
| TC-13: Missing vendor | `DSM-NIS-BOOKING` | SPV not in DSM-NIS-INPUT | Vendor lookup fails -> failure email |

### 16.2 Approval Email Test Cases

| Test Case | Email Subject | Expected Behaviour |
|-----------|--------------|-------------------|
| TC-20: HOD Approval | `NIS Approval - 7000088592` | Mark matching rows as HOD approved, forward to recipients |
| TC-21: FC Approval (NIS Checklist) | `NIS Checklist Print - 7000088592` | Download PDF, upload to folder, mark FC approved, trigger PDF merge |
| TC-22: HOD before SAP done | `NIS Approval - 7000088592` | Skipped if row status != "done" |
| TC-23: FC before HOD | `NIS Checklist Print - 7000088592` | Skipped (requires HOD approval first) |
| TC-24: Duplicate HOD | `NIS Approval - 7000088592` | Idempotent -- already approved rows skipped |
| TC-25: Unknown NIS number | `NIS Approval - 9999999999` | No matching row found -> skipped |

### 16.3 Debit/Credit Reply Test Cases

| Test Case | Email Subject | Expected Behaviour |
|-----------|--------------|-------------------|
| TC-30: Finance reply | `Re: Debit/Credit Notes \| Week 04 \| 20-Apr to 26-Apr` | Download PDFs, extract data, store in DB |
| TC-31: Reply with PDFs | `Fw: Debit/Credit Notes \| Week 04 \| ...` | Upload to booking/Debit-Credit-Notes/, extract per-PDF |
| TC-32: Reply no PDFs | `Re: Debit/Credit Notes \| Week 04 \| ...` | Skipped (no PDF attachments yet) |
| TC-33: SAP for reply PDF | (internal -- triggered after TC-30) | SAP checklist created per reply PDF with extracted values |
| TC-34: Duplicate reply | Same reply already processed | Idempotent -- reply_pdfs_uploaded already set |
| TC-35: Not a reply | `Debit/Credit Notes \| Week 04 \| ...` (no Re:/Fw:) | Ignored (our own outgoing email) |

### 16.4 Failure & Edge Case Test Cases

| Test Case | Scenario | Expected Behaviour |
|-----------|----------|-------------------|
| TC-40: Extraction failure | Email body has no HTML tables | Failure email sent with "extraction failed" |
| TC-41: Validation failure | Tables present but missing mandatory columns | Failure email with column validation details |
| TC-42: NIS portal timeout | Browser takes too long to load | Retried up to NIS_MAX_ROW_RETRIES times |
| TC-43: SAP COM error | SAP GUI not responding | Step retried up to MAX_STEP_RETRIES (3) times |
| TC-44: Crash resume | Service killed mid-booking | On restart, picks up from first incomplete step |
| TC-45: Parallel conflict | Two workers try same row | DB locks prevent double-booking |
| TC-46: Alert email | `[ALERT] Column Extraction Failed` | Excluded by subject filter (never processed) |
| TC-47: Deleted email | Email in Deleted Items folder | Filtered out by parentFolderId check |
| TC-48: Non-inbox email | Email in Sent Items | Not fetched (only Inbox queried) |

---

## 17. Configuration Reference

### Key Environment Variables:

| Variable | Default | Purpose |
|----------|---------|---------|
| `DSM_SUBJECT_KEYWORD` | `DSM-NIS-BOOKING` | Email subject filter keyword |
| `NIS_PARALLEL_COUNT` | `3` | Concurrent NIS booking workers |
| `NIS_MAX_ROW_RETRIES` | `3` | Max retries per booking row |
| `DB_BACKEND` | `postgres` | Database backend (postgres/sqlite) |
| `DEBIT_CREDIT_ENABLED` | `true` | Enable debit/credit flow |
| `DEBIT_CREDIT_TRIGGER_ROW_COUNT` | `5` | Number of SPV rows to trigger D/C |
| `DEBIT_CREDIT_SAP_ENABLED` | `true` | Enable SAP for D/C reply PDFs |
| `NIS_SKIP_INVOICE_EXTRACTION` | `false` | Skip PDF invoice extraction |
| `NIS_EMAIL_PDF_SPV_KEYWORDS` | `AWEK06L,AWEK03L,AWEK05L` | SPVs that use email.pdf |
| `NIS_AUTO_INVOICE_VENDOR_CODES` | `215943,215944,168856` | Vendors with auto-generated invoice |
| `LLM_USE_OPENAI_DIRECT` | `false` | Use public OpenAI API instead of Azure |
| `FAILURE_EMAIL_ENABLED` | `true` | Send failure notification emails |

---

## 18. File & Module Map

| File | Role |
|------|------|
| `run_nis_booking.py` | Entry point: NIS service (sets sys.path, calls nis_main) |
| `run_sap_automation.py` | Entry point: SAP service (sets sys.path, calls sap_main) |
| `nis_booking/main.py` | NIS orchestrator: email scan -> classify -> prepare -> parallel book -> debit/credit -> approvals |
| `nis_booking/nis_booking.py` | Browser automation: NIS portal form filling (Playwright + Edge) |
| `nis_booking/db_manager.py` | Database layer -- PostgreSQL/SQLite (shared by both services) |
| `nis_booking/config.py` | Single source of truth for all credentials, URLs, and settings |
| `nis_booking/read_email.py` | Fetch emails from Outlook via Microsoft Graph API |
| `nis_booking/classified_html_table_parser.py` | Email HTML -> structured metadata + SPV tables |
| `nis_booking/read_vendor_master_data.py` | SharePoint Excel -> vendor data (exact/fuzzy SPV match) |
| `nis_booking/invoice_extractor.py` | PDF -> invoice number (pdfplumber + LLM) |
| `nis_booking/add_new_row_data_nis.py` | Write tracker rows to DSM-NIS-OUTPUT.xlsx |
| `nis_booking/debit_credit_email.py` | Build debit/credit email HTML (5-row trigger) |
| `nis_booking/debit_credit_pdf_extractor.py` | Extract ref/amount/date from finance reply PDFs |
| `nis_booking/sharepoint_upload.py` | Upload/download files to/from SharePoint |
| `nis_booking/pdf_merger.py` | Merge multiple PDFs into one (for final completion) |
| `nis_booking/recipients_sp.py` | NIS completion/failure email recipients (SharePoint) |
| `nis_booking/finance_recipients_sp.py` | Finance team recipients for D/C emails (SharePoint) |
| `nis_booking/log_manager.py` | Thread-local file-based logging per SPV per date range |
| `nis_booking/llm_curl_client.py` | Azure OpenAI / curl-based LLM client for failure explanations |
| `sap_automation/sap_main.py` | SAP orchestrator: poll DB -> process pending SAP tasks |
| `sap_automation/sap_automation.py` | SAP GUI automation: form fill -> save -> print (COM scripting) |
| `sap_automation/checklist_capture_sap.py` | Capture SAP checklist number from status bar |
| `sap_automation/sap_config.py` | SAP-specific configuration (system, client, credentials) |

---

## 19. External Systems

```
+─────────────────────────────────────────────────────────────────+
│                    Microsoft Cloud (Azure AD)                     │
│                                                                   │
│  ┌───────────────────┐        ┌──────────────────────────────┐   │
│  │  Outlook Mailbox  │        │  SharePoint / OneDrive       │   │
│  │  (Graph API)      │        │  AGEL-Automation site        │   │
│  │                   │        │  - DSM-NIS-INPUT.xlsx        │   │
│  │  agel.nis.        │        │  - DSM-NIS-OUTPUT.xlsx       │   │
│  │  automation@      │        │  - booking/<SPV>/<dates>/    │   │
│  │  adani.com        │        │  - recipients sheets         │   │
│  └───────────────────┘        └──────────────────────────────┘   │
│                                                                   │
│  ┌───────────────────┐        ┌──────────────────────────────┐   │
│  │  Azure PostgreSQL │        │  Azure OpenAI                │   │
│  │  (shared DB)      │        │  (invoice extraction +       │   │
│  │  dsm_nis_booking  │        │   failure explanations)      │   │
│  └───────────────────┘        └──────────────────────────────┘   │
+─────────────────────────────────────────────────────────────────+

┌──────────────────────────┐    ┌──────────────────────────────────┐
│  NIS Fiori Launchpad     │    │  SAP GUI (Windows Desktop)       │
│  (SAP BTP / HANA Cloud)  │    │  Transaction: ZDCC               │
│  ael-nis-prd.launchpad   │    │  Non-PO Vendor Payment           │
│  .cfapps.in30.hana...    │    │  Checklist PDF printed locally   │
└──────────────────────────┘    └──────────────────────────────────┘
```

---

## 20. Polling & Execution Model

### NIS Booking Service (`run_nis_booking.py`):
```
┌─── ONE-SHOT with idle timeout ───────────────────────────────┐
│                                                               │
│  while True:                                                  │
│    1. _scan_and_prepare()      # fetch emails, seed rows      │
│    2. _book_pending_rows()     # parallel NIS booking         │
│    3. send_debit_credit()      # gated D/C emails             │
│    4. process_approvals()      # HOD/FC/merge/completion      │
│                                                               │
│    if no_work_found:                                          │
│      idle_scans++                                             │
│      if idle_scans >= 20: EXIT                                │
│      sleep(30s)                                               │
│    else:                                                      │
│      idle_scans = 0                                           │
│                                                               │
└───────────────────────────────────────────────────────────────┘
```

### SAP Automation Service (`run_sap_automation.py`):
```
┌─── One-shot OR Poll mode ─────────────────────────────────────┐
│                                                               │
│  One-shot (default):                                          │
│    process_all_pending_sap_tasks()  # then exit               │
│                                                               │
│  Poll mode (--poll --interval 60):                            │
│    while True:                                                │
│      process_all_pending_sap_tasks()                          │
│      process_debit_credit_sap_checklists()                    │
│      sleep(interval)                                          │
│                                                               │
└───────────────────────────────────────────────────────────────┘
```
