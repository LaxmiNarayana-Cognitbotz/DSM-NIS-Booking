# DSM-NIS Booking Automation Pipeline

End-to-end automation for **AGEL (Adani Green Energy Limited)** that monitors a dedicated Outlook mailbox for DSM (Deviation Settlement Mechanism) emails, classifies SPV financial fields, and routes each transaction to either the **NIS portal** (positive amounts via Playwright) or **SAP GUI** (negative amounts via COM scripting). Results are written to a SharePoint Excel tracker, followed by a multi-stage approval workflow (HOD/FC), PDF merging, and completion notifications.

---

## Table of Contents

- [Overview](#overview)
- [Architecture](#architecture)
- [Email Types & Triggers](#email-types--triggers)
- [Pipeline Steps](#pipeline-steps)
- [Project Structure](#project-structure)
- [Module Reference](#module-reference)
- [Configuration](#configuration)
- [Installation](#installation)
- [Running the Pipeline](#running-the-pipeline)
- [Test Cases](#test-cases)
- [Resume & Fault Tolerance](#resume--fault-tolerance)

---

## Overview

```
Outlook Email (DSM-NIS-BOOKING)
        |
        v
  Read & Parse Email (Graph API)
        |
        v
  Classify SPV Amounts (per row, per field)
   +--------+---------+
   |        |         |
 +ve amt  -ve amt   zero
   |        |         |
   v        v         v
NIS Portal  SAP GUI  SKIP
(Playwright)(COM)
   |        |
   +---+----+
       |
       v
 SharePoint Excel Tracker (DSM-NIS-OUTPUT.xlsx)
       |
       v
 HOD Auto-Approval -> FC Approval (email) -> PDF Merge -> Completion Email
```

The pipeline is **resume-aware**: every step is recorded in a shared PostgreSQL/SQLite database. If either service crashes, restarting automatically picks up from the first incomplete step.

---

## Architecture

### Two Independent Services

| Service | Entry Point | Host | Responsibility |
|---------|-------------|------|----------------|
| **NIS Booking** | `run_nis_booking.py` | Any (Playwright) | Email scan, extraction, NIS booking (+ve), tracker update, debit/credit, approvals |
| **SAP Automation** | `run_sap_automation.py` | Windows VM (SAP GUI) | SAP checklist creation (-ve), debit/credit reply SAP checklists |

Both share ONE database and never call each other directly.


```
+---------------------------------------------------------------------+
|                          main.py  (NIS Orchestrator)                  |
|                                                                       |
|  +------------+  +---------------+  +-------------+  +-------------+ |
|  |read_email  |  |classified_    |  |read_vendor_ |  |add_new_row_ | |
|  |   .py      |  |html_table_   |  |master_data  |  |data_nis.py  | |
|  |            |  |parser.py      |  |.py          |  |(SP write)   | |
|  +-----+------+  +------+--------+  +------+------+  +------+------+ |
|        |                |                   |                |        |
|  +-----v----------------v-------------------v----------------v------+ |
|  |              db_manager.py  (PostgreSQL / SQLite)                 | |
|  |    email_runs | booking_rows | run_steps | debit_credit_*        | |
|  +------------------------------------------------------------------+ |
|                                                                       |
|  +---------------------------+   +----------------------------------+ |
|  |     nis_booking.py        |   |  debit_credit_email.py           | |
|  |  (Playwright / Edge)      |   |  (5-row finance email trigger)   | |
|  |  Login + Form + Submit    |   +----------------------------------+ |
|  +---------------------------+                                        |
+-----------------------------------------------------------------------+

+-----------------------------------------------------------------------+
|                    sap_main.py  (SAP Orchestrator)                     |
|                                                                       |
|  Polls shared DB for pending SAP tasks                                |
|  +---------------------------+   +----------------------------------+ |
|  | sap_automation.py         |   | checklist_capture_sap.py         | |
|  | (SAP GUI COM scripting)   |   | (capture number from status bar) | |
|  +---------------------------+   +----------------------------------+ |
+-----------------------------------------------------------------------+
```

### External Systems

| System | Protocol | Purpose |
|--------|----------|---------|
| Outlook Mailbox | Graph API (MSAL) | Email ingestion + sending |
| SharePoint / OneDrive | Graph API | Vendor master, tracker Excel, booking folders |
| Azure PostgreSQL | psycopg2 | Shared pipeline state database |
| Azure OpenAI | REST API | Invoice extraction + failure explanations |
| NIS Fiori Launchpad | Playwright (Edge) | NIS portal form automation |
| SAP GUI (Windows) | win32com COM | SAP checklist creation |

---

## Email Types & Triggers

The system monitors `agel.nis.automation@adani.com` (Inbox only). Each email is classified by subject:

### Incoming Emails (Triggers)

| # | Type | Subject Pattern | Action |
|---|------|----------------|--------|
| 1 | **DSM Booking** | Contains `DSM-NIS-BOOKING` (case-insensitive) | Parse tables, classify amounts, book NIS/SAP |
| 2 | **FC Approval** | `NIS Checklist Print - <number>` | Download PDF, mark FC approved, trigger merge |
| 3 | **HOD Approval** | `NIS Approval - <number>` | Mark HOD approved for matching rows |
| 4 | **Debit/Credit Reply** | Contains `Debit/Credit Notes` + reply prefix (Re:/Fw:) | Download PDFs, extract data, trigger SAP |

### Outgoing Emails (Notifications)

| # | Type | Subject Pattern | When Sent |
|---|------|----------------|-----------|
| 5 | **Failure** | `[DSM-NIS Automation FAILED] <SPV> \| <from> to <to>` | Booking row fails after max retries |
| 6 | **Completion** | `DSM-NIS [All Stages Complete] \| <SPV> \| <dates>` | All stages done (NIS+SAP+HOD+FC+merge) |
| 7 | **Debit/Credit Request** | `Debit/Credit Notes \| Week <N> \| <from> to <to>` | 5-row table fully booked |
| 8 | **HOD Stage** | `DSM-NIS [HOD Approval Received] \| <SPV> \| <dates>` | HOD approval processed |
| 9 | **FC Stage** | `DSM-NIS [FC Approval Received & PDF Merged] \| ...` | FC approved + PDF merged |

### Subject Filter Rules
- Keyword `DSM-NIS-BOOKING` must appear anywhere in subject (case-insensitive)
- Subjects with `[ALERT]` or `alert:` are always excluded
- Only Inbox folder is queried (Sent Items, Deleted Items excluded)
- Duplicate emails are skipped (message_id already registered in DB)

---

## Pipeline Steps

Each step is tracked in the database with status: `pending -> running -> done / failed / skipped`.


### Phase 1: Email Extraction & Classification

| Step | Name | What It Does |
|------|------|-------------|
| 1 | `read_email` | Fetch DSM emails via Graph API, download PDF attachments |
| 2 | `email_extraction` | Parse two HTML tables: metadata (week/dates) + SPV DSM data (amounts) |
| 3 | `email_classification` | Classify each amount: +ve -> NIS, -ve -> SAP, 0 -> skip |
| 4 | `email_validation` | Confirm required fields present, at least one actionable amount |
| 5 | `read_vendor_master` | Read vendor data from SharePoint Excel (DSM-NIS-INPUT.xlsx) |

### Phase 2: NIS Booking (Positive Amounts)

| Step | Name | What It Does |
|------|------|-------------|
| 6 | `nis_booking_<field>` | Playwright automation: login -> upload PDF -> fill form -> add expense items -> submit -> capture NIS number |

**Parallel execution:** Up to `NIS_PARALLEL_COUNT` (default 3) bookings run simultaneously, each with its own browser session.

### Phase 3: SAP Automation (Negative Amounts)

| Step | Name | What It Does |
|------|------|-------------|
| 7 | `sap_automation_<field>` | SAP GUI COM: open ZDCC -> fill form -> save -> print PDF -> capture checklist number -> upload to SharePoint |

Runs on a separate Windows VM. Polls the shared DB for pending tasks.

### Phase 4: Post-Booking Stages

| Step | Trigger | What It Does |
|------|---------|-------------|
| 8 | Row done | Update DSM-NIS-OUTPUT.xlsx tracker (NIS No in col I, SAP No in col K) |
| 9 | NIS+SAP captured | HOD auto-approval |
| 10 | "NIS Checklist Print" email | FC approval + NIS checklist PDF upload |
| 11 | FC approved | PDF merge (all docs -> one merged PDF) |
| 12 | Merge done | Final completion email with merged PDF attached |
| 13 | 5-row table done | Debit/Credit email to finance team |

---

## Project Structure

```
DSM-ongoing/
|
+-- run_nis_booking.py               # Entry point: NIS service
+-- run_sap_automation.py            # Entry point: SAP service
+-- requirements.txt                 # Python dependencies
+-- .env.example                     # Environment variable template
|
+-- nis_booking/                     # NIS Booking service package
|   +-- main.py                      # Orchestrator (email scan -> book -> approve)
|   +-- config.py                    # All credentials, URLs, settings
|   +-- db_manager.py               # Database layer (PostgreSQL/SQLite)
|   +-- read_email.py               # Fetch emails via Graph API
|   +-- classified_html_table_parser.py  # Parse email HTML tables
|   +-- nis_booking.py              # Playwright browser automation
|   +-- read_vendor_master_data.py  # Vendor data from SharePoint Excel
|   +-- invoice_extractor.py        # PDF -> invoice number (LLM)
|   +-- add_new_row_data_nis.py     # Write to DSM-NIS-OUTPUT.xlsx
|   +-- sharepoint_upload.py        # Upload/download SharePoint files
|   +-- debit_credit_email.py       # Build 5-row D/C email HTML
|   +-- debit_credit_pdf_extractor.py  # Extract from finance reply PDFs
|   +-- pdf_merger.py               # Merge multiple PDFs into one
|   +-- recipients_sp.py            # NIS completion email recipients
|   +-- finance_recipients_sp.py    # Finance team recipients (D/C)
|   +-- log_manager.py              # Per-SPV file-based logging
|   +-- llm_curl_client.py          # Azure OpenAI client
|   +-- loginwithpassword.py        # Standalone NIS login (dev/test)
|   +-- nis_booking.db              # Local SQLite DB (if not using PG)
|   +-- temp-book/                  # Temporary working directory
|   +-- logs/                       # Per-SPV per-period log files
|
+-- sap_automation/                  # SAP Automation service package
|   +-- sap_main.py                 # SAP orchestrator (poll + process)
|   +-- sap_automation.py           # SAP GUI COM scripting
|   +-- checklist_capture_sap.py    # Capture SAP checklist number
|   +-- sap_config.py              # SAP-specific configuration
|   +-- add_new_row_data_nis.py     # Tracker update (SAP side)
|   +-- log_manager.py             # SAP logging
|   +-- temp-book/                  # SAP temp directory
|
+-- test-emails/                     # Sample test email files
+-- debit-credit-email-pdf-reply/    # Sample D/C reply PDFs
+-- invoice example/                 # Sample invoice PDFs
+-- DSM_NIS_WORKFLOW_ARCHITECTURE.md # Detailed architecture document
```

---

## Module Reference

| File | Role |
|------|------|
| `main.py` | NIS orchestrator: scan emails -> classify -> prepare rows -> parallel book -> D/C -> approvals -> completion |
| `config.py` | Single source of truth for all credentials, URLs, and settings. Every value overridable via env var. |
| `db_manager.py` | Dual-backend (PostgreSQL/SQLite) persistence. 5 tables: email_runs, booking_rows, run_steps, debit_credit_emails, debit_credit_pdfs. |
| `read_email.py` | Fetches emails from Outlook via Graph API. Filters by `DSM-NIS-BOOKING` keyword. Downloads PDF attachments. |
| `classified_html_table_parser.py` | Pure-Python HTML parser. Extracts metadata_table + spv_dsm_table. Normalizes dates to DD-MM-YYYY. |
| `nis_booking.py` | Full Playwright browser automation: login (password+OTP), form fill, PDF upload, expense items, approver, submit, NIS number capture. |
| `read_vendor_master_data.py` | Reads vendor master from DSM-NIS-INPUT.xlsx on SharePoint. Supports exact + fuzzy SPV name matching. |
| `invoice_extractor.py` | Extracts invoice number from PDF via pdfplumber + LLM. Supports INOX multi-page bills (3-row tables). |
| `add_new_row_data_nis.py` | Writes/merges rows to DSM-NIS-OUTPUT.xlsx on SharePoint. Per-SPV sheets, duplicate detection, col K merging. |
| `debit_credit_email.py` | Builds the transposed HTML email for 5-row trigger. Indian number format (21,63,738.00). |
| `debit_credit_pdf_extractor.py` | Extracts ref_number, amount, doc_date from finance reply PDFs (LLM + keyword fallback). |
| `sharepoint_upload.py` | Upload/download/list files on SharePoint. Handles large files via upload sessions. |
| `pdf_merger.py` | Merges multiple PDFs (NIS checklist + SAP checklist + invoice + email) into one. |
| `sap_main.py` | SAP orchestrator: finds pending tasks via DB scan, processes each, handles recovery markers. |
| `sap_automation.py` | SAP GUI COM scripting: ZDCC transaction, form fill, save, print PDF, capture checklist number. |

---

## Configuration

All configuration lives in `nis_booking/config.py`. Override via environment variables or `.env` file.

### Key Settings

| Variable | Default | Purpose |
|----------|---------|---------|
| `DSM_SUBJECT_KEYWORD` | `DSM-NIS-BOOKING` | Email subject filter keyword |
| `DB_BACKEND` | `postgres` | Database backend (`postgres` / `sqlite`) |
| `NIS_PARALLEL_COUNT` | `3` | Concurrent NIS booking workers |
| `NIS_MAX_ROW_RETRIES` | `3` | Max retries per booking row |
| `DEBIT_CREDIT_ENABLED` | `true` | Enable 5-row debit/credit flow |
| `DEBIT_CREDIT_TRIGGER_ROW_COUNT` | `5` | SPV row count that triggers D/C |
| `NIS_SKIP_INVOICE_EXTRACTION` | `false` | Skip PDF invoice number extraction |
| `NIS_EMAIL_PDF_SPV_KEYWORDS` | `AWEK06L,AWEK03L,AWEK05L` | SPVs that use email.pdf instead of invoice |
| `NIS_AUTO_INVOICE_VENDOR_CODES` | `215943,215944,168856` | Vendors with auto-generated invoice number |
| `FAILURE_EMAIL_ENABLED` | `true` | Send failure notifications |
| `LLM_USE_OPENAI_DIRECT` | `false` | Use public OpenAI instead of Azure endpoint |

### Credential Sections in config.py

| Section | Key Variables |
|---------|--------------|
| **Graph API (Email)** | `GRAPH_TENANT_ID`, `GRAPH_CLIENT_ID`, `GRAPH_CLIENT_SECRET`, `USER_EMAIL` |
| **Graph API (SharePoint)** | `SP_TENANT_ID`, `SP_CLIENT_ID`, `SP_CLIENT_SECRET` |
| **SharePoint Location** | `TENANT_HOST`, `SITE_NAME`, `DRIVE_NAME`, `FOLDER_PATH` |
| **NIS Portal** | `NIS_LOGIN_URL`, `NIS_USERNAME`, `NIS_PASSWORD`, `NIS_OTP_SECRET` |
| **Database** | `POSTGRES_HOST`, `POSTGRES_DB`, `POSTGRES_USER`, `POSTGRES_PASSWORD` |
| **Azure OpenAI** | `LLM_ENDPOINT`, `LLM_DEPLOYMENT`, `LLM_API_KEY` |

---

## Installation

```bash
# 1. Clone the repository
git clone <repo-url>
cd DSM-ongoing

# 2. Create virtual environment
python -m venv .venv
source .venv/bin/activate  # Linux/Mac
# .venv\Scripts\activate   # Windows

# 3. Install dependencies
pip install -r requirements.txt

# 4. Install Playwright browsers (NIS service only)
python -m playwright install chromium

# 5. Configure environment
cp .env.example .env
# Edit .env with your credentials

# 6. Initialize database (auto-creates tables on first run)
python -c "import sys; sys.path.insert(0,'nis_booking'); import db_manager; db_manager.init_db()"
```

### Requirements
- Python 3.10+
- PostgreSQL (production) or SQLite (local dev)
- Microsoft Edge (for NIS portal automation)
- SAP GUI + COM scripting (for SAP automation, Windows only)

---

## Running the Pipeline

### NIS Booking Service (any host with Playwright)

```bash
# One-shot: process all pending emails, then exit after 20 idle scans
python run_nis_booking.py --lookback 2

# Custom lookback period
python run_nis_booking.py --lookback 7
```

### SAP Automation Service (Windows VM with SAP GUI)

```bash
# One-shot: process all pending SAP tasks, then exit
python run_sap_automation.py

# Poll mode: continuously check every 60 seconds
python run_sap_automation.py --poll --interval 60
```

### Both services can run simultaneously on different hosts -- they coordinate via the shared database.

---

## Test Cases


### DSM Booking Emails

| ID | Email Subject | Rows | Expected Result |
|----|--------------|------|-----------------|
| TC-01 | `Fw: DSM-NIS-BOOKING` | 1 SPV | Standard single booking |
| TC-02 | `DSM-NIS-BOOKING Week 08` | 2 (Solar/Wind) | Combined into ONE NIS booking |
| TC-03 | `AUTO-DSM-NIS-Booking` | 3 SPVs | Per-row invoice pre-extraction |
| TC-04 | `Re: DSM-NIS-BOOKING` | 3 (INOX bill) | Per-page invoice extraction |
| TC-05 | `dsm-nis-booking week 34` | 5 SPVs | Book all + send D/C email |
| TC-06 | `FWD: DSM-NIS-Booking` | All +ve | NIS only, done immediately |
| TC-07 | `DSM-NIS-BOOKING` | All -ve | SAP only, done after SAP |
| TC-08 | `DSM-NIS-BOOKING` | Mixed +/-ve | NIS + SAP, done when both finish |
| TC-09 | `DSM-NIS-BOOKING` | All zero | Row skipped |
| TC-10 | `DSM-NIS-BOOKING` | AWEK SPVs | Uses email.pdf, auto invoice |
| TC-11 | Same message_id | Duplicate | Skipped |
| TC-12 | `DSM-NIS-BOOKING` | No tables | Failure email sent |
| TC-13 | `DSM-NIS-BOOKING` | Unknown SPV | Vendor lookup fails -> failure email |

### Approval Emails

| ID | Email Subject | Expected Result |
|----|--------------|-----------------|
| TC-20 | `NIS Approval - 7000088592` | HOD approved, forwarded to recipients |
| TC-21 | `NIS Checklist Print - 7000088592` | FC approved, PDF uploaded, merge triggered |
| TC-22 | `NIS Approval - ...` (row not done) | Skipped (requires row done first) |
| TC-23 | `NIS Checklist Print - ...` (no HOD) | Skipped (requires HOD first) |
| TC-24 | Duplicate approval | Idempotent (already approved) |
| TC-25 | Unknown NIS number | No matching row -> skipped |

### Debit/Credit Replies

| ID | Email Subject | Expected Result |
|----|--------------|-----------------|
| TC-30 | `Re: Debit/Credit Notes \| Week 04 \| ...` | PDFs downloaded + extracted |
| TC-31 | `Fw: Debit/Credit Notes \| Week 04 \| ...` | Upload to Debit-Credit-Notes/ subfolder |
| TC-32 | Reply with no PDFs | Skipped (no attachments) |
| TC-33 | (after TC-30) | SAP checklist per reply PDF |
| TC-34 | Duplicate reply | Idempotent (already uploaded) |
| TC-35 | `Debit/Credit Notes \| ...` (no Re:/Fw:) | Ignored (our own outgoing email) |

### Edge Cases

| ID | Scenario | Expected Result |
|----|----------|-----------------|
| TC-40 | No HTML tables in email | Failure email sent |
| TC-41 | Missing mandatory columns | Failure email with details |
| TC-42 | NIS portal timeout | Retried up to 3 times |
| TC-43 | SAP GUI error | Step retried up to 3 times |
| TC-44 | Service crash mid-booking | Resumes from incomplete step on restart |
| TC-45 | Two workers same row | DB locks prevent double-booking |
| TC-46 | `[ALERT] Column Extraction...` | Excluded by subject filter |
| TC-47 | Email in Deleted Items | Filtered by parentFolderId |

---

## Resume & Fault Tolerance

- **Crash-resume:** Every step is tracked in the DB. On restart, the pipeline queries for incomplete runs/rows and resumes from the first unfinished step.
- **Idempotent operations:** Duplicate emails, approvals, and replies are detected and skipped.
- **Cost-center booking units:** Rows from the same email with the same non-empty cost center are submitted in one NIS booking and share its NIS number. Rows with different or unresolved cost centers remain separate.
- **Parallel safety:** Threading locks protect the tracker Excel writes and vendor data JSON updates.
- **SAP recovery markers:** If SAP saves successfully but the DB update fails, a local JSON marker file allows recovery on next run.
- **Retry limits:** Rows are retried up to `NIS_MAX_ROW_RETRIES` (3) times before being marked `failed`.
- **Failure notifications:** An LLM-powered explanation email is sent on failure, describing which value is missing/mismatched.

---

## SharePoint Excel Layout (DSM-NIS-OUTPUT.xlsx)

Each SPV gets its own sheet. Columns A-L:

| Col | Header | Source |
|-----|--------|--------|
| A | Week No | metadata.week_no |
| B | Week | "from_date to to_date" |
| C | Statement date | dsm_statement_published_date |
| D | Working Received Date | email received date |
| E | DSM Charges | total_dsm_charges_payable |
| F | Revenue Diff amount | revenue_diff |
| G | Drawl Charges Payable | drawl_charges_payable |
| H | Total Amount payable/Receivable | net_payable_receivable |
| I | NIS No | NIS checklist number (from booking) |
| J | Due Date | due_date |
| K | CN/DN Checklist No | SAP checklist number |
| L | CN/DN Date | Booking date / CN-DN doc date |

---

## Further Documentation

For detailed architecture diagrams, state machines, and complete workflow flows, see:
**[DSM_NIS_WORKFLOW_ARCHITECTURE.md](DSM_NIS_WORKFLOW_ARCHITECTURE.md)**
