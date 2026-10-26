# DSM-NIS Booking Automation

## Business Process Document

**Project:** DSM-NIS Booking Automation  
**Department:** AGEL (Adani Green Energy Limited)  
**Purpose:** Automate the weekly DSM (Deviation Settlement Mechanism) and NIS (Non-Invoice Settlement) booking process  
**Document For:** Business Users & Stakeholders

---

## What This System Does

This automation handles the end-to-end process of DSM-NIS booking that was previously done manually. When an authorized user sends an email with DSM charges data, the system automatically:

1. Reads the email
2. Extracts the data from tables in the email
3. Books the entries in the NIS portal
4. Creates SAP checklists for negative amounts
5. Gets approvals (HOD and FC)
6. Merges all related documents into one PDF
7. Sends status notifications at every stage

---

## How It Works - Step by Step

### Step 1: User Sends an Email

The user sends an email to the automation mailbox: **agel.nis.automation@adani.com**

**Important Rule:** The email subject must contain the keyword **"DSM-NIS-BOOKING"** (case does not matter). Only emails with this keyword in the subject are picked up by the system. Any other email is ignored.

**Example valid subjects:**
- "DSM-NIS-BOOKING"
- "DSM-NIS-BOOKING Week 04"
- "AUTO-DSM-NIS-Booking"

**Example invalid subjects (ignored):**
- "DSM report"
- "NIS weekly update"

The email body must contain **two tables**:
1. **Metadata Table** - with Week Number, From Date, To Date, and Due Date
2. **SPV DSM Data Table** - with SPV Names and their charge amounts (DSM Charges, Drawl Charges, Revenue Difference, Revenue Loss)

The email should also have the **invoice PDF** attached.

---

### Step 2: Bot Runs Automatically (3 Times a Day)

The NIS Bot runs **3 times every day** automatically. Each time it runs, it:

- Checks for any new DSM-NIS-BOOKING emails received in the last 2 days
- Picks up pending bookings that could not be completed in the previous run
- Processes all pending tasks to completion

This means even if one run fails (due to portal downtime, network issue, etc.), the next run will pick it up and complete it.

---

### Step 3: Email Detection & Validation

When the bot finds an email with "DSM-NIS-BOOKING" in the subject, it:

1. **Checks if already processed** - If this email was already handled, it skips it (no duplicate bookings)
2. **Downloads attachments** - All PDF files attached to the email are saved
3. **Reads the email body** - Extracts the HTML tables from the email content

---

### Step 4: Table Extraction

The system reads the two tables from the email:

**Table 1 - Metadata (Booking Period Information):**

| Week No | From Date | To Date | Due Date |
|---------|-----------|---------|----------|
| 04 | 20-Apr-2026 | 26-Apr-2026 | 15-May-2026 |

**Table 2 - SPV Charges:**

| SPV Name | DSM Charges | Drawl Charges | Revenue Diff | Revenue Loss |
|----------|-------------|---------------|--------------|--------------|
| AWEMP1PL | 1,016 | 0 | 4,517 | 0 |
| AWEK06L | -500 | 0 | 200 | 0 |

---

### Step 5: Amount Classification

For each SPV row, the system classifies each charge:

- **Positive amounts (+)** = Need NIS Booking (payment to vendor)
- **Negative amounts (-)** = Need SAP Checklist (amount receivable)
- **Zero amounts (0)** = Skipped (nothing to process)

---

### Step 6: Vendor Data Lookup

The system reads the vendor master data from the **DSM-NIS-INPUT.xlsx** file on SharePoint to get:

- Vendor Code
- Company Code
- Cost Center
- Plant
- Bank Key

This data is matched by the SPV name from the email.

---

### Step 7: NIS Portal Booking (For Positive Amounts)

The system opens the NIS Portal in a browser and:

1. Logs in automatically
2. Creates a new NIS entry
3. Uploads the invoice PDF
4. Fills in company code, vendor code, bank key
5. Fills in invoice number and date
6. Adds expense items (one for each positive charge)
7. Submits for approval
8. Captures the **NIS Checklist Number** (e.g., 7000088592)

---

### Step 8: SAP Automation (For Negative Amounts)

For negative amounts, the system opens SAP and:

1. Fills the debit/credit note form
2. Enters company code, vendor code, amount
3. Saves the entry
4. Captures the **SAP Checklist Number** (e.g., 5001482493)
5. Prints and saves the checklist as PDF
6. Uploads the PDF to the booking folder on SharePoint

---

### Step 9: HOD Approval

After the NIS booking is submitted:

1. The NIS portal sends an approval request to the **HOD (Head of Department)**
2. When HOD approves, the system receives an email with subject "NIS Approval - [number]"
3. The system detects this email and marks the booking as **HOD Approved**
4. A notification email is sent to the team: "HOD Approval Received"

---

### Step 10: FC Approval

After HOD approval:

1. The FC (Financial Controller) reviews and approves the NIS entry
2. When FC approves, the system receives an email with subject "NIS Checklist Print - [number]"
3. The system detects this email and marks the booking as **FC Approved**
4. The NIS Checklist PDF is downloaded and uploaded to the booking folder
5. A notification email is sent to the team: "FC Approval Received"

---

### Step 11: PDF Merge

After FC approval is received, all related documents are merged into a **single combined PDF**:

**Documents merged (in order):**
1. email.pdf (the original email converted to PDF)
2. Invoice/DSM attachment PDF(s)
3. NIS Checklist PDF
4. SAP Checklist PDF

**Merged file name format:** `merged_<SPV Name>_<To Date>.pdf`

**Merged file location on SharePoint:**
```
Bots/AGEL142 DSM NIS Booking Automation/booking/<SPV Name>/<From Date>_<To Date>/merged_<SPV>_<date>.pdf
```

---

### Step 12: Completion Notification

After all stages are complete (NIS + SAP + HOD + FC + PDF Merge), a **final completion email** is sent to the team with:

- SPV Name
- Period (From Date to To Date)
- NIS Booking Number
- SAP Checklist Number
- Booking Folder path on SharePoint

---

## Flow Diagram

```
 USER SENDS EMAIL
 (Subject: "DSM-NIS-BOOKING")
        |
        v
 +------------------+
 | BOT DETECTS      |
 | (Runs 3x daily)  |
 +------------------+
        |
        v
 +------------------+
 | EMAIL VALIDATED  |
 | Tables present?  |
 | Keyword match?   |
 +------------------+
        |
        v
 +------------------+
 | TABLES EXTRACTED |
 | - Metadata       |
 | - SPV Charges    |
 +------------------+
        |
        v
 +------------------+
 | AMOUNTS SORTED   |
 | + = NIS Booking  |
 | - = SAP Checklist|
 +------------------+
        |
        +-------------------+
        |                   |
        v                   v
 +-----------+      +-----------+
 | NIS       |      | SAP       |
 | BOOKING   |      | CHECKLIST |
 | (Portal)  |      | (SAP GUI) |
 +-----------+      +-----------+
        |                   |
        +-------------------+
        |
        v
 +------------------+
 | HOD APPROVAL     |
 | (Auto-detected   |
 |  from email)     |
 +------------------+
        |
        v
 +------------------+
 | FC APPROVAL      |
 | (Auto-detected   |
 |  from email)     |
 +------------------+
        |
        v
 +------------------+
 | PDF MERGE        |
 | All docs combined|
 | into one file    |
 +------------------+
        |
        v
 +------------------+
 | COMPLETION EMAIL |
 | Sent to team     |
 +------------------+
        |
        v
       DONE
```

---

## Output File (Tracker)

After each booking is complete, a row is added to the **DSM-NIS-OUTPUT.xlsx** file on SharePoint. Each SPV gets its own sheet:

| Column | Content |
|--------|---------|
| Week No | 04 |
| Week | 20-Apr-2026 to 26-Apr-2026 |
| Statement Date | DSM Statement Published Date |
| Working Received Date | Date email was received |
| DSM Charges | Total DSM charges payable |
| Revenue Diff | Revenue difference amount |
| Drawl Charges | Drawl charges payable |
| Total Amount | Net payable/receivable |
| NIS No | NIS Checklist Number |
| Due Date | Payment due date |
| CN/DN Checklist No | SAP Checklist Number |
| CN/DN Date | Booking date |

---

## Booking Folder Structure on SharePoint

All files for each booking are stored in a structured folder:

```
SharePoint > Documents > Bots > AGEL142 DSM NIS Booking Automation > booking
    |
    +-- <SPV Name>
         |
         +-- <From Date>_<To Date>
              |
              +-- email.pdf              (original email as PDF)
              +-- DSM_<invoice>.pdf      (invoice attachment)
              +-- nis_checklist_XXXX.pdf (NIS checklist after FC approval)
              +-- sap_checklist_XXXX.pdf (SAP checklist PDF)
              +-- merged_<SPV>_<date>.pdf (all above merged into one)
              |
              +-- Debit-Credit-Notes/    (only for 5-row bookings)
                   +-- 1.AWEK06L.pdf
                   +-- 2.AWEK03L.pdf
                   +-- 3.AWEK05L.pdf
```

---

## Phase 1 - Single Row Automation

**What was delivered in Phase 1:**

Phase 1 handled the most common scenario - **emails with 1 SPV row only**.

**Scope:**
- Email detection with "DSM-NIS-BOOKING" subject keyword
- Table extraction (metadata + 1 SPV row)
- Amount classification (positive = NIS, negative = SAP)
- Vendor master data lookup from SharePoint
- NIS Portal booking (browser automation for positive amounts)
- SAP Checklist creation (SAP GUI automation for negative amounts)
- SharePoint folder creation with email PDF and invoice PDF upload
- Output tracker (DSM-NIS-OUTPUT.xlsx) row update
- Completion email notification
- Failure email notification

**Limitations of Phase 1:**
- Could only process emails with exactly 1 SPV row
- No multi-row handling
- No debit/credit finance flow
- No invoice number extraction from PDF
- No HOD/FC approval tracking
- No PDF merge

---

## Phase 2 - Multi-Row, Debit/Credit & Invoice Extraction

**What was added in Phase 2:**

Phase 2 extended the system to handle complex real-world scenarios with multiple SPV rows and additional financial workflows.

### Multi-Row NIS Grouping (Cost Center)

After vendor data is resolved for all extracted rows in one email:

- Every row with the same non-empty **Cost Center** is grouped into a **single NIS booking**, regardless of SPV name
- Each positive value becomes a separate expense item in that booking
- Grouped rows share **one NIS Checklist Number**
- Rows with different or unresolved cost centers are booked separately
- Every row is still tracked separately in the output file

### 3-Row Automation (Invoice Pre-Extraction)

When the email has 3 SPV rows:

- The system **pre-extracts invoice numbers** from the attached PDF before booking
- If the PDF is an INOX bill (contains "INOX" keyword) - each SPV page gets its own invoice number
- If not INOX - one invoice number is shared across all 3 rows
- NIS grouping then follows the resolved vendor cost center
- Invoice extraction uses AI (LLM/GPT) to read the bill number from the PDF text

### 5-Row Automation (Debit/Credit Flow)

When the email has exactly 5 SPV rows, an additional finance workflow is triggered:

**After all 5 rows are booked (NIS + SAP complete):**

1. System builds a **Debit/Credit Notes email** with a transposed table showing all 5 SPVs
2. Email is sent to the **Finance Team** (recipients from SharePoint sheet "finance_team")
3. Finance team reviews and replies with **Debit/Credit Note PDFs** attached (e.g., 1.AWEK06L.pdf, 2.AWEK3L.pdf)
4. System detects the reply email and downloads the attached PDFs
5. From each PDF, the system extracts:
   - Reference Number (e.g., DN/26-27/259-DSM)
   - Amount
   - Document Date
   - Note Type (Debit or Credit)
6. For each reply PDF, a **SAP Checklist** is created (same as negative amount SAP flow)
7. Results are written to the output tracker

### Invoice Extraction (All Rows)

For normal vendors (not AWEK special SPVs), the system:

1. Reads the attached PDF using a PDF reader
2. Sends the text to AI (Azure OpenAI / GPT) to identify the invoice number
3. Uses that invoice number in the NIS booking form

For special SPVs (AWEK06L, AWEK03L, AWEK05L) where the PDF has no invoice data:
- The email body itself is converted to PDF and uploaded
- Invoice number is auto-generated

### HOD & FC Approval Tracking (All Rows)

Phase 2 added full approval lifecycle tracking:

**HOD Approval:**
- Detected when email arrives with subject "NIS Approval - [number]"
- Marks booking row as HOD Approved
- Sends notification email to team

**FC Approval:**
- Detected when email arrives with subject "NIS Checklist Print - [number]"
- Only processes if HOD is already approved
- Downloads and uploads the NIS Checklist PDF
- Marks booking row as FC Approved
- Triggers PDF Merge

### PDF Merge (All Rows)

After FC approval, all documents in the booking folder are merged:
- email.pdf (first)
- Invoice/DSM PDFs
- NIS Checklist PDF
- SAP Checklist PDF

Output: **One single merged PDF** uploaded to the same booking folder.

### Parallel Booking

Phase 2 also enabled **parallel processing** - up to 3 bookings can run at the same time, each in its own browser session. This significantly reduces the time for multi-row emails.

---

## Complete Workflow Diagram

```
                    +========================+
                    |   USER SENDS EMAIL     |
                    | Subject: DSM-NIS-      |
                    |         BOOKING        |
                    | Body: 2 tables         |
                    | Attachment: Invoice PDF |
                    +========================+
                              |
                              v
                    +========================+
                    |   BOT RUNS (3x DAILY)  |
                    | Checks mailbox for new |
                    | emails & pending tasks |
                    +========================+
                              |
                              v
                    +========================+
                    |   EMAIL DETECTED       |
                    | "DSM-NIS-BOOKING" in   |
                    | subject? YES -> process|
                    +========================+
                              |
                              v
                    +========================+
                    |   EXTRACT TABLES       |
                    | Table 1: Metadata      |
                    | Table 2: SPV Charges   |
                    | Download PDF attachment|
                    +========================+
                              |
                              v
                    +========================+
                    | HOW MANY SPV ROWS?     |
                    +========================+
                     /    |      |       \
                    /     |      |        \
                   v      v      v         v
              1 ROW  2 ROW  3 ROW    5 ROW
              (std)  (pair) (invoice) (debit/credit)
                   \      |      |        /
                    \     |      |       /
                     v    v      v      v
                    +========================+
                    | CLASSIFY AMOUNTS       |
                    | + amounts = NIS Book   |
                    | - amounts = SAP Check  |
                    | 0 amounts = Skip       |
                    +========================+
                              |
                    +---------+---------+
                    |                   |
                    v                   v
          +================+   +================+
          | NIS BOOKING    |   | SAP CHECKLIST  |
          | (positive amt) |   | (negative amt) |
          | Browser auto   |   | SAP GUI auto   |
          | -> NIS Number  |   | -> SAP Number  |
          | -> PDF upload  |   | -> PDF upload  |
          +================+   +================+
                    |                   |
                    +---------+---------+
                              |
                              v
                    +========================+
                    | HOD APPROVAL           |
                    | (Auto-detected email)  |
                    | "NIS Approval - XXXX"  |
                    +========================+
                              |
                              v
                    +========================+
                    | FC APPROVAL            |
                    | (Auto-detected email)  |
                    | "NIS Checklist Print   |
                    |  - XXXX"              |
                    +========================+
                              |
                              v
                    +========================+
                    | PDF MERGE              |
                    | All docs -> 1 merged   |
                    | PDF in booking folder  |
                    +========================+
                              |
                              v
                    +========================+
                    | TRACKER UPDATE         |
                    | DSM-NIS-OUTPUT.xlsx    |
                    | Row added per SPV      |
                    +========================+
                              |
                              v
                    +========================+
                    | COMPLETION EMAIL       |
                    | Sent to recipients     |
                    +========================+
                              |
                              v
           (IF 5-ROW TABLE)  |
                              v
                    +========================+
                    | DEBIT/CREDIT EMAIL     |
                    | Sent to Finance Team   |
                    +========================+
                              |
                              v
                    +========================+
                    | FINANCE TEAM REPLIES   |
                    | With DN/CN PDF files   |
                    +========================+
                              |
                              v
                    +========================+
                    | SAP CHECKLIST per PDF  |
                    | Extract ref/amount/    |
                    | date from each PDF     |
                    | Create SAP entry       |
                    +========================+
                              |
                              v
                           COMPLETE
```

---

## Summary Table: Phase 1 vs Phase 2

| Feature | Phase 1 | Phase 2 |
|---------|---------|---------|
| Email detection (DSM-NIS-BOOKING keyword) | Yes | Yes |
| Table extraction from email | Yes | Yes |
| 1-row SPV booking | Yes | Yes |
| 2-row Solar/Wind pair booking | No | Yes |
| 3-row with invoice pre-extraction | No | Yes |
| 5-row with Debit/Credit flow | No | Yes |
| Invoice number extraction from PDF (AI) | No | Yes |
| HOD Approval tracking | No | Yes |
| FC Approval tracking | No | Yes |
| PDF Merge (all docs into one) | No | Yes |
| Debit/Credit email to Finance | No | Yes |
| Finance reply PDF processing | No | Yes |
| SAP checklist for Debit/Credit notes | No | Yes |
| Parallel booking (up to 3 simultaneous) | No | Yes |
| Stage-wise notification emails | No | Yes |
| SharePoint booking folder with all docs | Yes | Yes |
| Output tracker (DSM-NIS-OUTPUT.xlsx) | Yes | Yes |

---

## Key Points for Business Users

1. **Only emails with "DSM-NIS-BOOKING" in the subject are processed** - any other email is safely ignored.

2. **The bot runs 3 times daily** - no manual trigger needed. If a task fails, it automatically retries in the next run.

3. **All documents are stored on SharePoint** at:  
   `Documents > Bots > AGEL142 DSM NIS Booking Automation > booking > [SPV Name] > [Period]`

4. **The final merged PDF** contains everything in one file - email, invoice, NIS checklist, and SAP checklist.

5. **Notifications are sent at every stage** - HOD approval received, FC approval received, and final completion.

6. **No duplicate bookings** - the system tracks every email by its unique ID and never processes the same email twice.

7. **Approvals are automatic** - the system detects HOD and FC approval emails and progresses the workflow without manual intervention.

---

*Document prepared for business stakeholders - DSM-NIS Booking Automation Project*
