"""
config.py  (nis_booking)
========================
Configuration for the NIS Booking service.
Covers: email, NIS portal, SharePoint, Sheet DB, logging.
SAP GUI settings are excluded (handled by sap_automation/).
"""

import os

def _load_env():
    _candidates = [
        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"),
        os.path.join(os.getcwd(), ".env"),
    ]
    for _path in _candidates:
        if os.path.isfile(_path):
            try:
                with open(_path, "r", encoding="utf-8") as _f:
                    for _line in _f:
                        _line = _line.strip()
                        if _line and not _line.startswith("#") and "=" in _line:
                            _k, _v = _line.split("=", 1)
                            _k = _k.strip()
                            _v = _v.strip().strip("'\"")
                            if _k not in os.environ:
                                os.environ[_k] = _v
            except Exception:
                pass

# Load environment variables from .env file
try:
    from dotenv import load_dotenv
    _root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    load_dotenv(os.path.join(_root_dir, ".env"))
    load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))
    load_dotenv()
except Exception:
    pass

_load_env()

# =============================================================================
# Microsoft Azure AD / Graph API  (shared by read_email, vendor master, tracker)
# =============================================================================

# --- App registration used for reading emails (Outlook / Graph) ---
GRAPH_TENANT_ID     = os.getenv("GRAPH_TENANT_ID")
GRAPH_CLIENT_ID     = os.getenv("GRAPH_CLIENT_ID")
GRAPH_CLIENT_SECRET = os.getenv("GRAPH_CLIENT_SECRET")

# Mailbox to monitor for DSM-NIS emails
USER_EMAIL = os.getenv("USER_EMAIL")

# --- App registration used for SharePoint / Excel operations ---
SP_TENANT_ID     = os.getenv("SP_TENANT_ID")
SP_CLIENT_ID     = os.getenv("SP_CLIENT_ID")
SP_CLIENT_SECRET = os.getenv("SP_CLIENT_SECRET")

# Microsoft Graph OAuth authority & scope
GRAPH_AUTHORITY = f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}"
SP_AUTHORITY    = f"https://login.microsoftonline.com/{SP_TENANT_ID}"
GRAPH_SCOPE     = ["https://graph.microsoft.com/.default"]

# =============================================================================
# SharePoint / OneDrive  (vendor master + NIS booking tracker Excel file)
# =============================================================================

TENANT_HOST = os.getenv("SP_TENANT_HOST")
SITE_NAME   = os.getenv("SP_SITE_NAME")
DRIVE_NAME  = os.getenv("SP_DRIVE_NAME")
FOLDER_PATH = os.getenv("SP_FOLDER_PATH")

FILE_NAME   = os.getenv("SP_FILE_NAME")

INPUT_FILE_NAME  = os.getenv("SP_INPUT_FILE_NAME")
INPUT_SHEET_NAME = os.getenv("SP_INPUT_SHEET_NAME")
OUTPUT_FILE_NAME = os.getenv("SP_OUTPUT_FILE_NAME")

TARGET_SHEET = os.getenv("SP_TARGET_SHEET")

SP_BOOKING_BASE = os.getenv("SP_BOOKING_BASE")

# =============================================================================
# NIS Portal (Playwright browser automation)
# =============================================================================

NIS_LOGIN_URL = os.getenv("NIS_LOGIN_URL")

NIS_EDGE_PATH = os.getenv("NIS_EDGE_PATH")

NIS_AUTH_FILE = os.getenv("NIS_AUTH_FILE")

NIS_USERNAME = os.getenv("NIS_USERNAME")
NIS_PASSWORD = os.getenv("NIS_PASSWORD")
NIS_OTP_SECRET = os.getenv("NIS_OTP_SECRET")

NIS_USE_COOKIES = (os.getenv("NIS_USE_COOKIES") or "").lower() in ("1", "true", "yes")

# =============================================================================
# Parallel / concurrent NIS booking
# =============================================================================
# How many NIS bookings to run AT THE SAME TIME. Each booking runs in its own
# thread with its own browser session and its own working directory, so their
# values never overlap. Rows are pulled from the global pool of pending booking
# rows across ALL emails — so if one email only has 2 rows, the remaining
# worker slot is filled with a row from another email.
#   1 → fully sequential (original behaviour)
#   N → up to N bookings in parallel
NIS_PARALLEL_COUNT = int(os.getenv("NIS_PARALLEL_COUNT")) if os.getenv("NIS_PARALLEL_COUNT") else None

# Max times a single booking row may be retried across runs before it is left
# in the "failed" state (prevents an unbookable row from looping forever).
NIS_MAX_ROW_RETRIES = int(os.getenv("NIS_MAX_ROW_RETRIES")) if os.getenv("NIS_MAX_ROW_RETRIES") else None

# Seconds to wait between launching each parallel worker. Default 0 → every
# worker opens its browser and logs in at the SAME time. Increase only if the
# portal struggles with a simultaneous login burst.
NIS_PARALLEL_LAUNCH_STAGGER = float(os.getenv("NIS_PARALLEL_LAUNCH_STAGGER")) if os.getenv("NIS_PARALLEL_LAUNCH_STAGGER") else 0.0

# When parallel booking is on, all sessions open + log in together and then
# wait at a barrier so NONE of them starts filling the form until EVERY session
# has finished logging in. This is the max seconds to wait at that barrier
# before a session gives up waiting and proceeds on its own (so one stuck login
# can't hang the whole batch). Set to 0 to disable the barrier entirely.
NIS_LOGIN_BARRIER_TIMEOUT = float(os.getenv("NIS_LOGIN_BARRIER_TIMEOUT")) if os.getenv("NIS_LOGIN_BARRIER_TIMEOUT") else 0.0
NIS_APP_IFRAME_TIMEOUT_MS = int(os.getenv("NIS_APP_IFRAME_TIMEOUT_MS")) if os.getenv("NIS_APP_IFRAME_TIMEOUT_MS") else None

# =============================================================================
# Invoice number handling
# =============================================================================
# Master switch: when TRUE, the invoice number is ALWAYS auto-generated for
# every vendor and the PDF invoice-number extraction (pdfplumber + LLM) is
# skipped entirely. Set to FALSE to enable PDF extraction for non-special
# vendors. Defaults to FALSE.
NIS_SKIP_INVOICE_EXTRACTION = (os.getenv("NIS_SKIP_INVOICE_EXTRACTION") or "").lower() in ("1", "true", "yes")

# For these vendor codes the attached PDF carries NO invoice data, so the NIS
# invoice number must keep being auto-generated. For every other vendor the
# invoice number is extracted from the PDF (pdfplumber + LLM) — but only when
# NIS_SKIP_INVOICE_EXTRACTION is False.
NIS_AUTO_INVOICE_VENDOR_CODES = {
    code.strip()
    for code in (os.getenv("NIS_AUTO_INVOICE_VENDOR_CODES") or "").split(",")
    if code.strip()
}

# SPV name keywords that trigger the "email PDF" exception rule:
#   - upload email.pdf (email converted to PDF) instead of the invoice PDF
#   - always auto-generate the invoice number (never extract from PDF)
# Comma-separated partial match keywords, case-insensitive.
# Example: "AWEK06L,AWEK03L,AWEK05L"
NIS_EMAIL_PDF_SPV_KEYWORDS = {
    kw.strip().upper()
    for kw in (os.getenv("NIS_EMAIL_PDF_SPV_KEYWORDS") or "").split(",")
    if kw.strip()
}

# --- Azure OpenAI (used to extract the invoice number from PDF text) ---
LLM_ENDPOINT    = os.getenv("LLM_ENDPOINT")
LLM_DEPLOYMENT  = os.getenv("LLM_DEPLOYMENT")
LLM_API_VERSION = os.getenv("LLM_API_VERSION")
LLM_API_KEY     = os.getenv("LLM_API_KEY")

# When TRUE, the invoice extractor bypasses the Azure llm_client endpoint above
# and calls the public OpenAI API directly using OPENAI_API_KEY (read from the
# .env file). Use this on machines where the internal Azure endpoint is not
# reachable.
LLM_USE_OPENAI_DIRECT = (os.getenv("LLM_USE_OPENAI_DIRECT") or "").lower() in ("1", "true", "yes")
OPENAI_API_KEY  = os.getenv("OPENAI_API_KEY")
OPENAI_MODEL    = os.getenv("OPENAI_MODEL")
OPENAI_BASE_URL = os.getenv("OPENAI_BASE_URL")

# When TRUE, Debit/Credit reply-note PDFs are parsed with the LLM (same Azure
# endpoint as llm_client.bat / invoice_extractor) to pull Amount, Reference
# Number and Date. When FALSE, the deterministic keyword/regex parser is used.
# On any LLM failure the extractor automatically falls back to the keyword
# parser so the flow never breaks.
USE_LLM_EXTRACTION = (os.getenv("USE_LLM_EXTRACTION") or "").lower() in ("1", "true", "yes")

# Keyword whose presence in an attached invoice PDF signals a MULTI-invoice
# bill (one invoice per SPV page). Used only for 3-row DSM tables: if present,
# each row's invoice number is extracted from the page matching its SPV name;
# if absent, the single-page invoice number is shared by all rows.
INOX_KEYWORD = os.getenv("INOX_KEYWORD")

# =============================================================================
# Email subject keyword filter
# =============================================================================

DSM_SUBJECT_KEYWORD = os.getenv("DSM_SUBJECT_KEYWORD")

# =============================================================================
# Sheet DB storage (legacy SharePoint/Excel persistence — kept for reference)
# =============================================================================

SHEET_DB_PATH = os.getenv("SHEET_DB_PATH")
SHEET_DB_USE_SHAREPOINT = True

# =============================================================================
# Database backend (SQLite or PostgreSQL)
# =============================================================================
# DB_BACKEND controls which database engine is used for persistence.
#   "sqlite"   → local SQLite file (default, original behaviour)
#   "postgres"  → Azure/remote PostgreSQL server
# When set to "postgres", tables are auto-created if they don't exist.
DB_BACKEND = (os.getenv("DB_BACKEND") or "").strip().lower()

# =============================================================================
# SQLite DB storage
# =============================================================================
SQLITE_DB_PATH = os.getenv("SQLITE_DB_PATH")

# POSTGRES_HOST = "az10psqldmrcbtp01.postgres.database.azure.com"
#     POSTGRES_USER = "psqladmin"
#     POSTGRES_PORT = "5432"

# =============================================================================
# PostgreSQL connection (only used when DB_BACKEND=postgres)
# =============================================================================
# Standard libpq-style connection parameters. All can be set individually or
# via a single DSN connection string (POSTGRES_DSN takes precedence).
POSTGRES_DSN      = os.getenv("POSTGRES_DSN")
POSTGRES_HOST     = os.getenv("POSTGRES_HOST")
POSTGRES_PORT     = int(os.getenv("POSTGRES_PORT")) if os.getenv("POSTGRES_PORT") else None
POSTGRES_DB       = os.getenv("POSTGRES_DB")
POSTGRES_USER     = os.getenv("POSTGRES_USER")
POSTGRES_PASSWORD = os.getenv("POSTGRES_PASSWORD")
POSTGRES_SSLMODE  = os.getenv("POSTGRES_SSLMODE")

# =============================================================================
# Completion-email recipients (stored on SharePoint, NOT in SQLite)
# =============================================================================
# Recipients live in a SharePoint workbook (sheet "nis_completion_recipients").
NIS_RECIPIENTS_SP_FOLDER = os.getenv("NIS_RECIPIENTS_SP_FOLDER")
NIS_RECIPIENTS_FILE_NAME = os.getenv("NIS_RECIPIENTS_FILE_NAME")
NIS_RECIPIENTS_SHEET     = os.getenv("NIS_RECIPIENTS_SHEET")

# =============================================================================
# Attachment storage
# =============================================================================

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))

ATTACHMENT_DIR = os.getenv("ATTACHMENT_DIR")

# =============================================================================
# Logging
# =============================================================================
LOG_USE_CUSTOM_PATH = (os.getenv("LOG_USE_CUSTOM_PATH") or "").lower() in ("1", "true", "yes")

LOG_BASE_PATH = os.getenv("LOG_BASE_PATH")

# =============================================================================
# Failure email notification
# =============================================================================
FAILURE_EMAIL_ENABLED = (os.getenv("FAILURE_EMAIL_ENABLED") or "").lower() in ("1", "true", "yes")
FAILURE_EMAIL_TO      = os.getenv("FAILURE_EMAIL_TO")

# Fallback recipients used only when the SharePoint recipients sheet is empty
# or cannot be read. Comma or semicolon separated. Defaults to FAILURE_EMAIL_TO.
NIS_COMPLETION_RECIPIENTS = os.getenv("NIS_COMPLETION_RECIPIENTS")


# =============================================================================
# Debit/Credit Note email to Finance team (5-row DSM table trigger)
# =============================================================================
# When a DSM email's SPV table contains EXACTLY this many valid (non-Total) SPV
# rows, the automation builds a "Debit/Credit Notes" email (mirroring the
# template in send-debit-credit-email-finance/) from the same table values and
# sends it to the finance team. A finance team member later replies with the
# Debit/Credit Note PDFs (typically 3), which the automation downloads and
# stores in the booking folder. Every step is tracked in the SQLite DB
# (table `debit_credit_emails`).
DEBIT_CREDIT_ENABLED = (os.getenv("DEBIT_CREDIT_ENABLED") or "").lower() in ("1", "true", "yes")

# Number of valid SPV rows that triggers the debit/credit flow.
DEBIT_CREDIT_TRIGGER_ROW_COUNT = int(os.getenv("DEBIT_CREDIT_TRIGGER_ROW_COUNT")) if os.getenv("DEBIT_CREDIT_TRIGGER_ROW_COUNT") else None

# Subject used for the outgoing debit/credit email. A run-specific marker
# (week / period) is appended so replies can be matched back to the run.
DEBIT_CREDIT_SUBJECT = os.getenv("DEBIT_CREDIT_SUBJECT")

# Keyword that must appear in an incoming subject for it to be considered a
# possible finance-team reply carrying the Debit/Credit Note PDFs.
DEBIT_CREDIT_SUBJECT_KEYWORD = os.getenv("DEBIT_CREDIT_SUBJECT_KEYWORD")

# Company code shown on the debit/credit email ("C. Code"). Left blank by
# default; set it here or via env if Finance expects a fixed value.
DEBIT_CREDIT_COMPANY_CODE = os.getenv("DEBIT_CREDIT_COMPANY_CODE")

# Sub-folder (inside each booking folder) where the finance-team reply PDFs are
# stored on SharePoint.
DEBIT_CREDIT_REPLY_SUBFOLDER = os.getenv("DEBIT_CREDIT_REPLY_SUBFOLDER")

# Finance-team recipients live in the SharePoint workbook `automation_sheet_db.xlsx`
# (config.SHEET_DB_PATH) on sheet `finance_team`, single column header `emails`.
FINANCE_TEAM_FILE_NAME  = os.getenv("FINANCE_TEAM_FILE_NAME")
FINANCE_TEAM_SP_FOLDER  = os.getenv("FINANCE_TEAM_SP_FOLDER")
FINANCE_TEAM_SHEET      = os.getenv("FINANCE_TEAM_SHEET")

# Fallback recipients used only when the finance_team sheet is empty/unreadable.
FINANCE_TEAM_RECIPIENTS = os.getenv("FINANCE_TEAM_RECIPIENTS")

# Run SAP checklist automation for each finance-reply Debit/Credit PDF. Requires
# the Windows SAP GUI environment (the sap_automation package). When it cannot
# be imported (e.g. non-Windows host), the PDF rows stay pending for retry.
DEBIT_CREDIT_SAP_ENABLED = (os.getenv("DEBIT_CREDIT_SAP_ENABLED") or "").lower() in ("1", "true", "yes")
