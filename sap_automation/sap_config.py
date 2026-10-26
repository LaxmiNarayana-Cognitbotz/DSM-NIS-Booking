"""
sap_config.py  (sap_automation)
===============================
Configuration for the SAP Automation service.
Covers: SAP GUI, SharePoint, Sheet DB, logging, vendor master.
This file allows sap_main.py to run independently.
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
# Microsoft Azure AD / Graph API
# =============================================================================

# App registration used for Outlook / Graph email notifications.
# Keep these aligned with nis_booking/config.py so both services send mail
# through the same Graph app and mailbox.
GRAPH_TENANT_ID     = os.getenv("GRAPH_TENANT_ID")
GRAPH_CLIENT_ID     = os.getenv("GRAPH_CLIENT_ID")
GRAPH_CLIENT_SECRET = os.getenv("GRAPH_CLIENT_SECRET")

USER_EMAIL = os.getenv("USER_EMAIL")

# App registration used for SharePoint / Sheet DB access.
SP_TENANT_ID     = os.getenv("SP_TENANT_ID")
SP_CLIENT_ID     = os.getenv("SP_CLIENT_ID")
SP_CLIENT_SECRET = os.getenv("SP_CLIENT_SECRET")

GRAPH_AUTHORITY = f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}"
SP_AUTHORITY    = f"https://login.microsoftonline.com/{SP_TENANT_ID}"
GRAPH_SCOPE     = ["https://graph.microsoft.com/.default"]

# =============================================================================
# SharePoint / OneDrive
# =============================================================================

TENANT_HOST = os.getenv("SP_TENANT_HOST")
SITE_NAME   = os.getenv("SP_SITE_NAME")
DRIVE_NAME  = os.getenv("SP_DRIVE_NAME")
FOLDER_PATH = os.getenv("SP_FOLDER_PATH")

INPUT_FILE_NAME  = os.getenv("SP_INPUT_FILE_NAME")
INPUT_SHEET_NAME = os.getenv("SP_INPUT_SHEET_NAME")
OUTPUT_FILE_NAME = os.getenv("SP_OUTPUT_FILE_NAME")

SP_BOOKING_BASE = os.getenv("SP_BOOKING_BASE")

# =============================================================================
# SAP GUI Automation
# =============================================================================

SAP_LOGON_PATH = os.getenv("SAP_LOGON_PATH")
SAP_SYSTEM     = os.getenv("SAP_SYSTEM")
SAP_CLIENT     = os.getenv("SAP_CLIENT")
SAP_USERNAME   = os.getenv("SAP_USERNAME")
SAP_PASSWORD   = os.getenv("SAP_PASSWORD")
SAP_LANGUAGE   = os.getenv("SAP_LANGUAGE")

# SAP form field defaults
SAP_CHK_DOC_TYP  = os.getenv("SAP_CHK_DOC_TYP")
SAP_DIG_SIGN_INV = os.getenv("SAP_DIG_SIGN_INV")
SAP_GJAHR        = os.getenv("SAP_GJAHR")
SAP_DEB_NOT_REF  = os.getenv("SAP_DEB_NOT_REF")
SAP_DOC_SUB_CAT  = os.getenv("SAP_DOC_SUB_CAT")
SAP_DO_PREVIEW   = os.getenv("SAP_DO_PREVIEW")

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))

SAP_PDF_TARGET_DIR = os.getenv("SAP_PDF_TARGET_DIR")

# =============================================================================
# Sheet DB storage (primary persistence — shared with nis_booking)
# =============================================================================

SHEET_DB_PATH = os.getenv("SHEET_DB_PATH")
SHEET_DB_USE_SHAREPOINT = True

# =============================================================================
# Logging
# =============================================================================
LOG_USE_CUSTOM_PATH = (os.getenv("LOG_USE_CUSTOM_PATH") or "").lower() in ("1", "true", "yes")

LOG_BASE_PATH = os.getenv("LOG_BASE_PATH")

# =============================================================================
# SAP Automation – Retry settings
# =============================================================================
MAX_STEP_RETRIES = int(os.getenv("SAP_MAX_STEP_RETRIES")) if os.getenv("SAP_MAX_STEP_RETRIES") else None

# Dialog timeout settings (in seconds)
SAP_PRINT_DIALOG_WAIT = int(os.getenv("SAP_PRINT_DIALOG_WAIT")) if os.getenv("SAP_PRINT_DIALOG_WAIT") else None
SAP_SAVE_DIALOG_CONNECT_TIMEOUT = int(os.getenv("SAP_SAVE_DIALOG_CONNECT_TIMEOUT")) if os.getenv("SAP_SAVE_DIALOG_CONNECT_TIMEOUT") else None
SAP_SAVE_DIALOG_VISIBLE_TIMEOUT = int(os.getenv("SAP_SAVE_DIALOG_VISIBLE_TIMEOUT")) if os.getenv("SAP_SAVE_DIALOG_VISIBLE_TIMEOUT") else None
SAP_SAVE_DIALOG_RETRIES = int(os.getenv("SAP_SAVE_DIALOG_RETRIES")) if os.getenv("SAP_SAVE_DIALOG_RETRIES") else None
