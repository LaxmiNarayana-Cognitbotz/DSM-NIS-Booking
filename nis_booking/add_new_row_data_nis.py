import os
import io
import sys
import json
import time
import requests
import msal
import math
import config

# =========================
# CONFIG  (all values loaded from config.py)
# =========================
TENANT_ID     = config.SP_TENANT_ID
CLIENT_ID     = config.SP_CLIENT_ID
CLIENT_SECRET = config.SP_CLIENT_SECRET

TENANT_HOST  = config.TENANT_HOST
SITE_NAME    = config.SITE_NAME
DRIVE_NAME   = config.DRIVE_NAME
FOLDER_PATH  = config.FOLDER_PATH

# ── Output file settings ─────────────────────────────────────────────────
OUTPUT_FILE_NAME = config.OUTPUT_FILE_NAME   # DSM-NIS-OUTPUT.xlsx

AUTHORITY = config.SP_AUTHORITY
SCOPES    = config.GRAPH_SCOPE

# =========================
# SHEET LAYOUT  (DSM-NIS-OUTPUT.xlsx)
# =========================
# Each sheet is named after an SPV.
# Columns A to L (12 columns total):
#   A  = Week No                                     ← week_no
#   B  = Week                                        ← "from_date TO to_date"
#   C  = Statement date                              ← dsm_statement_published_date
#   D  = Working Received Date                       ← email received date
#   E  = DSM Charges                                 ← total_dsm_charges_payable
#   F  = Revenue Diff amount                         ← revenue_diff
#   G  = Drawl Charges Payable                       ← drawl_charges_payable
#   H  = Total Amount payable/Receivble (in Rs)      ← net_payable_receivable
#   I  = NIS No                                      ← checklist_id (NIS booking number)
#   J  = Due Date                                    ← due_date
#   K  = CN/DN Checklist No                          ← sap_checklist_number
#   L  = CN/DN Date                                  ← nis_booking_date

HEADER_ROW      = 1
START_DATA_ROW  = HEADER_ROW + 1   # row 2
START_COL, END_COL = "A", "L"
MAX_SCAN_ROW    = 5000

EXPECTED_COLS = ord(END_COL) - ord(START_COL) + 1  # 12 columns A..L

OUTPUT_HEADERS = [
    "Week No",
    "Week",
    "Statement date",
    "Working Received Date",
    "DSM Charges",
    "Revenue Diff amount",
    "Drawl Charges Payable",
    "Total Amount payable/Receivble (in Rs)",
    "NIS No",
    "Due Date",
    "CN/DN Checklist No",
    "CN/DN Date",
]


# =========================
# AUTH
# =========================
def get_token(retries: int = 0, backoff_sec: float = 1.5) -> str:
    app = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
    )
    last_err = None
    for attempt in range(retries + 1):
        result = app.acquire_token_for_client(scopes=SCOPES)
        if "access_token" in result:
            return result["access_token"]
        last_err = result
        if result.get("error") == "invalid_client":
            break
        if attempt < retries:
            time.sleep(backoff_sec * (2 ** attempt))
    raise RuntimeError(
        "Failed to acquire token. If error is 'invalid_client', verify the client secret.\n"
        f"MSAL response:\n{json.dumps(last_err, indent=2)}"
    )


# =========================
# HTTP HELPERS
# =========================
def get_json(url: str, headers: dict) -> dict:
    r = requests.get(url, headers=headers)
    try:
        r.raise_for_status()
    except Exception:
        print("GET failed:", url)
        print("Status:", r.status_code)
        print(r.text[:2048])
        raise
    try:
        return r.json()
    except Exception:
        return {}


def patch_json(url: str, headers: dict, payload: dict) -> dict:
    r = requests.patch(url, headers={**headers, "Content-Type": "application/json"}, json=payload)
    try:
        r.raise_for_status()
    except Exception:
        print("PATCH failed:", url)
        print("Status:", r.status_code)
        print(r.text[:2048])
        raise
    try:
        return r.json()
    except Exception:
        return {}


# =========================
# GRAPH HELPERS
# =========================
def resolve_site_and_drive(headers_auth: dict):
    site_url = f"https://graph.microsoft.com/v1.0/sites/{TENANT_HOST}:/sites/{SITE_NAME}"
    site = get_json(site_url, headers_auth) or {}
    site_id = site.get("id")
    if not site_id:
        raise RuntimeError("Failed to resolve Site ID.")

    drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    drives = get_json(drives_url, headers_auth) or {}
    all_drives = drives.get("value") or []

    drive_id = None
    for d in all_drives:
        if d.get("name") == DRIVE_NAME:
            drive_id = d.get("id")
            break
    if not drive_id:
        default_drive = get_json(
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive", headers_auth
        ) or {}
        drive_id = default_drive.get("id")
    if not drive_id:
        raise RuntimeError("Failed to resolve Drive ID.")
    return site_id, drive_id


def resolve_item_id_by_path(drive_id: str, item_path: str, headers_auth: dict) -> str:
    item_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{item_path}"
    item = get_json(item_url, headers_auth) or {}
    item_id = item.get("id")
    if not item_id:
        raise RuntimeError("Unable to resolve item ID for the Excel file.")
    return item_id


def _set_column_widths(drive_id: str, item_id: str, sheet_name: str, headers_auth: dict):
    """
    Set explicit column widths for columns A-L so the data is readable.
    Graph API does not support true auto-fit, so we set generous fixed widths
    based on the expected header/data lengths.
    """
    # (column letter, width in points)  – tuned for the OUTPUT headers
    col_widths = [
        ("A", 70),   # Week No
        ("B", 170),  # Week (from_date to to_date)
        ("C", 120),  # Statement date
        ("D", 160),  # Working Received Date
        ("E", 120),  # DSM Charges
        ("F", 140),  # Revenue Diff amount
        ("G", 150),  # Drawl Charges Payable
        ("H", 220),  # Total Amount payable/Receivble (in Rs)
        ("I", 140),  # NIS No
        ("J", 100),  # Due Date
        ("K", 150),  # CN/DN Checklist No
        ("L", 110),  # CN/DN Date
    ]
    for col_letter, width in col_widths:
        col_url = (
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
            f"/workbook/worksheets('{sheet_name}')"
            f"/range(address='{col_letter}:{col_letter}')/format"
        )
        try:
            patch_json(col_url, headers_auth, {"columnWidth": width})
        except Exception as exc:
            print(f"[WARN] Could not set width for column {col_letter}: {exc}")
    print(f"[OK] Column widths set for sheet '{sheet_name}'")


def ensure_sheet_name(drive_id: str, item_id: str, headers_auth: dict, target: str) -> str:
    """
    Check if a worksheet named *target* exists in the workbook.
    If it does, return the name.  If not, **create** it, write the header
    row (A1:L1), and return the name.
    """
    ws_url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets"
    )
    sheets = get_json(ws_url, headers_auth) or {}
    names = [s.get("name") for s in (sheets.get("value") or []) if s.get("name")]

    if target in names:
        return target

    # ── Sheet does not exist → create it ──────────────────────────────────
    print(f"[INFO] Sheet '{target}' not found in OUTPUT file – creating …")
    create_payload = {"name": target}
    r = requests.post(
        ws_url,
        headers={**headers_auth, "Content-Type": "application/json"},
        json=create_payload,
    )
    try:
        r.raise_for_status()
    except Exception:
        print(f"[WARN] Sheet create response: {r.status_code} {r.text[:300]}")
        raise RuntimeError(f"Failed to create worksheet '{target}' in OUTPUT file.")

    print(f"[OK] Created worksheet '{target}'")

    # ── Write header row A1:L1 ────────────────────────────────────────────
    header_url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets('{target}')/range(address='A1:L1')"
    )
    patch_json(header_url, headers_auth, {"values": [OUTPUT_HEADERS]})

    # Bold + centre-align the header row
    fmt_url = header_url + "/format"
    patch_json(fmt_url, headers_auth, {
        "horizontalAlignment": "Center",
        "verticalAlignment": "Center",
    })
    font_url = fmt_url + "/font"
    patch_json(font_url, headers_auth, {"bold": True, "size": 12, "color": "#FFFFFF"})

    # ── Blue background fill on the header row ────────────────────────────
    fill_url = fmt_url + "/fill"
    patch_json(fill_url, headers_auth, {"color": "#4472C4"})
    print(f"[OK] Header row written with blue background for sheet '{target}'")

    # ── Set column widths so data is readable ─────────────────────────────
    _set_column_widths(drive_id, item_id, target, headers_auth)

    return target


def read_col_A_values(
    drive_id: str, item_id: str, sheet_name: str, headers_auth: dict, max_row: int = MAX_SCAN_ROW
):
    """Read column A (Week No) to find the first blank row."""
    address = f"A1:A{max_row}"
    url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets('{sheet_name}')/range(address='{address}')"
    )
    resp = get_json(url, headers_auth) or {}
    col_vals = resp.get("values")
    if not isinstance(col_vals, list):
        return []
    return col_vals


def _read_cols_A_B(
    drive_id: str, item_id: str, sheet_name: str, headers_auth: dict, max_row: int = MAX_SCAN_ROW
) -> list:
    """Read columns A-B (Week No + Week) to support duplicate detection."""
    address = f"A1:B{max_row}"
    url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets('{sheet_name}')/range(address='{address}')"
    )
    resp = get_json(url, headers_auth) or {}
    vals = resp.get("values")
    if not isinstance(vals, list):
        return []
    return vals


def _read_cols_A_to_L(
    drive_id: str, item_id: str, sheet_name: str, headers_auth: dict, max_row: int = MAX_SCAN_ROW
) -> list:
    """Read the full A-L range so we can locate an existing row and inspect its
    NIS (col I, index 8) and SAP/CN-DN (col K, index 10) cells for merging."""
    address = f"A1:L{max_row}"
    url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets('{sheet_name}')/range(address='{address}')"
    )
    resp = get_json(url, headers_auth) or {}
    vals = resp.get("values")
    if not isinstance(vals, list):
        return []
    return vals


def _split_checklist_numbers(cell) -> list:
    """Split a Column-K style cell into individual checklist numbers."""
    return [part.strip() for part in str(cell or "").split(",") if part.strip()]


def merge_checklist_numbers(existing_cell, new_number) -> str:
    """Merge *new_number* into the existing Column-K value, preserving every
    number already present and de-duplicating. Returns a comma-joined string.
    Guarantees no previously stored SAP checklist number is ever lost."""
    numbers = _split_checklist_numbers(existing_cell)
    for candidate in _split_checklist_numbers(new_number):
        if candidate and candidate not in numbers:
            numbers.append(candidate)
    return ", ".join(numbers)


def _is_duplicate_row(
    cols_ab: list,
    week_no: str,
    week_range: str,
    require_nis: bool = False,
) -> bool:
    """
    Return True if the sheet already contains a row with the same Week No (col A)
    AND Week range (col B).  Comparison is case-insensitive, stripped.
    Skips the header row (index 0).

    When *require_nis* is True, only rows that themselves carry a NIS booking
    number (col I, index 8) count as duplicates. This lets a NIS-booking write
    ignore Debit/Credit reply rows (which have a blank NIS column) for the same
    week, so a NIS number is never suppressed by a pre-existing reply row.
    """
    target_wno = str(week_no).strip().lower()
    target_wr  = str(week_range).strip().lower()

    for idx, row in enumerate(cols_ab):
        if idx == 0:              # skip header row
            continue
        if not isinstance(row, list) or len(row) < 2:
            continue
        cell_a = str(row[0]).strip().lower() if row[0] is not None else ""
        cell_b = str(row[1]).strip().lower() if row[1] is not None else ""
        if cell_a == target_wno and cell_b == target_wr:
            if require_nis:
                cell_i = str(row[8]).strip() if len(row) > 8 and row[8] is not None else ""
                if not cell_i:
                    continue   # a reply row (blank NIS) never blocks a NIS write
            return True
    return False


def write_row_A_to_M(
    drive_id: str,
    item_id: str,
    sheet_name: str,
    row_index: int,
    values: list,
    headers_auth: dict,
):
    address = f"{START_COL}{row_index}:{END_COL}{row_index}"
    url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets('{sheet_name}')/range(address='{address}')"
    )
    payload = {"values": [values]}
    patch_json(url, headers_auth, payload)
    print(f"[OK] Wrote row at {address}")

    fmt_url = url + "/format"
    patch_json(fmt_url, headers_auth, {
        "horizontalAlignment": "Center",
        "verticalAlignment": "Center",
    })

    font_url = fmt_url + "/font"
    patch_json(font_url, headers_auth, {"size": 12})
    print(f"[OK] Applied font size 12 and centre alignment to {address}")


# =========================
# ROW COMPUTATION
# =========================
def is_blank_cell(val) -> bool:
    return (val is None) or (isinstance(val, str) and val.strip() == "")


def try_parse_int(val):
    if val is None:
        return None
    if isinstance(val, int):
        return int(val)
    if isinstance(val, float):
        if math.isnan(val):
            return None
        return int(val) if float(val).is_integer() else None
    if isinstance(val, str):
        s = val.strip()
        if s.isdigit():
            try:
                return int(s)
            except Exception:
                return None
    return None


def find_first_blank_row_in_A(col_vals: list, start_row: int, max_row: int) -> int:
    """Find the first blank cell in column A (Week No), starting from start_row."""
    n = len(col_vals)
    row = start_row
    while row <= max_row:
        if row <= n:
            row_vals = col_vals[row - 1]
            cell_val = (
                row_vals[0]
                if (isinstance(row_vals, list) and len(row_vals) > 0)
                else None
            )
            if is_blank_cell(cell_val):
                return row
        else:
            return row
        row += 1
    return max_row


def compute_next_week_no(col_vals: list, start_row: int) -> int:
    max_seen = 0
    for idx, row in enumerate(col_vals, start=1):
        if idx < start_row:
            continue
        cell_val = (
            row[0] if (isinstance(row, list) and len(row) > 0) else None
        )
        v = try_parse_int(cell_val)
        if v is not None and v > max_seen:
            max_seen = v
    return max_seen + 1 if max_seen >= 0 else 1


# =========================
# VALUE HELPERS
# =========================
def safe_float(val):
    """Convert a string like '1,016.00' or '4517646' to float. Returns 0.0 on failure."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    cleaned = str(val).replace(",", "").strip()
    try:
        return float(cleaned)
    except (ValueError, TypeError):
        return 0.0


def build_row_values(
    metadata_row: dict,
    spv_dsm_row: dict,
    checklist_id: str = "",
    sap_checklist_number: str = "",
    email_received_date: str = "",
    nis_booking_date: str = "",
) -> list:
    """
    Build the 12-value list for columns A..L  (DSM-NIS-OUTPUT layout).

    Column mapping (0-based index → spreadsheet column):
      0  A  Week No                                    ← week_no
      1  B  Week                                       ← "from_date to to_date"
      2  C  Statement date                             ← dsm_statement_published_date
      3  D  Working Received Date                      ← email_received_date
      4  E  DSM Charges                                ← total_dsm_charges_payable
      5  F  Revenue Diff amount                        ← revenue_diff
      6  G  Drawl Charges Payable                      ← drawl_charges_payable
      7  H  Total Amount payable/Receivble (in Rs)     ← net_payable_receivable
      8  I  NIS No                                     ← checklist_id
      9  J  Due Date                                   ← due_date
     10  K  CN/DN Checklist No                         ← sap_checklist_number
     11  L  CN/DN Date                                 ← nis_booking_date
    """
    # ---- metadata fields ----
    week_no   = str(metadata_row.get("week_no", "")).strip()
    from_date = str(metadata_row.get("from_date", "")).strip()
    to_date   = str(metadata_row.get("to_date", "")).strip()
    due_date  = str(metadata_row.get("due_date", "")).strip()
    dsm_statement_published_date = str(metadata_row.get("dsm_statement_published_date", "")).strip()

    # Column B: "DD-MM-YYYY to DD-MM-YYYY"
    week_range = f"{from_date} to {to_date}" if from_date and to_date else ""

    # ---- spv dsm fields ----
    dsm_charges            = safe_float(spv_dsm_row.get("total_dsm_charges_payable"))
    revenue_diff           = safe_float(spv_dsm_row.get("revenue_diff"))
    drawl_charges_payable  = safe_float(spv_dsm_row.get("drawl_charges_payable"))
    net_payable_receivable = safe_float(spv_dsm_row.get("net_payable_receivable"))

    return [
        week_no,                        # A  – Week No
        week_range,                     # B  – Week ("from_date to to_date")
        dsm_statement_published_date,   # C  – Statement date
        email_received_date or "",      # D  – Working Received Date
        dsm_charges,                    # E  – DSM Charges
        revenue_diff,                   # F  – Revenue Diff amount
        drawl_charges_payable,          # G  – Drawl Charges Payable
        net_payable_receivable,         # H  – Total Amount payable/Receivble (in Rs)
        checklist_id or "",             # I  – NIS No
        due_date,                       # J  – Due Date
        sap_checklist_number or "",     # K  – CN/DN Checklist No
        nis_booking_date or "",         # L  – CN/DN Date
    ]


# =========================
# MAIN PUBLIC FUNCTION
# =========================
def add_incremental_week_row(
    metadata_row: dict,
    spv_dsm_row: dict,
    sheet_name: str = "",
    checklist_id: str = "",
    sap_checklist_number: str = "",
    email_received_date: str = "",
    nis_booking_date: str = "",
    allow_duplicate_week: bool = False,
):
    """
    Adds one new row to the DSM-NIS-OUTPUT Excel file on SharePoint.

    The target worksheet is the SPV name (*sheet_name*).  If the sheet does
    not exist it is created with headers in row 1 (columns A-L).

    Args:
        metadata_row          : first dict from parsed_tables["metadata_table"]
        spv_dsm_row           : matching dict from parsed_tables["spv_dsm_table"]
        sheet_name            : target worksheet = SPV name
        checklist_id          : NIS checklist/booking number (→ column I)
        sap_checklist_number  : SAP checklist number (→ column K)
        email_received_date   : date the email was received (→ column D)
        nis_booking_date      : date the NIS booking was done (→ column L)
        allow_duplicate_week  : when True, skip the (Week No + Week) duplicate
                                guard. Used for Wind/Solar SPV pairs that are
                                booked together and legitimately need their own
                                row for the same week (each variant is a row).
    """
    token = get_token()
    headers_auth = {"Authorization": f"Bearer {token}"}

    _, drive_id = resolve_site_and_drive(headers_auth)
    item_path   = f"{FOLDER_PATH}/{OUTPUT_FILE_NAME}".strip("/")
    item_id     = resolve_item_id_by_path(drive_id, item_path, headers_auth)
    resolved_sheet = ensure_sheet_name(drive_id, item_id, headers_auth, sheet_name)

    # ── Duplicate check: read cols A-L and look for same week_no + week ──
    # A NIS booking row is a duplicate only of another NIS row (col I filled),
    # so a Debit/Credit reply row (blank NIS) for the same week never blocks it.
    cols_ab = _read_cols_A_to_L(drive_id, item_id, resolved_sheet, headers_auth, max_row=MAX_SCAN_ROW)

    from_date = str(metadata_row.get("from_date", "")).strip()
    to_date   = str(metadata_row.get("to_date", "")).strip()
    week_no   = str(metadata_row.get("week_no", "")).strip()
    week_range = f"{from_date} to {to_date}" if from_date and to_date else ""

    if not allow_duplicate_week and _is_duplicate_row(cols_ab, week_no, week_range, require_nis=True):
        print(
            f"[DEDUP] Row with Week No={week_no} and Week='{week_range}' "
            f"already exists on sheet '{resolved_sheet}' – skipping insert."
        )
        return

    # ── Find first blank row in column A ─────────────────────────────────
    # Extract column A from the A+B data we already fetched
    col_vals = [[row[0]] if isinstance(row, list) and len(row) > 0 else [None] for row in cols_ab]
    target_row = find_first_blank_row_in_A(col_vals, START_DATA_ROW, MAX_SCAN_ROW)

    row_values = build_row_values(metadata_row, spv_dsm_row,
                                   checklist_id=checklist_id,
                                   sap_checklist_number=sap_checklist_number,
                                   email_received_date=email_received_date,
                                   nis_booking_date=nis_booking_date)

    if len(row_values) != EXPECTED_COLS:
        raise ValueError(
            f"row_values must have exactly {EXPECTED_COLS} items (A..L); got {len(row_values)}."
        )

    write_row_A_to_M(drive_id, item_id, resolved_sheet, target_row, row_values, headers_auth)
    print(f"[INFO] Inserted WEEK NO={metadata_row.get('week_no', '?')} at row {target_row} on sheet '{resolved_sheet}' in {OUTPUT_FILE_NAME}")


def _write_cell(drive_id, item_id, sheet_name, cell_address, value, headers_auth):
    """Patch a single cell (e.g. 'K7') so we can merge into an existing row
    without rewriting the whole A-L range."""
    url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets('{sheet_name}')/range(address='{cell_address}')"
    )
    patch_json(url, headers_auth, {"values": [[value]]})
    print(f"[OK] Updated cell {cell_address} = {value!r}")


def _build_reply_row_values(metadata_row, sap_checklist_number, email_received_date, cn_dn_date):
    """12-value A-L row for a Debit/Credit PDF-reply SAP checklist.

    These rows exist only to record a SAP checklist created from a finance
    reply PDF, so the NIS number (col I) and the amount columns (E-H, incl.
    Revenue Diff) are intentionally left blank. Only the shared header fields
    (week, dates, due date) and the SAP checklist number (col K) are filled.
    """
    from_date = str(metadata_row.get("from_date", "")).strip()
    to_date   = str(metadata_row.get("to_date", "")).strip()
    week_range = f"{from_date} to {to_date}" if from_date and to_date else ""
    return [
        str(metadata_row.get("week_no", "")).strip(),                      # A Week No
        week_range,                                                        # B Week
        str(metadata_row.get("dsm_statement_published_date", "")).strip(), # C Statement date
        email_received_date or "",                                         # D Working Received Date
        "",                                                                # E DSM Charges (blank)
        "",                                                                # F Revenue Diff (blank)
        "",                                                                # G Drawl Charges (blank)
        "",                                                                # H Total (blank)
        "",                                                                # I NIS No (blank — reply row)
        str(metadata_row.get("due_date", "")).strip(),                     # J Due Date
        sap_checklist_number or "",                                        # K CN/DN Checklist No
        cn_dn_date or "",                                                  # L CN/DN Date
    ]


def update_sap_checklist_on_existing_row(
    metadata_row: dict,
    sheet_name: str,
    sap_checklist_number: str,
):
    """Update column K (CN/DN Checklist No) on an EXISTING NIS row for this week/SPV.

    When SAP automation completes, the NIS booking row already exists in the
    DSM-NIS-OUTPUT Excel (written during NIS booking). This function finds that
    row by matching Week (col B) and merges the SAP checklist number into
    column K (comma-separated, de-duplicated).

    If no existing NIS row is found (shouldn't happen), falls back to writing
    a new row via add_incremental_week_row.
    """
    sap_number = str(sap_checklist_number or "").strip()
    if not sap_number:
        print("[SAP-UPDATE] No SAP checklist number given — nothing to update.")
        return

    token = get_token()
    headers_auth = {"Authorization": f"Bearer {token}"}
    _, drive_id = resolve_site_and_drive(headers_auth)
    item_path = f"{FOLDER_PATH}/{OUTPUT_FILE_NAME}".strip("/")
    item_id = resolve_item_id_by_path(drive_id, item_path, headers_auth)
    resolved_sheet = ensure_sheet_name(drive_id, item_id, headers_auth, sheet_name)

    from_date = str(metadata_row.get("from_date", "")).strip()
    to_date = str(metadata_row.get("to_date", "")).strip()
    week_range = f"{from_date} to {to_date}" if from_date and to_date else ""

    rows = _read_cols_A_to_L(drive_id, item_id, resolved_sheet, headers_auth, max_row=MAX_SCAN_ROW)

    # Find the existing NIS row for this week (col B matches, col I has NIS number)
    for idx, row in enumerate(rows):
        if idx == 0 or not isinstance(row, list):
            continue
        cell_b = str(row[1]).strip().lower() if len(row) > 1 and row[1] is not None else ""
        cell_i = str(row[8]).strip() if len(row) > 8 and row[8] is not None else ""
        if cell_b == week_range.lower() and cell_i:
            # Found the NIS row — merge SAP number into column K
            existing_k = str(row[10]).strip() if len(row) > 10 and row[10] is not None else ""
            merged = merge_checklist_numbers(existing_k, sap_number)
            if merged == existing_k:
                print(f"[SAP-UPDATE] {sap_number} already present in col K "
                      f"on sheet '{resolved_sheet}' row {idx + 1} — no change.")
                return
            _write_cell(drive_id, item_id, resolved_sheet, f"K{idx + 1}", merged, headers_auth)
            print(f"[SAP-UPDATE] Updated col K on existing NIS row {idx + 1} "
                  f"sheet '{resolved_sheet}': K='{merged}'")
            return

    # Fallback: no NIS row found — try a broader search (any row with matching week)
    for idx, row in enumerate(rows):
        if idx == 0 or not isinstance(row, list):
            continue
        cell_b = str(row[1]).strip().lower() if len(row) > 1 and row[1] is not None else ""
        if cell_b == week_range.lower():
            existing_k = str(row[10]).strip() if len(row) > 10 and row[10] is not None else ""
            merged = merge_checklist_numbers(existing_k, sap_number)
            if merged == existing_k:
                print(f"[SAP-UPDATE] {sap_number} already present in col K "
                      f"on sheet '{resolved_sheet}' row {idx + 1} — no change.")
                return
            _write_cell(drive_id, item_id, resolved_sheet, f"K{idx + 1}", merged, headers_auth)
            print(f"[SAP-UPDATE] Updated col K on row {idx + 1} "
                  f"sheet '{resolved_sheet}': K='{merged}' (fallback match)")
            return

    # No row found at all — this shouldn't happen but log it
    print(f"[SAP-UPDATE] ⚠ No existing row found for week '{week_range}' on "
          f"sheet '{resolved_sheet}' — SAP checklist number not written to tracker.")


def add_sap_checklist_reply_row(
    metadata_row: dict,
    sheet_name: str,
    sap_checklist_number: str,
    email_received_date: str = "",
    cn_dn_date: str = "",
):
    """Record a SAP checklist created from a finance-reply Debit/Credit PDF.

    Per the DSM-NIS-OUTPUT rules, a reply PDF produces a NEW row on the SPV's
    sheet that carries NO NIS booking number — only the SAP checklist number in
    column K plus the shared header fields copied from the original email.

    Idempotent + preservation-safe: if a reply row (a data row with a blank NIS
    column I) already exists for the same week, the new SAP checklist number is
    MERGED into that row's column K (de-duplicated, existing numbers kept)
    instead of adding a duplicate row. This makes SAP retries safe and lets one
    SPV/week accumulate multiple comma-separated SAP checklist numbers.
    """
    sap_number = str(sap_checklist_number or "").strip()
    if not sap_number:
        print("[SAP-REPLY] No SAP checklist number given — nothing to write.")
        return

    token = get_token()
    headers_auth = {"Authorization": f"Bearer {token}"}
    _, drive_id = resolve_site_and_drive(headers_auth)
    item_path   = f"{FOLDER_PATH}/{OUTPUT_FILE_NAME}".strip("/")
    item_id     = resolve_item_id_by_path(drive_id, item_path, headers_auth)
    resolved_sheet = ensure_sheet_name(drive_id, item_id, headers_auth, sheet_name)

    from_date = str(metadata_row.get("from_date", "")).strip()
    to_date   = str(metadata_row.get("to_date", "")).strip()
    week_range = f"{from_date} to {to_date}" if from_date and to_date else ""

    rows = _read_cols_A_to_L(drive_id, item_id, resolved_sheet, headers_auth, max_row=MAX_SCAN_ROW)

    # ── Look for an existing reply row for this week (col I / NIS blank) ──
    for idx, row in enumerate(rows):
        if idx == 0 or not isinstance(row, list):   # skip header
            continue
        cell_b = str(row[1]).strip().lower() if len(row) > 1 and row[1] is not None else ""
        cell_i = str(row[8]).strip() if len(row) > 8 and row[8] is not None else ""
        if cell_b == week_range.lower() and not cell_i:
            existing_k = str(row[10]).strip() if len(row) > 10 and row[10] is not None else ""
            merged = merge_checklist_numbers(existing_k, sap_number)
            if merged == existing_k:
                print(f"[SAP-REPLY] {sap_number} already present on sheet "
                      f"'{resolved_sheet}' (row {idx + 1}) — nothing to add.")
                return
            _write_cell(drive_id, item_id, resolved_sheet, f"K{idx + 1}", merged, headers_auth)
            print(f"[SAP-REPLY] Merged SAP checklist into existing reply row "
                  f"{idx + 1} on '{resolved_sheet}': K='{merged}'")
            return

    # ── No reply row yet → insert a new one (NIS blank, K = SAP number) ──
    col_vals = [[r[0]] if isinstance(r, list) and len(r) > 0 else [None] for r in rows]
    target_row = find_first_blank_row_in_A(col_vals, START_DATA_ROW, MAX_SCAN_ROW)
    row_values = _build_reply_row_values(metadata_row, sap_number, email_received_date, cn_dn_date)
    if len(row_values) != EXPECTED_COLS:
        raise ValueError(
            f"reply row_values must have exactly {EXPECTED_COLS} items (A..L); got {len(row_values)}."
        )
    write_row_A_to_M(drive_id, item_id, resolved_sheet, target_row, row_values, headers_auth)
    print(f"[SAP-REPLY] Inserted SAP-checklist reply row at {target_row} on "
          f"'{resolved_sheet}' (week='{week_range}', K='{sap_number}', NIS blank).")


# =========================
# STANDALONE ENTRY POINT
# =========================
def main():
    """
    Standalone test: uses hard-coded example data so the script can be run
    independently without the full email-reading pipeline.
    """
    print("\n=== Insert Row (A..M) into DSM-NIS-OUTPUT – standalone test ===")
    print("TENANT_HOST      :", TENANT_HOST)
    print("SITE_NAME        :", SITE_NAME)
    print("DRIVE_NAME       :", DRIVE_NAME)
    print("FOLDER_PATH      :", FOLDER_PATH)
    print("OUTPUT_FILE_NAME :", OUTPUT_FILE_NAME)
    print("=======================================================================\n")

    # Example data matching the email parser output format
    example_metadata = {
        "week_no":                       "34",
        "from_date":                     "17-11-2025",
        "to_date":                       "23-11-2025",
        "dsm_statement_published_date":  "2-Dec-25",
        "due_date":                      "12-Dec-25",
    }

    example_spv_dsm = {
        "spv_name":                  "AWEMP1PL",
        "total_dsm_charges_payable":  1016,
        "drawl_charges_payable":      0,
        "revenue_diff":               4517646,
        "revenue_loss":               0,
        "net_payable_receivable":     4518662.00,
    }

    add_incremental_week_row(
        metadata_row=example_metadata,
        spv_dsm_row=example_spv_dsm,
        sheet_name="AWEMP1PL",                    # SPV name = sheet name
        checklist_id="NIS-2025-0001234",           # example NIS checklist ID
        sap_checklist_number="SAP-CKL-0005678",    # example SAP checklist number
        email_received_date="08-Apr-2026",          # example email received date
        nis_booking_date="08-Apr-2026",             # example NIS booking date
    )


if __name__ == "__main__":
    main()