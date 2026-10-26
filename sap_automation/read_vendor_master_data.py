import json
import re
import requests
import msal
import sap_config as config

# All credentials and SharePoint settings loaded from config.py
TENANT_ID     = config.SP_TENANT_ID
CLIENT_ID     = config.SP_CLIENT_ID
CLIENT_SECRET = config.SP_CLIENT_SECRET

TENANT_HOST = config.TENANT_HOST
SITE_NAME   = config.SITE_NAME
DRIVE_NAME  = config.DRIVE_NAME
FOLDER_PATH = config.FOLDER_PATH

# ── Input file settings ──────────────────────────────────────────────────
INPUT_FILE_NAME  = config.INPUT_FILE_NAME    # DSM-NIS-INPUT.xlsx
INPUT_SHEET_NAME = config.INPUT_SHEET_NAME   # "main"

AUTHORITY = config.SP_AUTHORITY
SCOPES    = config.GRAPH_SCOPE


def get_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )
    return app.acquire_token_for_client(scopes=SCOPES)["access_token"]


def get_json(url, headers):
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.json()


def _extract_spv_code(name: str) -> str:
    """Extract the core SPV code, ignoring MW/capacity suffixes and whitespace.

    Examples:
        'AWEK06L-75 MW'     → 'AWEK06L'
        'AWEK06L 75MW'      → 'AWEK06L'
        'AWEMP1PL-Solar AWEMP1PL 700MW' → 'AWEMP1PL'
        'AHEJ4L'            → 'AHEJ4L'
    """
    name = str(name or "").strip()
    if not name:
        return ""
    # Remove common suffixes: dash/space + digits + MW/KW/GW
    cleaned = re.sub(r"[\s\-_]+\d+\s*(MW|KW|GW)\b.*$", "", name, flags=re.IGNORECASE)
    # Also remove trailing " MW", "MW" alone
    cleaned = re.sub(r"\s+\d+\s*$", "", cleaned)
    # Take only the first token (the plant code) if there are multiple parts
    tokens = re.split(r"[\s\-_]+", cleaned.strip())
    return tokens[0].strip() if tokens else cleaned.strip()


def read_vendor_data(spv_name: str) -> dict:
    """
    Read vendor/SPV master data from DSM-NIS-INPUT.xlsx → 'main' sheet.

    The sheet has columns (row 1 = header, data from row 2):
        A: SPV Name
        B: Vendor
        C: company
        D: cost center
        E: Plant
        F: Bank Key
        G: GL

    Finds the row whose column A matches *spv_name* (case-insensitive) and
    returns a dict like:
        {
            "Vendor":      "...",
            "Company":     "...",
            "Cost Center": "...",
            "Plant":       "...",
            "Bank Key":    "...",
            "GL":          "...",
        }

    Raises RuntimeError if the SPV is not found.
    """
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}

    # Resolve site and drive
    site = get_json(
        f"https://graph.microsoft.com/v1.0/sites/{TENANT_HOST}:/sites/{SITE_NAME}",
        headers
    )
    site_id = site["id"]

    drives = get_json(
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives",
        headers
    )["value"]

    drive_id = next(d["id"] for d in drives if d["name"] == DRIVE_NAME)

    # Resolve the INPUT file
    item = get_json(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{FOLDER_PATH}/{INPUT_FILE_NAME}",
        headers
    )

    # Read all data rows from A2:G5000 (row 1 is header)
    range_url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item['id']}"
        f"/workbook/worksheets('{INPUT_SHEET_NAME}')/range(address='A2:G5000')"
    )

    resp = get_json(range_url, headers)
    rows = resp.get("values", [])

    # Column indices: A=0(SPV Name), B=1(Vendor), C=2(company),
    #                 D=3(cost center), E=4(Plant), F=5(Bank Key), G=6(GL)
    target = spv_name.strip().upper()

    # 1) Exact full-name match
    for row in rows:
        if len(row) < 7:
            continue
        cell_spv = str(row[0] or "").strip().upper()
        if not cell_spv:
            continue  # skip blank rows
        if cell_spv == target:
            return {
                "Vendor":      str(row[1] or "").strip(),
                "Company":     str(row[2] or "").strip(),
                "Cost Center": str(row[3] or "").strip(),
                "Plant":       str(row[4] or "").strip(),
                "Bank Key":    str(row[5] or "").strip(),
                "GL":          str(row[6] or "").strip(),
            }

    # 2) Match by core SPV code (strip MW suffix, whitespace, etc.)
    target_code = _extract_spv_code(spv_name).upper()
    if target_code:
        for row in rows:
            if len(row) < 7:
                continue
            cell_spv = str(row[0] or "").strip()
            if not cell_spv:
                continue
            row_code = _extract_spv_code(cell_spv).upper()
            if row_code == target_code:
                print(f"  [VENDOR] Matched by core SPV code: '{spv_name}' → "
                      f"'{cell_spv}' (code={target_code})")
                return {
                    "Vendor":      str(row[1] or "").strip(),
                    "Company":     str(row[2] or "").strip(),
                    "Cost Center": str(row[3] or "").strip(),
                    "Plant":       str(row[4] or "").strip(),
                    "Bank Key":    str(row[5] or "").strip(),
                    "GL":          str(row[6] or "").strip(),
                }

    raise RuntimeError(
        f"SPV '{spv_name}' not found in {INPUT_FILE_NAME} / sheet '{INPUT_SHEET_NAME}'. "
        f"Checked {len(rows)} data row(s). Core code tried: '{target_code}'."
    )


# ─────────────────────────────────────────────────────────────────────────────
# Debit/Credit reply-PDF vendor resolution (SPV from filename → DSM-NIS-INPUT)
# ─────────────────────────────────────────────────────────────────────────────
import re
import difflib

_INPUT_ROWS_CACHE = None


def _load_input_rows() -> list:
    """Load all DSM-NIS-INPUT rows as dicts (cached for the process lifetime).

    Returns rows like:
        {"SPV Name": ..., "Vendor": ..., "Company": ..., "Cost Center": ...,
         "Plant": ..., "Bank Key": ..., "GL": ...}
    """
    global _INPUT_ROWS_CACHE
    if _INPUT_ROWS_CACHE is not None:
        return _INPUT_ROWS_CACHE

    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    site = get_json(
        f"https://graph.microsoft.com/v1.0/sites/{TENANT_HOST}:/sites/{SITE_NAME}", headers)
    site_id = site["id"]
    drives = get_json(
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives", headers)["value"]
    drive_id = next(d["id"] for d in drives if d["name"] == DRIVE_NAME)
    item = get_json(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{FOLDER_PATH}/{INPUT_FILE_NAME}",
        headers)
    range_url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item['id']}"
        f"/workbook/worksheets('{INPUT_SHEET_NAME}')/range(address='A2:G5000')"
    )
    resp = get_json(range_url, headers)
    raw = resp.get("values", []) or []

    rows = []
    for row in raw:
        if len(row) < 7:
            continue
        spv = str(row[0] or "").strip()
        if not spv:
            continue
        rows.append({
            "SPV Name":    spv,
            "Vendor":      str(row[1] or "").strip(),
            "Company":     str(row[2] or "").strip(),
            "Cost Center": str(row[3] or "").strip(),
            "Plant":       str(row[4] or "").strip(),
            "Bank Key":    str(row[5] or "").strip(),
            "GL":          str(row[6] or "").strip(),
        })
    _INPUT_ROWS_CACHE = rows
    return rows


def _normalize_spv(name) -> str:
    return re.sub(r"\s+", " ", str(name or "").strip()).upper()


def _row_fields(row: dict) -> dict:
    return {
        "Vendor":      str(row.get("Vendor") or "").strip(),
        "Company":     str(row.get("Company") or "").strip(),
        "Cost Center": str(row.get("Cost Center") or "").strip(),
        "Plant":       str(row.get("Plant") or "").strip(),
        "Bank Key":    str(row.get("Bank Key") or "").strip(),
        "GL":          str(row.get("GL") or "").strip(),
    }


def spv_name_from_pdf_filename(filename: str) -> str:
    """Derive the SPV name from a debit/credit-note PDF filename.

    Examples:
        '1.AWEK06L.pdf'    → 'AWEK06L'
        'DSM_3.AWEK5L.pdf' → 'AWEK5L'
    Strips any directory, a leading numeric index, and the extension.
    """
    base = str(filename or "").replace("\\", "/").split("/")[-1]
    base = re.sub(r"\.pdf$", "", base, flags=re.IGNORECASE)
    base = re.sub(r"^DSM_", "", base, flags=re.IGNORECASE)
    base = re.sub(r"^\s*\d+\s*[\.\)\-_]\s*", "", base)
    return base.strip()


def find_closest_spv(spv_name: str, rows: list = None) -> dict:
    """Return the input row whose SPV Name is the closest match to *spv_name*."""
    if rows is None:
        rows = _load_input_rows()
    if not rows:
        return {}
    target = _normalize_spv(spv_name)
    if not target:
        return {}

    exact, prefix, contains = [], [], []
    for row in rows:
        norm = _normalize_spv(row.get("SPV Name"))
        if not norm:
            continue
        if norm == target:
            exact.append(row)
        elif norm.startswith(target):
            prefix.append(row)
        elif target in norm:
            contains.append(row)

    def _best(cands):
        if len(cands) == 1:
            return cands[0]
        return sorted(
            cands,
            key=lambda r: (
                len(_normalize_spv(r.get("SPV Name"))),
                -difflib.SequenceMatcher(None, target, _normalize_spv(r.get("SPV Name"))).ratio(),
            ),
        )[0]

    for bucket in (exact, prefix, contains):
        if bucket:
            return _best(bucket)

    scored = sorted(
        rows,
        key=lambda r: difflib.SequenceMatcher(None, target, _normalize_spv(r.get("SPV Name"))).ratio(),
        reverse=True,
    )
    best = scored[0]
    ratio = difflib.SequenceMatcher(None, target, _normalize_spv(best.get("SPV Name"))).ratio()
    return best if ratio >= 0.5 else {}


def _keywords_from_filename(filename: str) -> list:
    core = spv_name_from_pdf_filename(filename)
    return [t for t in re.split(r"[^A-Za-z0-9]+", core.upper()) if t]


def find_spv_by_keywords(filename_or_name: str, rows: list = None) -> dict:
    """Find the SPV row by keyword-matching the PDF filename to the SPV Name
    column (primary plant-code keyword first), falling back to fuzzy match."""
    if rows is None:
        rows = _load_input_rows()
    if not rows:
        return {}
    keywords = _keywords_from_filename(filename_or_name)
    if not keywords:
        return {}
    primary = keywords[0]
    norm_primary = _normalize_spv(primary)
    for row in rows:
        norm = _normalize_spv(row.get("SPV Name"))
        if norm_primary and norm_primary in norm:
            return row
    return find_closest_spv(primary, rows)


def read_vendor_data_by_pdf_filename(filename: str) -> dict:
    """Resolve vendor data for a Debit/Credit reply PDF by its filename.

    Parses the SPV from the filename, keyword-matches it to DSM-NIS-INPUT, and
    returns the vendor fields plus the matched SPV name. Returns {} if unmatched.
    """
    row = find_spv_by_keywords(filename)
    if not row:
        return {}
    result = _row_fields(row)
    result["Matched SPV"] = str(row.get("SPV Name") or "").strip()
    result["SPV From Filename"] = spv_name_from_pdf_filename(filename)
    return result
