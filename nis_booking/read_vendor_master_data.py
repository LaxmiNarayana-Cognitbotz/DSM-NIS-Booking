import json
import re
import difflib
import requests
import msal
import config

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
    # e.g. "AWEMP1PL-Solar AWEMP1PL 700MW" → split on dash/space → "AWEMP1PL"
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

    Matching strategy:
      1. Exact case-insensitive match on full SPV Name.
      2. Match by core SPV code (ignoring MW suffix / whitespace differences).
      3. Falls through to find_closest_spv() in the caller if both fail.

    Returns a dict like:
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
    rows = _load_input_rows()
    target = spv_name.strip().upper()

    # 1) Exact full-name match
    for row in rows:
        if str(row.get("SPV Name") or "").strip().upper() == target:
            return _row_fields(row)

    # 2) Match by core SPV code (strip MW suffix, whitespace, etc.)
    target_code = _extract_spv_code(spv_name).upper()
    if target_code:
        for row in rows:
            row_code = _extract_spv_code(row.get("SPV Name")).upper()
            if row_code == target_code:
                print(f"  [VENDOR] Matched by core SPV code: '{spv_name}' → "
                      f"'{row.get('SPV Name')}' (code={target_code})")
                return _row_fields(row)

    raise RuntimeError(
        f"SPV '{spv_name}' not found in {INPUT_FILE_NAME} / sheet '{INPUT_SHEET_NAME}'. "
        f"Checked {len(rows)} data row(s). Core code tried: '{target_code}'."
    )


def _row_fields(row: dict) -> dict:
    """Project a raw input row dict to the vendor-data shape used by callers."""
    return {
        "Vendor":      str(row.get("Vendor") or "").strip(),
        "Company":     str(row.get("Company") or "").strip(),
        "Cost Center": str(row.get("Cost Center") or "").strip(),
        "Plant":       str(row.get("Plant") or "").strip(),
        "Bank Key":    str(row.get("Bank Key") or "").strip(),
        "GL":          str(row.get("GL") or "").strip(),
    }


def _load_input_rows() -> list:
    """Fetch every data row from DSM-NIS-INPUT.xlsx / 'main' as a list of dicts.

    Each dict carries the SPV name plus vendor fields. Blank SPV rows are
    skipped. This is the single Graph fetch shared by the exact and
    closest-name lookups.
    """
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}

    site = get_json(
        f"https://graph.microsoft.com/v1.0/sites/{TENANT_HOST}:/sites/{SITE_NAME}",
        headers,
    )
    site_id = site["id"]

    drives = get_json(
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives",
        headers,
    )["value"]
    drive_id = next(d["id"] for d in drives if d["name"] == DRIVE_NAME)

    item = get_json(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{FOLDER_PATH}/{INPUT_FILE_NAME}",
        headers,
    )

    range_url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item['id']}"
        f"/workbook/worksheets('{INPUT_SHEET_NAME}')/range(address='A2:G5000')"
    )
    resp = get_json(range_url, headers)
    rows = resp.get("values", [])

    out = []
    for row in rows:
        if len(row) < 7:
            continue
        spv = str(row[0] or "").strip()
        if not spv:
            continue
        out.append({
            "SPV Name":    spv,
            "Vendor":      row[1],
            "Company":     row[2],
            "Cost Center": row[3],
            "Plant":       row[4],
            "Bank Key":    row[5],
            "GL":          row[6],
        })
    return out


def _normalize_spv(name: str) -> str:
    """Uppercase and strip all non-alphanumerics for robust name comparison."""
    return re.sub(r"[^A-Z0-9]", "", str(name or "").upper())


def spv_name_from_pdf_filename(filename: str) -> str:
    """Derive the SPV name from a debit/credit-note PDF filename.

    Examples:
        '1.AWEK06L.pdf'      → 'AWEK06L'
        '2.AWEK3L.pdf'       → 'AWEK3L'
        'DSM_3.AWEK5L.pdf'   → 'AWEK5L'
        'AWEK01L II.pdf'     → 'AWEK01L II'
    Strips any directory, a leading numeric index ('1.', '02 -', 'DSM_1.'),
    and the extension.
    """
    base = str(filename or "").replace("\\", "/").split("/")[-1]
    base = re.sub(r"\.pdf$", "", base, flags=re.IGNORECASE)
    base = re.sub(r"^DSM_", "", base, flags=re.IGNORECASE)
    # Remove a leading index like "1.", "01 -", "3_", "2)"
    base = re.sub(r"^\s*\d+\s*[\.\)\-_]\s*", "", base)
    return base.strip()


def find_closest_spv(spv_name: str, rows: list = None) -> dict:
    """Return the input row whose SPV Name is the closest match to *spv_name*.

    Matching strategy (best first):
      1. exact (normalized) match,
      2. input SPV starts with the target (e.g. 'AWEK06L' → 'AWEK06L_MSEDCL_75 MW'),
      3. target contained in input SPV,
      4. difflib closest ratio (tie-break / fallback).
    Returns {} when no reasonable match is found.
    """
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
        # Prefer the shortest normalized name (closest to the bare SPV code),
        # then the highest difflib ratio.
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

    # Fallback: overall closest by ratio, with a sane floor.
    scored = sorted(
        rows,
        key=lambda r: difflib.SequenceMatcher(None, target, _normalize_spv(r.get("SPV Name"))).ratio(),
        reverse=True,
    )
    best = scored[0]
    ratio = difflib.SequenceMatcher(None, target, _normalize_spv(best.get("SPV Name"))).ratio()
    return best if ratio >= 0.5 else {}


def read_vendor_data_by_closest_name(spv_name: str) -> dict:
    """Vendor data for the SPV whose name best matches *spv_name* (fuzzy).

    Used for finance-reply Debit/Credit PDFs, where the SPV comes from the PDF
    filename and may be a short code (e.g. 'AWEK06L') that must be matched to a
    fuller DSM-NIS-INPUT key (e.g. 'AWEK06L_MSEDCL_75 MW').
    Returns the vendor-data dict, or {} if nothing matches.
    """
    row = find_closest_spv(spv_name)
    if not row:
        return {}
    result = _row_fields(row)
    result["Matched SPV"] = str(row.get("SPV Name") or "").strip()
    return result


def _keywords_from_filename(filename: str) -> list:
    """Break a Debit/Credit PDF filename into comparison keywords.

    e.g. '1.AWEK06L_50 MW.pdf' → ['AWEK06L', '50', 'MW']  (index + ext removed).
    The leading plant code (e.g. 'AWEK06L') is always first so it is treated as
    the primary keyword.
    """
    core = spv_name_from_pdf_filename(filename)
    tokens = [t for t in re.split(r"[^A-Za-z0-9]+", core.upper()) if t]
    return tokens


def find_spv_by_keywords(filename_or_name: str, rows: list = None) -> dict:
    """Find the SPV row by keyword-matching the PDF filename to the SPV Name
    column, returning the FIRST matching row (per spec).

    The filename may not contain the exact SPV name (it can include extra tokens
    like '50 MW'), so we match on the leading plant-code keyword and prefer rows
    that share the most keywords. Falls back to the fuzzy closest match.
    """
    if rows is None:
        rows = _load_input_rows()
    if not rows:
        return {}

    keywords = _keywords_from_filename(filename_or_name)
    if not keywords:
        return {}
    primary = keywords[0]  # the plant code, e.g. AWEK06L

    # First row whose SPV Name contains the primary plant-code keyword.
    for row in rows:
        norm = _normalize_spv(row.get("SPV Name"))
        if _normalize_spv(primary) and _normalize_spv(primary) in norm:
            return row

    # No direct plant-code hit → fall back to the fuzzy closest match.
    return find_closest_spv(primary, rows)


def read_vendor_data_by_pdf_filename(filename: str) -> dict:
    """Resolve vendor data for a Debit/Credit reply PDF by its filename.

    Parses the SPV from the filename, keyword-matches it to DSM-NIS-INPUT, and
    returns the vendor fields (Vendor code, Company, Cost Center, Plant, …) of
    the first matching row, plus the matched SPV name. Returns {} if unmatched.
    """
    row = find_spv_by_keywords(filename)
    if not row:
        return {}
    result = _row_fields(row)
    result["Matched SPV"] = str(row.get("SPV Name") or "").strip()
    result["SPV From Filename"] = spv_name_from_pdf_filename(filename)
    return result

