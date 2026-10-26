import json
import re
from html import unescape
from typing import List, Dict, Optional

# --------------------------------------------------
# Header maps  –  key = any normalized substring that
# uniquely identifies the column.
# Multiple keys can map to the same field (aliases).
# The LONGEST matching key wins (avoids "date" matching
# both "from date" and "to date" ambiguously).
# --------------------------------------------------
METADATA_HEADER_MAP = {
    # ── week number ────────────────────────────────
    "week no":          "week_no",
    "week number":      "week_no",
    "weekno":           "week_no",

    # ── period start ───────────────────────────────
    "from date":        "from_date",
    "date from":        "from_date",
    "start date":       "from_date",
    "week from":        "from_date",
    "period from":      "from_date",

    # ── period end ─────────────────────────────────
    "to date":          "to_date",
    "date to":          "to_date",
    "end date":         "to_date",
    "week to":          "to_date",
    "period to":        "to_date",
    "week end":         "to_date",

    # ── published date ──────────────────────────────
    "dsm statement published date": "dsm_statement_published_date",
    "statement published":          "dsm_statement_published_date",
    "published date":               "dsm_statement_published_date",

    # ── due date ────────────────────────────────────
    "due date":         "due_date",
    "payment due":      "due_date",
}

SPV_DSM_HEADER_MAP = {
    "spv name":                     "spv_name",
    "spvname":                      "spv_name",
    "name of spv":                  "spv_name",
    "spv":                          "spv_name",
    # ── plant name is an alias for spv name ────────
    # Some DSM statements label the identifier column "Plant Name"
    # (or just "Plant") instead of "SPV Name".  The plant name value
    # is identical to the SPV name, so map it to spv_name.
    "plant name":                   "spv_name",
    "plant":                        "spv_name",
    "total dsm charges payable":    "total_dsm_charges_payable",
    "dsm charges payable":          "total_dsm_charges_payable",
    "dsm charges":                  "total_dsm_charges_payable",
    "drawl charges payable":        "drawl_charges_payable",
    "drawl charges":                "drawl_charges_payable",
    "revenue diff":                 "revenue_diff",
    "revenue difference":           "revenue_diff",
    "revenue loss":                 "revenue_loss",
    "net payable":                  "net_payable_receivable",
    "net receivable":               "net_payable_receivable",
    "net amount":                   "net_payable_receivable",
}

# --------------------------------------------------
# Utility functions
# --------------------------------------------------
def normalize(text: str) -> str:
    text = unescape(text.lower())
    # ── Step 1: expand meaningful parentheticals BEFORE stripping ────────────
    # "From (Date)" → "from date", "To (Date)" → "to date"
    # Without this, the blanket strip below turns them into bare "from"/"to"
    # which never match the header-map keys "from date" / "to date".
    text = re.sub(r"\(\s*date\s*\)", " date", text)
    # ── Step 2: strip only short label/unit codes like (A), (B), (Rs.), (In) ─
    # Limit to ≤ 10 chars inside parens so long descriptive phrases survive.

# --------------------------------------------------
# Utility functions
# --------------------------------------------------
def normalize(text: str) -> str:
    text = unescape(text.lower())
    # ── Step 1: expand meaningful parentheticals BEFORE stripping ────────────
    # "From (Date)" → "from date", "To (Date)" → "to date"
    # Without this, the blanket strip below turns them into bare "from"/"to"
    # which never match the header-map keys "from date" / "to date".
    text = re.sub(r"\(\s*date\s*\)", " date", text)
    # ── Step 2: strip only short label/unit codes like (A), (B), (Rs.), (In) ─
    # Limit to ≤ 10 chars inside parens so long descriptive phrases survive.
    text = re.sub(r"\([^)]{0,10}\)", " ", text)
    # ── Step 3: remove remaining punctuation, collapse whitespace ─────────────
    text = re.sub(r"[^a-z0-9 ]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()

def clean_cell(text: str) -> str:
    if not text:
        return ""
    text = unescape(text)
    text = re.sub(r"<br\s*/?>", " ", text, flags=re.I)
    text = re.sub(r"<[^>]+>", "", text)
    return text.strip()

# Set of extracted metadata fields that contain date values.
# normalize_date_cell() is applied to these fields in build_table().
_DATE_FIELDS = {
    "from_date",
    "to_date",
    "due_date",
    "dsm_statement_published_date",
}

# All strptime formats tried in order.  We list the most common first;
# full month name (e.g. "July") and abbreviated (e.g. "Jul") are both
# covered so variants like "14 Jul 2025", "14-Jul-25", "July 14 2025",
# "14/07/2025", "14-07-2025" all parse correctly.
_DATE_FORMATS = [
    # Numeric-only ─────────────────────────────────────────────────────────────
    "%d-%m-%Y",   # 14-07-2025
    "%d/%m/%Y",   # 14/07/2025
    "%d.%m.%Y",   # 14.07.2025
    "%d-%m-%y",   # 14-07-25
    "%d/%m/%y",   # 14/07/25
    "%Y-%m-%d",   # 2025-07-14  (ISO)
    "%m/%d/%Y",   # 07/14/2025  (US)
    "%m-%d-%Y",   # 07-14-2025  (US)
    # Abbreviated month name ───────────────────────────────────────────────────
    "%d %b %Y",   # 14 Jul 2025
    "%d-%b-%Y",   # 14-Jul-2025
    "%d/%b/%Y",   # 14/Jul/2025
    "%d %b %y",   # 14 Jul 25
    "%d-%b-%y",   # 14-Jul-25
    "%b %d %Y",   # Jul 14 2025
    "%b %d, %Y",  # Jul 14, 2025
    "%b-%d-%Y",   # Jul-14-2025
    # Full month name ──────────────────────────────────────────────────────────
    "%d %B %Y",   # 14 July 2025
    "%d-%B-%Y",   # 14-July-2025
    "%d/%B/%Y",   # 14/July/2025
    "%d %B %y",   # 14 July 25
    "%B %d %Y",   # July 14 2025
    "%B %d, %Y",  # July 14, 2025
    "%B-%d-%Y",   # July-14-2025
]

def normalize_date_cell(raw: str) -> str:
    """
    Try to parse *raw* as a date using every format in _DATE_FORMATS and
    return a canonical ``DD-MM-YYYY`` string that the NIS portal expects.

    If the value cannot be parsed the original string is returned unchanged
    so that no data is silently lost.

    Examples
    --------
    >>> normalize_date_cell("14 Jul 2025")   → "14-07-2025"
    >>> normalize_date_cell("14-Jul-25")     → "14-07-2025"
    >>> normalize_date_cell("July 14 2025")  → "14-07-2025"
    >>> normalize_date_cell("14/07/2025")   → "14-07-2025"
    >>> normalize_date_cell("14-07-2025")   → "14-07-2025"  (unchanged)
    >>> normalize_date_cell("not a date")   → "not a date"  (unchanged)
    """
    from datetime import datetime as _dt
    raw = str(raw or "").strip()
    if not raw:
        return raw
    # Strip ordinal suffixes: "14th", "1st", "2nd", "3rd" → "14", "1", …
    cleaned = re.sub(r"(\d+)(st|nd|rd|th)\b", r"\1", raw, flags=re.I)
    # Collapse multiple whitespace / non-breaking spaces
    cleaned = re.sub(r"[\u00a0\s]+", " ", cleaned).strip()
    for fmt in _DATE_FORMATS:
        try:
            return _dt.strptime(cleaned, fmt).strftime("%d-%m-%Y")
        except ValueError:
            continue
    # Last-resort: if raw is already DD-MM-YYYY leave it as-is
    return raw

def extract_tables(html: str) -> List[str]:
    return re.findall(r"(?is)<table.*?</table>", html)

def extract_rows(table_html: str) -> List[str]:
    return re.findall(r"(?is)<tr.*?</tr>", table_html)

def extract_cells(row_html: str) -> List[str]:
    return re.findall(r"(?is)<t[dh].*?>(.*?)</t[dh]>", row_html)

# --------------------------------------------------
# Core table parser
# --------------------------------------------------
def parse_single_table(table_html: str) -> Optional[dict]:
    rows_html = extract_rows(table_html)
    rows = []

    for r in rows_html:
        cells = [clean_cell(c) for c in extract_cells(r)]
        if any(cells):
            rows.append(cells)

    if len(rows) < 2:
        return None

    headers = [normalize(h) for h in rows[0]]
    data_rows = rows[1:]

    return {
        "headers":     headers,
        "raw_headers": rows[0],
        "data":        data_rows,
    }

def match_headers(headers: List[str], header_map: Dict[str, str]) -> bool:
    return sum(1 for h in headers for k in header_map if k in h) >= 1


def header_match_score(headers: List[str], header_map: Dict[str, str]) -> int:
    """Count distinct mapped fields so the most complete table wins."""
    matched_fields = set()
    for header in headers:
        best_key = ""
        best_mapped = ""
        for key, mapped in header_map.items():
            if key in header and len(key) > len(best_key):
                best_key = key
                best_mapped = mapped
        if best_mapped:
            matched_fields.add(best_mapped)
    return len(matched_fields)

def build_table(parsed: dict, header_map: Dict[str, str]) -> List[Dict]:
    """
    Map column indices to field names using a longest-match-wins strategy:
    for each header cell, all matching keys are collected and the LONGEST
    key string is declared the winner.  This prevents short aliases like
    "date" or "spv" from stealing a column that a longer, more specific
    key (e.g. "from date", "spv name") also matches.

    Date fields (from_date, to_date, due_date, dsm_statement_published_date)
    are automatically normalised to DD-MM-YYYY via normalize_date_cell() so
    variants like "14 Jul 2025", "14-Jul-25", "14/07/2025" all become the
    same canonical format before being stored or sent to the NIS portal.
    """
    headers   = parsed["headers"]
    data_rows = parsed["data"]

    idx_map: Dict[int, str] = {}
    for i, h in enumerate(headers):
        best_key    = ""
        best_mapped = ""
        for key, mapped in header_map.items():
            if key in h and len(key) > len(best_key):
                best_key    = key
                best_mapped = mapped
        if best_mapped:
            idx_map[i] = best_mapped

    result = []
    for row in data_rows:
        record = {
            mapped: row[i] if i < len(row) else ""
            for i, mapped in idx_map.items()
        }
        if record.get("spv_name") and normalize(record.get("spv_name")) == "total":
            continue
        # Skip summary/sub-header rows where spv_name is blank or just
        # unit labels like "(A)", "(In Rs.)", "(B)" — these come from
        # multi-row headers with rowspan that the parser cannot flatten.
        # Only apply when spv_name is an expected field in this table.
        if "spv_name" in idx_map.values():
            spv_val = record.get("spv_name", "").strip()
            if not spv_val or re.match(r"^\(.*\)$", spv_val):
                continue
        # Normalise every date field to DD-MM-YYYY
        for field in _DATE_FIELDS:
            if field in record and record[field]:
                record[field] = normalize_date_cell(record[field])
        if any(v.strip() for v in record.values()):
            result.append(record)

    return result


def _visible_email_text(body_html: str) -> str:
    text = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", body_html or "")
    text = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</tr>", "\n", text)
    text = re.sub(r"(?s)<[^>]+>", " ", text)
    text = unescape(text)
    return re.sub(r"[ \t\r\f\v]+", " ", text).strip()


def _parse_llm_spv_names(reply: str, expected_count: int) -> List[str]:
    raw = str(reply or "").strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw, flags=re.I)
    raw = re.sub(r"\s*```$", "", raw)
    value = json.loads(raw)
    if isinstance(value, dict):
        value = value.get("spv_names")
    if not isinstance(value, list):
        raise ValueError("LLM response must contain an spv_names array")
    names = [str(name or "").strip() for name in value]
    if len(names) != expected_count or any(not name for name in names):
        raise ValueError(
            f"LLM returned {len(names)} usable positions; expected {expected_count}"
        )
    return names


def recover_missing_spv_names(
    body_html: str, spv_rows: List[Dict]
) -> List[Dict]:
    """Use the LLM only when deterministic parsing leaves an SPV name blank."""
    rows = [dict(row) for row in (spv_rows or [])]
    if not rows or all(str(row.get("spv_name") or "").strip() for row in rows):
        return rows

    from llm_curl_client import call_llm_curl

    visible_text = _visible_email_text(body_html)
    row_values = [
        {key: value for key, value in row.items() if key != "spv_name"}
        for row in rows
    ]
    messages = [
        {
            "role": "system",
            "content": (
                "Extract SPV/plant names from a DSM booking email. Return JSON only "
                'as {"spv_names":["name for row 1", "..."]}. Preserve the exact '
                "name, capacity, and row order shown in the DSM calculation table. "
                'Return "TOTAL" for an aggregate total row. Do not invent a name.'
            ),
        },
        {
            "role": "user",
            "content": (
                f"Expected DSM data rows: {len(rows)}\n"
                f"Already parsed row values: {json.dumps(row_values, ensure_ascii=False)}\n"
                f"Email text:\n{visible_text[:16000]}"
            ),
        },
    ]
    reply = call_llm_curl(messages, max_tokens=max(120, len(rows) * 40))
    names = _parse_llm_spv_names(reply, len(rows))
    for row, name in zip(rows, names):
        if not str(row.get("spv_name") or "").strip():
            row["spv_name"] = name
    print(f"[SPV-RECOVERY] LLM recovered {sum(1 for row in spv_rows if not str(row.get('spv_name') or '').strip())} missing SPV name(s)")
    return rows
# --------------------------------------------------
# Debug helper – prints raw extracted table structure
# --------------------------------------------------
def print_raw_tables(body_html: str) -> None:
    """
    Pretty-print every <table> found in the HTML so you can manually verify
    what the parser actually received before any mapping is applied.

    For each table this prints:
      1. The raw HTML snippet (truncated to 2 000 chars so the log stays readable)
      2. The raw cell text of EVERY row (no cleaning, no normalisation)
      3. The normalised headers after clean_cell + normalize()
      4. Up to 5 data rows after clean_cell()
    """
    tables = extract_tables(body_html)
    print(f"\n{'='*70}")
    print(f"[RAW EXTRACT] {len(tables)} table(s) found in email body")
    print(f"{'='*70}")

    for t_idx, t in enumerate(tables):
        print(f"\n---- Table[{t_idx}] raw HTML (first 2000 chars) ----")
        print(t[:2000])
        if len(t) > 2000:
            print(f"  ... [{len(t) - 2000} more chars truncated]")

        rows_html = extract_rows(t)
        print(f"\n---- Table[{t_idx}] raw cell text (NO cleaning, {len(rows_html)} row(s)) ----")
        for r_idx, row_html in enumerate(rows_html):
            # pull raw inner text without any clean_cell processing
            raw_cells = [
                re.sub(r"<[^>]+>", "", unescape(cell)).strip()
                for cell in re.findall(r"(?is)<t[dh].*?>(.*?)</t[dh]>", row_html)
            ]
            print(f"  Row[{r_idx}] raw  : {raw_cells}")

        parsed = parse_single_table(t)
        if not parsed:
            print(f"  (skipped – fewer than 2 usable rows after cleaning)")
            continue

        print(f"\n---- Table[{t_idx}] after clean_cell + normalize ----")
        print(f"  Raw  headers : {parsed['raw_headers']}")
        print(f"  Norm headers : {parsed['headers']}")
        for r_idx, row in enumerate(parsed["data"][:5]):
            print(f"  Data row[{r_idx}]  : {row}")

# --------------------------------------------------
# Debug helper – prints raw extracted table structure
# --------------------------------------------------
def print_raw_tables(body_html: str) -> None:
    """
    Pretty-print every <table> found in the HTML so you can manually verify
    what the parser actually received before any mapping is applied.

    For each table this prints:
      1. The raw HTML snippet (truncated to 2 000 chars so the log stays readable)
      2. The raw cell text of EVERY row (no cleaning, no normalisation)
      3. The normalised headers after clean_cell + normalize()
      4. Up to 5 data rows after clean_cell()
    """
    tables = extract_tables(body_html)
    print(f"\n{'='*70}")
    print(f"[RAW EXTRACT] {len(tables)} table(s) found in email body")
    print(f"{'='*70}")

    for t_idx, t in enumerate(tables):
        print(f"\n---- Table[{t_idx}] raw HTML (first 2000 chars) ----")
        print(t[:2000])
        if len(t) > 2000:
            print(f"  ... [{len(t) - 2000} more chars truncated]")

        rows_html = extract_rows(t)
        print(f"\n---- Table[{t_idx}] raw cell text (NO cleaning, {len(rows_html)} row(s)) ----")
        for r_idx, row_html in enumerate(rows_html):
            # pull raw inner text without any clean_cell processing
            raw_cells = [
                re.sub(r"<[^>]+>", "", unescape(cell)).strip()
                for cell in re.findall(r"(?is)<t[dh].*?>(.*?)</t[dh]>", row_html)
            ]
            print(f"  Row[{r_idx}] raw  : {raw_cells}")

        parsed = parse_single_table(t)
        if not parsed:
            print(f"  (skipped – fewer than 2 usable rows after cleaning)")
            continue

        print(f"\n---- Table[{t_idx}] after clean_cell + normalize ----")
        print(f"  Raw  headers : {parsed['raw_headers']}")
        print(f"  Norm headers : {parsed['headers']}")
        for r_idx, row in enumerate(parsed["data"][:5]):
            print(f"  Data row[{r_idx}]  : {row}")
        if len(parsed["data"]) > 5:
            print(f"  ... ({len(parsed['data']) - 5} more data rows)")

    print(f"\n{'='*70}\n")

# --------------------------------------------------
# Mandatory fields definitions & validation
# --------------------------------------------------
MANDATORY_METADATA_FIELDS = {
    "week_no": "Week No / Week Number",
    "from_date": "From Date / Start Date",
    "to_date": "To Date / End Date",
    "due_date": "Due Date / Payment Due",
}

MANDATORY_SPV_DSM_FIELDS = {
    "spv_name": "SPV Name",
    "total_dsm_charges_payable": "Total DSM Charges Payable",
    "drawl_charges_payable": "Drawl Charges Payable",
    "revenue_diff": "Revenue Difference / Loss",
    "net_payable_receivable": "Net Payable / Receivable",
}

def validate_mandatory_columns(parsed_tables: dict) -> dict:
    """
    Validates that parsed HTML tables contain all required tables and mandatory columns.
    
    Returns a dict with:
      - is_valid: bool
      - missing_tables: list of str
      - missing_metadata_cols: list of str
      - missing_spv_cols: list of str
      - error_summary: str
    """
    metadata_rows = (parsed_tables or {}).get("metadata_table", [])
    spv_rows = (parsed_tables or {}).get("spv_dsm_table", [])
    
    missing_tables = []
    missing_metadata_cols = []
    missing_spv_cols = []

    if not metadata_rows:
        missing_tables.append("Metadata Table")
    else:
        first_meta = metadata_rows[0] if isinstance(metadata_rows[0], dict) else {}
        for key, display_name in MANDATORY_METADATA_FIELDS.items():
            if not str(first_meta.get(key, "")).strip():
                missing_metadata_cols.append(f"{display_name} [{key}]")

    if not spv_rows:
        missing_tables.append("SPV DSM Table")
    else:
        data_rows = [
            r for r in spv_rows
            if isinstance(r, dict)
            and str(r.get("spv_name", "")).strip().upper() != "TOTAL"
        ]
        if not data_rows:
            missing_tables.append("SPV DSM Table (No valid SPV data rows found)")
        else:
            for key, display_name in MANDATORY_SPV_DSM_FIELDS.items():
                if key == "spv_name":
                    has_missing_value = any(
                        not str(row.get(key, "")).strip() for row in data_rows
                    )
                else:
                    has_missing_value = not any(
                        str(row.get(key, "")).strip() for row in data_rows
                    )
                if has_missing_value:
                    missing_spv_cols.append(f"{display_name} [{key}]")

    is_valid = len(missing_tables) == 0 and len(missing_metadata_cols) == 0 and len(missing_spv_cols) == 0

    errors = []
    if missing_tables:
        errors.append(f"Missing Table(s): {', '.join(missing_tables)}")
    if missing_metadata_cols:
        errors.append(f"Missing Metadata Column(s): {', '.join(missing_metadata_cols)}")
    if missing_spv_cols:
        errors.append(f"Missing SPV DSM Column(s): {', '.join(missing_spv_cols)}")

    return {
        "is_valid": is_valid,
        "missing_tables": missing_tables,
        "missing_metadata_cols": missing_metadata_cols,
        "missing_spv_cols": missing_spv_cols,
        "error_summary": " | ".join(errors) if errors else ""
    }

# --------------------------------------------------
# ✅ PUBLIC FUNCTION
# --------------------------------------------------
def parse_html_tables(body_html: str) -> Dict[str, List[Dict]]:
    tables   = extract_tables(body_html)
    metadata = []
    spv_dsm  = []
    metadata_score = 0
    spv_score = 0

    # Outlook can wrap the business tables in layout tables or prepend other
    # tables. Inspect all candidates and retain the most complete header match.
    for t in tables:
        parsed = parse_single_table(t)
        if not parsed:
            continue

        headers = parsed["headers"]
        candidate_metadata_score = header_match_score(headers, METADATA_HEADER_MAP)
        candidate_spv_score = header_match_score(headers, SPV_DSM_HEADER_MAP)

        candidate_metadata = build_table(parsed, METADATA_HEADER_MAP)
        if candidate_metadata and (
            candidate_metadata_score > metadata_score
            or (
                candidate_metadata_score == metadata_score
                and len(candidate_metadata) > len(metadata)
            )
        ):
            metadata = candidate_metadata
            metadata_score = candidate_metadata_score

        candidate_spv = build_table(parsed, SPV_DSM_HEADER_MAP)
        if candidate_spv and (
            candidate_spv_score > spv_score
            or (
                candidate_spv_score == spv_score
                and len(candidate_spv) > len(spv_dsm)
            )
        ):
            spv_dsm = candidate_spv
            spv_score = candidate_spv_score

    if spv_dsm and any(not str(row.get("spv_name") or "").strip() for row in spv_dsm):
        try:
            spv_dsm = recover_missing_spv_names(body_html, spv_dsm)
        except Exception as exc:
            print(f"[SPV-RECOVERY] LLM recovery failed; email will be retried: {exc}")

    return {
        "metadata_table": metadata,
        "spv_dsm_table":  spv_dsm,
    }
