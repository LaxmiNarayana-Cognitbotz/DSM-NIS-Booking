"""
debit_credit_email.py
======================
Builds the "Debit/Credit Notes" email that is sent to the Finance team when a
DSM email's SPV table has exactly ``config.DEBIT_CREDIT_TRIGGER_ROW_COUNT``
(default 5) valid SPV rows.

The email mirrors the template in
``send-debit-credit-email-finance/Debit_Credit Notes.eml``: a transposed table
where each SPV becomes a column, values are rendered in the Indian number
format (e.g. ``21,63,738.00``), and a header block carries the total DSM amount,
company code and CN/DN due date.

This module is pure (no network / DB). Sending + tracking is orchestrated by
``main.py``; recipients come from ``finance_recipients_sp.get_finance_recipients``.
"""

from __future__ import annotations

import html
import re
from datetime import datetime

import config

_TOTAL_KEY = "TOTAL"

# SPV-table field keys (produced by classified_html_table_parser).
_F_DSM = "total_dsm_charges_payable"
_F_REVDIFF = "revenue_diff"
_F_NET = "net_payable_receivable"


# ─────────────────────────────────────────────────────────────────────────────
# Value helpers
# ─────────────────────────────────────────────────────────────────────────────
def _to_float(value) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "").strip())
    except Exception:
        return 0.0


def format_inr(value) -> str:
    """Format a number using the Indian grouping system with 2 decimals.

    Examples: 3716 -> '3,716.00', -3932200 -> '-39,32,200.00',
              2163738 -> '21,63,738.00'.
    Non-numeric input is returned as a trimmed string.
    """
    raw = str(value if value is not None else "").strip()
    if raw == "":
        return ""
    try:
        num = float(raw.replace(",", ""))
    except Exception:
        return raw

    neg = num < 0
    num = abs(round(num, 2))
    int_part = int(num)
    cents = int(round((num - int_part) * 100))
    if cents == 100:  # rounding edge
        int_part += 1
        cents = 0

    s = str(int_part)
    if len(s) > 3:
        last3 = s[-3:]
        rest = s[:-3]
        parts = []
        while len(rest) > 2:
            parts.insert(0, rest[-2:])
            rest = rest[:-2]
        if rest:
            parts.insert(0, rest)
        grouped = ",".join(parts) + "," + last3
    else:
        grouped = s

    result = f"{grouped}.{cents:02d}"
    return ("-" + result) if neg else result


def _reformat_date(raw: str) -> str:
    """Reformat a date like '13-May-2026' or '20-Apr-2026' to '13/05/2026'.

    Returns the original string unchanged if it cannot be parsed.
    """
    raw = str(raw or "").strip()
    if not raw:
        return ""
    for fmt in ("%d-%b-%Y", "%d-%B-%Y", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%d/%m/%Y")
        except Exception:
            continue
    return raw


# ─────────────────────────────────────────────────────────────────────────────
# Row selection
# ─────────────────────────────────────────────────────────────────────────────
def _is_total_row(row: dict) -> bool:
    name = re.sub(r"\s+", " ", str((row or {}).get("spv_name") or "").strip()).upper()
    return name == _TOTAL_KEY


def valid_spv_rows(spv_rows: list) -> list:
    """Return SPV rows that have a name and are not the 'Total' summary row."""
    out = []
    for row in spv_rows or []:
        if not isinstance(row, dict) or _is_total_row(row):
            continue
        if str(row.get("spv_name") or "").strip():
            out.append(row)
    return out


def total_row(spv_rows: list) -> dict:
    """Return the email's 'Total' row if present, else an empty dict."""
    for row in spv_rows or []:
        if isinstance(row, dict) and _is_total_row(row):
            return row
    return {}


def should_send(spv_rows: list) -> bool:
    """True when the SPV table has exactly the trigger number of valid rows."""
    target = int(getattr(config, "DEBIT_CREDIT_TRIGGER_ROW_COUNT", 5))
    return len(valid_spv_rows(spv_rows)) == target


# ─────────────────────────────────────────────────────────────────────────────
# Totals
# ─────────────────────────────────────────────────────────────────────────────
def _column_totals(rows: list, total: dict) -> dict:
    """Totals for each metric. Prefer the email's Total row; fall back to a
    computed sum across the valid SPV rows when no Total row is present.
    """
    out = {}
    for key in (_F_DSM, _F_REVDIFF, _F_NET):
        if total and str(total.get(key, "")).strip() != "":
            out[key] = _to_float(total.get(key))
        else:
            out[key] = sum(_to_float(r.get(key)) for r in rows)
    return out


# ─────────────────────────────────────────────────────────────────────────────
# Email building
# ─────────────────────────────────────────────────────────────────────────────
def build_subject(metadata_row: dict, spv_rows: list) -> str:
    """Subject including a run marker so replies can be matched back.

    e.g. 'Debit/Credit Notes | Week 04 | 20-Apr-2026 to 26-Apr-2026'
    """
    base = str(getattr(config, "DEBIT_CREDIT_SUBJECT", "Debit/Credit Notes"))
    week_no = str((metadata_row or {}).get("week_no") or "").strip()
    from_date = str((metadata_row or {}).get("from_date") or "").strip()
    to_date = str((metadata_row or {}).get("to_date") or "").strip()
    parts = [base]
    if week_no:
        parts.append(f"Week {week_no}")
    if from_date or to_date:
        parts.append(f"{from_date} to {to_date}".strip())
    return " | ".join(parts)


def build_html(metadata_row: dict, spv_rows: list) -> str:
    """Build the HTML body for the debit/credit email from the SPV table."""
    rows = valid_spv_rows(spv_rows)
    total = total_row(spv_rows)
    totals = _column_totals(rows, total)

    week_no = str((metadata_row or {}).get("week_no") or "").strip()
    from_date = str((metadata_row or {}).get("from_date") or "").strip()
    to_date = str((metadata_row or {}).get("to_date") or "").strip()
    due_date = _reformat_date((metadata_row or {}).get("due_date"))
    company_code = str(getattr(config, "DEBIT_CREDIT_COMPANY_CODE", "") or "").strip()
    dsm_amount = format_inr(totals.get(_F_NET))

    esc = lambda v: html.escape(str(v if v is not None else ""))

    # ── Header lines ──────────────────────────────────────────────────────
    header_lines = [
        "<p>Dear Team,</p>",
        "<p>Please provide Debit/Credit Notes</p>",
        f"<p><b>DSM Amount</b> - {esc(dsm_amount)}</p>",
    ]
    if company_code:
        header_lines.append(f"<p><b>C. Code</b> : {esc(company_code)}</p>")
    if due_date:
        header_lines.append(f"<p><b>Due Date CN/DN</b> : {esc(due_date)}</p>")

    # ── Transposed table: one column per SPV + a Total column ─────────────
    spv_headers = "".join(f"<th>{esc(r.get('spv_name'))}</th>" for r in rows)
    week_label = f"Week {week_no}".strip()

    dsm_cells = "".join(f"<td>{esc(format_inr(r.get(_F_DSM)))}</td>" for r in rows)
    rev_cells = "".join(f"<td>{esc(format_inr(r.get(_F_REVDIFF)))}</td>" for r in rows)
    net_cells = "".join(f"<td>{esc(format_inr(r.get(_F_NET)))}</td>" for r in rows)

    table = f"""
<table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse;">
  <tr>
    <th>{esc(week_label)}</th>
    {spv_headers}
    <th>Total as per Statement (Payable (+)/ Receivable (-))</th>
  </tr>
  <tr>
    <td>{esc(from_date)}<br>{esc(to_date)}</td>
    {''.join('<td></td>' for _ in rows)}
    <td></td>
  </tr>
  <tr>
    <td>DSM Charges (Rs.)</td>
    {dsm_cells}
    <td>{esc(format_inr(totals.get(_F_DSM)))}</td>
  </tr>
  <tr>
    <td>Revenue Difference Payable (+)/ Receivable (-)</td>
    {rev_cells}
    <td>{esc(format_inr(totals.get(_F_REVDIFF)))}</td>
  </tr>
  <tr>
    <td>Total</td>
    {net_cells}
    <td>{esc(format_inr(totals.get(_F_NET)))}</td>
  </tr>
</table>
"""

    return "".join(header_lines) + table + "<p>Thanks &amp; Regards</p>"


def build_email(metadata_row: dict, spv_rows: list) -> dict:
    """Convenience: return subject + html + key derived values for tracking."""
    rows = valid_spv_rows(spv_rows)
    totals = _column_totals(rows, total_row(spv_rows))
    return {
        "subject": build_subject(metadata_row, spv_rows),
        "html": build_html(metadata_row, spv_rows),
        "dsm_amount": format_inr(totals.get(_F_NET)),
        "company_code": str(getattr(config, "DEBIT_CREDIT_COMPANY_CODE", "") or "").strip(),
        "due_date_cn_dn": _reformat_date((metadata_row or {}).get("due_date")),
        "week_no": str((metadata_row or {}).get("week_no") or "").strip(),
        "from_date": str((metadata_row or {}).get("from_date") or "").strip(),
        "to_date": str((metadata_row or {}).get("to_date") or "").strip(),
    }
