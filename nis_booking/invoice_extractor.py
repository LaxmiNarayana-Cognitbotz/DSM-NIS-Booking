"""
invoice_extractor.py  (nis_booking)
===================================
Extracts the invoice / bill number from a vendor PDF attached to the DSM-NIS
email.

Pipeline:
    1. pdfplumber  → pull raw text out of the PDF.
    2. LLM (GPT)   → read that text and return ONLY the invoice number.

The LLM backend is selectable via ``config.LLM_USE_OPENAI_DIRECT``:
    * False (default) → internal Azure OpenAI endpoint (llm_client credentials).
    * True            → public OpenAI API using OPENAI_API_KEY from the .env file.

This is used for vendors whose PDFs *do* carry an invoice number. For the
special vendor codes (see ``config.NIS_AUTO_INVOICE_VENDOR_CODES``) the PDF has
no invoice data, so the caller keeps auto-generating the number instead.

Example
-------
The bill ``WR_26_W08_0005_Adani Wind Energy Kutchh One Limited_Ratadiya.pdf``
contains the line ``Bill NO . : WR/26/W08/0005`` → invoice number is
``WR/26/W08/0005``.
"""

import json
import re
from pathlib import Path

from llm_curl_client import call_llm_curl

import config


# ── PDF → text ────────────────────────────────────────────────────────────────
def extract_pdf_text(pdf_path: str) -> str:
    """Return all text from *pdf_path* using pdfplumber (empty string on failure)."""
    if not pdf_path:
        return ""
    p = Path(pdf_path)
    if not p.exists() or p.suffix.lower() != ".pdf":
        print(f"[InvoiceExtract] Not a usable PDF: {pdf_path}")
        return ""
    try:
        import pdfplumber
    except ImportError:
        print("[InvoiceExtract] pdfplumber is not installed — cannot read PDF text.")
        return ""

    text_parts = []
    try:
        with pdfplumber.open(str(p)) as pdf:
            for page in pdf.pages:
                text_parts.append(page.extract_text() or "")
    except Exception as exc:
        print(f"[InvoiceExtract] Failed to read PDF text: {exc}")
        return ""
    return "\n".join(text_parts).strip()


# ── LLM call (Azure OpenAI) ─────────────────────────────────────────────────────
_SYSTEM_PROMPT = (
    "You are a precise data-extraction assistant. You are given the raw text of "
    "a vendor reimbursement / DSM charges invoice. Find the invoice number "
    "(it may be labelled 'Bill No', 'Bill NO', 'Invoice No', 'Invoice Number' "
    "or similar). Respond with a compact JSON object of the exact form "
    '{\"invoice_number\": \"<value>\"} and nothing else. Preserve the value '
    "exactly as printed, including slashes and leading zeros. If no invoice "
    'number is present, respond {\"invoice_number\": \"\"}.'
)


def _call_llm(pdf_text: str, timeout: int = 60) -> str:
    """Send *pdf_text* to the configured LLM via curl and return the raw reply.

    Uses ``llm_curl_client.call_llm_curl`` which mirrors llm_client.bat —
    writes the payload to a temp JSON file and calls curl via subprocess.
    This works on the VM where Python requests cannot reach the endpoint.
    """
    messages = [
        {"role": "system", "content": _SYSTEM_PROMPT},
        {"role": "user", "content": "Extract the invoice number from this invoice text:\n\n" + pdf_text},
    ]
    return call_llm_curl(messages, max_tokens=100, timeout=timeout)


def _parse_invoice_number(reply: str) -> str:
    """Pull the invoice number out of the model reply (JSON or plain text)."""
    if not reply:
        return ""
    # Preferred: a JSON object {"invoice_number": "..."}
    match = re.search(r"\{.*\}", reply, re.DOTALL)
    if match:
        try:
            obj = json.loads(match.group(0))
            # If the model returned the expected key, trust it verbatim — even
            # an empty string means "no invoice number found", so don't fall
            # through to the plain-text heuristics below.
            if "invoice_number" in obj:
                return str(obj.get("invoice_number") or "").strip()
        except Exception:
            pass
    # Fallback: strip common labels / quotes from a plain-text reply.
    cleaned = reply.strip().strip('"').strip()
    cleaned = re.sub(r"(?i)^(invoice|bill)\s*(number|no)\s*[:\-]?\s*", "", cleaned)
    return cleaned.strip().strip('"').strip()


def extract_invoice_number_from_pdf(pdf_path: str) -> str:
    """
    Full pipeline: PDF → text (pdfplumber) → invoice number (LLM).

    Returns the extracted invoice number, or "" when it cannot be determined
    (missing PDF, empty text, or LLM failure). The caller decides how to handle
    an empty result.
    """
    pdf_text = extract_pdf_text(pdf_path)
    if not pdf_text:
        print(f"[InvoiceExtract] No text extracted from PDF: {pdf_path}")
        return ""
    try:
        reply = _call_llm(pdf_text)
    except Exception as exc:
        print(f"[InvoiceExtract] LLM call failed: {exc}")
        return ""
    invoice_number = _parse_invoice_number(reply)
    if invoice_number:
        print(f"[InvoiceExtract] Extracted invoice number: {invoice_number}")
    else:
        print(f"[InvoiceExtract] LLM returned no invoice number (reply={reply!r}).")
    return invoice_number


def extract_pdf_pages_text(pdf_path: str) -> list:
    """Return a list of per-page text strings (empty list on failure)."""
    if not pdf_path:
        return []
    try:
        import pdfplumber
    except ImportError:
        print("[InvoiceExtract] pdfplumber not installed — cannot read PDF pages.")
        return []
    pages = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                pages.append(page.extract_text() or "")
    except Exception as exc:
        print(f"[InvoiceExtract] Failed to read PDF pages: {exc}")
        return []
    return pages


def _norm(s: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(s or "").upper())


def _primary_spv_token(spv_name: str) -> str:
    """Leading plant-code token of an SPV name, e.g. 'AWEK06L_MSEDCL_75 MW' → 'AWEK06L'."""
    toks = [t for t in re.split(r"[^A-Za-z0-9]+", str(spv_name or "")) if t]
    return toks[0] if toks else ""


def _extract_invoice_from_text(text: str) -> str:
    """Run the invoice-number LLM prompt on a single block of text."""
    if not str(text or "").strip():
        return ""
    try:
        reply = _call_llm(text)
    except Exception as exc:
        print(f"[InvoiceExtract] LLM call failed on page text: {exc}")
        return ""
    return _parse_invoice_number(reply)


def _page_for_spv(pages: list, spv_name: str) -> str:
    """Return the page text whose content matches *spv_name* (by plant-code
    token, then by a longer normalized prefix). '' when no page matches."""
    token = _norm(_primary_spv_token(spv_name))
    if token:
        for pt in pages:
            if token in _norm(pt):
                return pt
    full = _norm(spv_name)
    if len(full) >= 5:
        for pt in pages:
            if full[:6] in _norm(pt):
                return pt
    return ""


def _assign_pages_to_spvs(pages: list, spv_names: list) -> dict:
    """Assign each SPV to a UNIQUE page using best-match scoring.

    Strategy (in priority order):
      1. Score each (spv, page) pair by how many unique tokens of the SPV name
         appear in the page text. Assign greedily: best score first, no page
         reused.
      2. For any SPV left unassigned (no token matched any remaining page),
         fall back to positional assignment: row N → page N (when page count
         matches SPV count).

    Returns {spv_name: page_text} (may have "" for unresolved SPVs).
    """
    n_spvs = len(spv_names)
    n_pages = len(pages)

    # Normalize page texts once
    norm_pages = [_norm(pt) for pt in pages]

    # Build a score matrix: how well does each SPV match each page?
    # Score = number of distinct alphanumeric tokens from the SPV name found
    # in the page text, weighted by token length (longer tokens = stronger signal).
    scores = []  # list of (score, spv_idx, page_idx)
    for si, spv in enumerate(spv_names):
        spv_tokens = [t for t in re.split(r"[^A-Za-z0-9]+", str(spv)) if t]
        for pi, norm_page in enumerate(norm_pages):
            score = 0
            for tok in spv_tokens:
                ntok = tok.upper()
                if ntok and ntok in norm_page:
                    score += len(ntok)  # weight by token length
            if score > 0:
                scores.append((score, si, pi))

    # Greedy assignment: highest score first, each SPV and page used at most once
    scores.sort(key=lambda x: -x[0])
    assigned_spvs = set()
    assigned_pages = set()
    result = {s: "" for s in spv_names}

    for score, si, pi in scores:
        if si in assigned_spvs or pi in assigned_pages:
            continue
        result[spv_names[si]] = pages[pi]
        assigned_spvs.add(si)
        assigned_pages.add(pi)
        print(f"[InvoiceExtract]   Page {pi+1} matched to '{spv_names[si]}' (score={score})")

    # Fallback: positional assignment for unmatched SPVs when page count == SPV count
    if len(assigned_spvs) < n_spvs and n_pages == n_spvs:
        for si in range(n_spvs):
            if si not in assigned_spvs and si not in assigned_pages:
                result[spv_names[si]] = pages[si]
                assigned_spvs.add(si)
                assigned_pages.add(si)
                print(f"[InvoiceExtract]   Page {si+1} assigned to '{spv_names[si]}' "
                      f"(positional fallback)")

    # Log unmatched SPVs
    for si in range(n_spvs):
        if si not in assigned_spvs:
            print(f"[InvoiceExtract]   No page matched for '{spv_names[si]}'")

    return result


def extract_invoice_numbers_for_rows(pdf_path: str, spv_rows: list) -> dict:
    """Resolve the invoice number for each SPV row from a single attached PDF.

    Rule (applies to 3-row DSM tables — the caller decides when to use it):

      * If the PDF text contains ``config.INOX_KEYWORD`` ("INOX") → the bill has
        one invoice per SPV page. For each row, find the page whose text matches
        the row's SPV name and extract THAT page's invoice number. Each row thus
        gets its own invoice number.
      * Otherwise → the bill has a single invoice; extract it once and share the
        same number across all rows.

    Returns ``{spv_name: invoice_number}`` (invoice_number may be "" if a page
    or number could not be resolved — the booking layer then falls back to its
    normal generated/extracted value).
    """
    spv_names = [str((r or {}).get("spv_name") or "").strip()
                 for r in (spv_rows or [])]
    spv_names = [s for s in spv_names if s]
    result = {s: "" for s in spv_names}

    pages = extract_pdf_pages_text(pdf_path)
    if not pages:
        return result
    full_text = "\n".join(pages)

    inox_kw = str(getattr(config, "INOX_KEYWORD", "INOX") or "").upper()
    is_inox = bool(inox_kw) and inox_kw in full_text.upper()

    if is_inox:
        print(f"[InvoiceExtract] INOX bill detected ({len(pages)} page(s)) — "
              f"extracting a per-SPV invoice number.")
        # Use the improved page assignment that ensures each SPV gets a unique page
        page_map = _assign_pages_to_spvs(pages, spv_names)
        for spv in spv_names:
            page_text = page_map.get(spv, "")
            if not page_text:
                print(f"[InvoiceExtract] No page matched SPV '{spv}'.")
                continue
            result[spv] = _extract_invoice_from_text(page_text)
            print(f"[InvoiceExtract]   {spv} → {result[spv] or '(none)'}")
    else:
        shared = _extract_invoice_from_text(full_text)
        print(f"[InvoiceExtract] Non-INOX bill — single invoice '{shared}' "
              f"shared by {len(spv_names)} row(s).")
        for spv in spv_names:
            result[spv] = shared

    return result


if __name__ == "__main__":
    import sys

    target = sys.argv[1] if len(sys.argv) > 1 else (
        "../invoice example/WR_26_W08_0005_Adani Wind Energy Kutchh One "
        "Limited_Ratadiya.pdf"
    )
    print("Invoice number:", extract_invoice_number_from_pdf(target))