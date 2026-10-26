"""
debit_credit_pdf_extractor.py  (nis_booking)
============================================
Extract the fields needed for SAP checklist creation from a finance-team
Debit/Credit Note PDF (the reply attachments to the Debit/Credit Notes email).

For each note PDF we pull:
    * ref_number : the document number, e.g. "DN/26-27/259-DSM" / "CN/26-27/288-DSM"
    * amount     : the note amount, e.g. 3928484.00  (always returned positive)
    * doc_date   : the note date, normalised to SAP format DD.MM.YYYY
    * note_type  : "DN" (debit) or "CN" (credit) when detectable
    * dsm_week   : the "DSM Week: <from> to <to>" line when present

Text is read with pdfplumber (same dependency used by invoice_extractor). All
fields are best-effort: a missing field comes back as "" (or 0.0 for amount)
so the caller can decide how to handle an incomplete note.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

from llm_curl_client import call_llm_curl

import config

# ── PDF → text ────────────────────────────────────────────────────────────────
def extract_pdf_text(pdf_path: str) -> str:
    if not pdf_path:
        return ""
    p = Path(pdf_path)
    if not p.exists() or p.suffix.lower() != ".pdf":
        print(f"[DC-PDF] Not a usable PDF: {pdf_path}")
        return ""
    try:
        import pdfplumber
    except ImportError:
        print("[DC-PDF] pdfplumber is not installed — cannot read PDF text.")
        return ""
    parts = []
    try:
        with pdfplumber.open(str(p)) as pdf:
            for page in pdf.pages:
                parts.append(page.extract_text() or "")
    except Exception as exc:
        print(f"[DC-PDF] Failed to read PDF text: {exc}")
        return ""
    return "\n".join(parts).strip()


# ── Field parsers ───────────────────────────────────────────────────────────
_REF_RE = re.compile(r"Doc\.?\s*No\.?\s*[:\-]?\s*((?:DN|CN)\s*/\s*[0-9A-Za-z/\-]+)", re.IGNORECASE)
_REF_FALLBACK_RE = re.compile(r"\b((?:DN|CN)\s*/\s*[0-9]{2}\s*-\s*[0-9]{2}\s*/\s*[0-9A-Za-z\-]+)", re.IGNORECASE)
_DATE_RE = re.compile(r"Date\s*[:\-]?\s*(\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4})", re.IGNORECASE)
_AMOUNT_RE = re.compile(r"(\d{1,3}(?:,\d{2,3})+\.\d{2})")
_NOTE_TYPE_RE = re.compile(r"\b(DEBIT|CREDIT)\s+NOTE\b", re.IGNORECASE)
_WEEK_RE = re.compile(r"DSM\s*Week\s*[:\-]?\s*(.+)", re.IGNORECASE)


def _clean_ref(raw: str) -> str:
    return re.sub(r"\s+", "", str(raw or "")).upper()


def _to_amount(raw: str) -> float:
    try:
        return abs(float(str(raw).replace(",", "").strip()))
    except Exception:
        return 0.0


def _normalize_date(raw: str) -> str:
    """Return the date as SAP-style DD.MM.YYYY, best-effort."""
    raw = str(raw or "").strip()
    if not raw:
        return ""
    m = re.match(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})", raw)
    if not m:
        return raw
    d, mth, y = m.group(1), m.group(2), m.group(3)
    if len(y) == 2:
        y = "20" + y
    return f"{int(d):02d}.{int(mth):02d}.{y}"


def extract_dc_pdf_data(pdf_path: str) -> dict:
    """Parse a Debit/Credit Note PDF and return its key fields.

    Returns a dict:
        {ref_number, amount, amount_text, doc_date, note_type, dsm_week,
         raw_text_len, extraction_method}

    Method selection (config.USE_LLM_EXTRACTION):
      * True  → LLM extraction (same Azure endpoint as llm_client.bat), with an
                automatic fall back to the keyword parser for any field the LLM
                does not return (or if the whole LLM call fails).
      * False → deterministic keyword/regex parser only.
    """
    text = extract_pdf_text(pdf_path)
    result = {
        "ref_number": "",
        "amount": 0.0,
        "amount_text": "",
        "doc_date": "",
        "note_type": "",
        "dsm_week": "",
        "raw_text_len": len(text or ""),
        "extraction_method": "none",
    }
    if not text:
        return result

    use_llm = getattr(config, "USE_LLM_EXTRACTION", False)
    kw = _keyword_extract(text)

    if not use_llm:
        kw["extraction_method"] = "keyword"
        result.update(kw)
        return result

    # LLM path — fall back to keyword values for anything the LLM misses.
    llm = {}
    try:
        llm = _llm_extract(text)
    except Exception as exc:
        print(f"[DC-PDF] LLM extraction failed ({exc}) — falling back to keyword parser.")

    merged = dict(kw)
    for key in ("ref_number", "amount", "amount_text", "doc_date", "note_type"):
        val = llm.get(key)
        if key == "amount":
            if _to_amount(val) > 0:
                merged["amount"] = _to_amount(val)
        elif val:
            merged[key] = val
    # keep dsm_week from keyword parse (not requested from the LLM)
    merged["extraction_method"] = "llm" if llm else "keyword_fallback"
    result.update(merged)
    return result


def _keyword_extract(text: str) -> dict:
    """Deterministic regex extraction of ref/amount/date/note_type/week."""
    out = {"ref_number": "", "amount": 0.0, "amount_text": "",
           "doc_date": "", "note_type": "", "dsm_week": ""}
    if not text:
        return out

    m = _REF_RE.search(text) or _REF_FALLBACK_RE.search(text)
    if m:
        out["ref_number"] = _clean_ref(m.group(1))

    m = _DATE_RE.search(text)
    if m:
        out["doc_date"] = _normalize_date(m.group(1))

    m = _AMOUNT_RE.search(text)
    if m:
        out["amount_text"] = m.group(1)
        out["amount"] = _to_amount(m.group(1))

    m = _NOTE_TYPE_RE.search(text)
    if m:
        out["note_type"] = "DN" if m.group(1).upper().startswith("DEBIT") else "CN"
    elif out["ref_number"].startswith(("DN", "CN")):
        out["note_type"] = out["ref_number"][:2]

    m = _WEEK_RE.search(text)
    if m:
        out["dsm_week"] = re.sub(r"\s+", " ", m.group(1)).strip()[:120]

    return out


# ── LLM extraction (Azure OpenAI — same endpoint as llm_client.bat) ───────────
_LLM_SYSTEM_PROMPT = (
    "You are a precise data-extraction assistant. You are given the raw text of "
    "a vendor Debit Note or Credit Note PDF. Extract exactly three fields: the "
    "note Amount, the Reference / Document Number (labelled 'Doc.No.', e.g. "
    "'DN/26-27/259-DSM' or 'CN/26-27/288-DSM'), and the note Date (labelled "
    "'Date'). Respond ONLY with a compact JSON object of the exact form "
    '{"amount": "<value>", "reference_number": "<value>", "date": "<value>"} '
    "and nothing else. Preserve the values exactly as printed (keep the amount's "
    "commas and decimals, keep the reference slashes, keep the date separators). "
    "If a field is not present, use an empty string for it."
)


def _call_llm(messages: list, timeout: int = 60) -> str:
    """Send chat *messages* to Azure OpenAI via curl and return the raw reply text.

    Uses ``llm_curl_client.call_llm_curl`` which mirrors llm_client.bat —
    writes the payload to a temp JSON file and calls curl via subprocess.
    This works on the VM where Python requests cannot reach the endpoint.
    """
    return call_llm_curl(messages, max_tokens=200, timeout=timeout)


def _llm_extract(text: str) -> dict:
    """Ask the LLM for {ref_number, amount, amount_text, doc_date, note_type}.

    Returns {} when the model reply cannot be parsed.
    """
    reply = _call_llm([
        {"role": "system", "content": _LLM_SYSTEM_PROMPT},
        {"role": "user", "content": "Extract the fields from this note text:\n\n" + text},
    ])
    match = re.search(r"\{.*\}", reply or "", re.DOTALL)
    if not match:
        return {}
    try:
        obj = json.loads(match.group(0))
    except Exception:
        return {}

    amount_text = str(obj.get("amount") or "").strip()
    ref = _clean_ref(obj.get("reference_number") or "")
    date_raw = str(obj.get("date") or "").strip()
    note_type = ""
    if ref.startswith("DN"):
        note_type = "DN"
    elif ref.startswith("CN"):
        note_type = "CN"

    return {
        "ref_number": ref,
        "amount": _to_amount(amount_text),
        "amount_text": amount_text,
        "doc_date": _normalize_date(date_raw),
        "note_type": note_type,
    }


if __name__ == "__main__":
    import sys, glob, json
    targets = sys.argv[1:] or glob.glob("../debit-credit-email-pdf-reply/*.pdf")
    for t in targets:
        print("=" * 60)
        print(t)
        print(json.dumps(extract_dc_pdf_data(t), indent=2))
