import re

import config


def _normalize_spv_name(value: object) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())


def is_email_pdf_spv(spv_name: object) -> bool:
    """Return whether an SPV must use email.pdf and an auto-generated invoice."""
    normalized_name = _normalize_spv_name(spv_name)
    if not normalized_name:
        return False

    keywords = getattr(config, "NIS_EMAIL_PDF_SPV_KEYWORDS", set())
    return any(
        normalized_keyword and normalized_keyword in normalized_name
        for normalized_keyword in (_normalize_spv_name(keyword) for keyword in keywords)
    )
