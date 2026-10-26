"""
recipients_sp.py
================
SharePoint-backed store for the NIS completion / failure email recipients.

This is the ONLY tracking-style data the nis_booking service keeps on
SharePoint. Everything else (email_runs, run_steps, booking_rows) lives in the
local SQLite database (see ``db_manager.py``).

The recipients are stored in a SharePoint workbook (default
``nis_completion_recipients.xlsx``) on a single sheet (default
``nis_completion_recipients``) with one column header ``emails``. Each row may
hold one address or a comma/semicolon separated list.

This module is READ-ONLY — it never creates or modifies the workbook.
If the file does not exist on SharePoint, it falls back to the
``NIS_COMPLETION_RECIPIENTS`` config variable. The workbook must be
created and maintained manually on SharePoint.
"""

from __future__ import annotations

import urllib.parse

import requests

import config
import sharepoint_upload as sp

_SHEET = str(getattr(config, "NIS_RECIPIENTS_SHEET", "nis_completion_recipients"))
_FOLDER = str(getattr(config, "NIS_RECIPIENTS_SP_FOLDER", getattr(config, "FOLDER_PATH", "")) or "").strip("/")
_FILE_NAME = str(getattr(config, "NIS_RECIPIENTS_FILE_NAME", "nis_completion_recipients.xlsx"))
_HEADERS = ["emails"]

_ITEM_ID_CACHE = ""


def _sp_file_path() -> str:
    return f"{_FOLDER}/{_FILE_NAME}".strip("/") if _FOLDER else _FILE_NAME


def _encode_sp_path(path: str) -> str:
    return "/".join(urllib.parse.quote(seg, safe="") for seg in path.split("/") if seg)


def _graph_headers(content_type: bool = False) -> dict:
    h = dict(sp._headers())
    if content_type:
        h["Content-Type"] = "application/json"
    return h


def _graph_drive_url() -> str:
    return f"https://graph.microsoft.com/v1.0/drives/{sp._get_drive_id()}"


def _item_id() -> str:
    global _ITEM_ID_CACHE
    if _ITEM_ID_CACHE:
        return _ITEM_ID_CACHE

    if _FOLDER:
        sp.ensure_sp_folder(_FOLDER)

    url = f"{_graph_drive_url()}/root:/{_encode_sp_path(_sp_file_path())}"
    resp = requests.get(url, headers=_graph_headers(), timeout=60)
    if resp.status_code == 200:
        _ITEM_ID_CACHE = resp.json()["id"]
        return _ITEM_ID_CACHE
    if resp.status_code == 404:
        # File does not exist on SharePoint. Do NOT create it — this file is
        # managed manually and should never be overwritten by the automation.
        # The caller will fall back to config.NIS_COMPLETION_RECIPIENTS.
        raise FileNotFoundError(
            f"Recipients workbook not found at {_sp_file_path()}. "
            f"Please create it manually on SharePoint with a sheet "
            f"'{_SHEET}' and an 'emails' column."
        )
    raise RuntimeError(
        f"Could not resolve recipients workbook: {resp.status_code} {resp.text[:200]}"
    )


def _get_values() -> list:
    item_id = _item_id()
    url = (f"{_graph_drive_url()}/items/{item_id}/workbook/worksheets('{_SHEET}')"
           f"/usedRange(valuesOnly=true)?$select=values")
    resp = requests.get(url, headers=_graph_headers(), timeout=60)
    if resp.status_code == 200:
        return resp.json().get("values", []) or []
    if resp.status_code == 404:
        return []
    raise RuntimeError(
        f"Could not read recipients sheet '{_SHEET}': {resp.status_code} {resp.text[:200]}"
    )


def _split_addresses(raw: str, seen: set, out: list):
    for addr in str(raw or "").replace(";", ",").split(","):
        email = addr.strip()
        key = email.lower()
        if email and key not in seen:
            seen.add(key)
            out.append(email)


def get_recipients() -> list:
    """Return the de-duplicated recipient address list from SharePoint.

    Falls back to ``config.NIS_COMPLETION_RECIPIENTS`` if the SharePoint sheet
    is empty or cannot be read, so notifications never silently break.
    """
    seen: set = set()
    recipients: list = []
    try:
        values = _get_values()
        header = [str(v or "").strip().lower() for v in (values[0] if values else [])]
        email_idx = header.index("emails") if "emails" in header else 0
        for raw_row in values[1:]:
            cell = raw_row[email_idx] if email_idx < len(raw_row) else ""
            _split_addresses(cell, seen, recipients)
    except Exception as exc:
        print(f"[RECIPIENTS-SP] Could not read recipients from SharePoint: {exc}")

    if not recipients:
        fallback = str(getattr(config, "NIS_COMPLETION_RECIPIENTS", "") or "").strip()
        if fallback:
            _split_addresses(fallback, seen, recipients)
    return recipients
