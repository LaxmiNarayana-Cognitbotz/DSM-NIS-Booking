"""
read_email.py
=============
Fetches emails from Outlook via Microsoft Graph API.

Public API
----------
  fetch_emails_last_n_days(days=2)  → list of raw email dicts
  process_email(message)            → enriched dict with parsed tables + downloads
  process_latest_email()            → backwards-compat: process only the newest email

Each returned dict includes:
  message_id, received_at, title, body_html,
  downloaded_files, parsed_tables
"""

import os
import json
import re
import base64
import msal
import requests
from datetime import datetime, timezone, timedelta
from urllib.parse import unquote, urlparse

from classified_html_table_parser import parse_html_tables, validate_mandatory_columns
import config

# ─────────────────────────────────────────────────────────────────────────────
# Configuration  (all values loaded from config.py)
# ─────────────────────────────────────────────────────────────────────────────
TENANT_ID     = config.GRAPH_TENANT_ID
CLIENT_ID     = config.GRAPH_CLIENT_ID
CLIENT_SECRET = config.GRAPH_CLIENT_SECRET
USER_EMAIL    = config.USER_EMAIL

AUTHORITY = config.GRAPH_AUTHORITY
SCOPE     = config.GRAPH_SCOPE

# DSM PDFs are initially downloaded into temp-book/<msg_id>/ next to main.py.
# main.py moves them into temp-book/<spv_name>/ once the SPV name is known,
# uploads them to SharePoint, then deletes the local temp tree.
_SCRIPT_DIR    = os.path.dirname(os.path.abspath(__file__))
TEMP_BOOK_ROOT = os.path.join(_SCRIPT_DIR, "temp-book")

# ─────────────────────────────────────────────────────────────────────────────
# Keyword that must appear somewhere in the subject to qualify as a DSM/NIS
# automation email.  Matching is case-insensitive (so "dsm-nis-booking",
# "DSM-NIS-BOOKING", "Re: DSM-NIS-Booking weekly" all match).
# ─────────────────────────────────────────────────────────────────────────────
DSM_SUBJECT_KEYWORD = config.DSM_SUBJECT_KEYWORD

DELETED_MAIL_FOLDER_NAMES = (
    "deleteditems",
    "recoverableitemsdeletions",
    "recoverableitemspurges",
)


# ─────────────────────────────────────────────────────────────────────────────
# Auth
# ─────────────────────────────────────────────────────────────────────────────
def _get_access_token() -> str:
    app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET,
    )
    token = app.acquire_token_for_client(scopes=SCOPE)
    if "access_token" not in token:
        raise RuntimeError(f"MSAL token error: {token}")
    return token["access_token"]


# ─────────────────────────────────────────────────────────────────────────────
# PDF helpers
# ─────────────────────────────────────────────────────────────────────────────
def _extract_pdf_links(html: str) -> list:
    pdfs    = []
    pattern = re.compile(
        r'<a[^>]+href=["\'"]([^"\']+\.pdf[^"\']*)["\'][^>]*>(.*?)</a>',
        re.IGNORECASE | re.DOTALL,
    )
    for url, text in pattern.findall(html):
        filename = text.strip()
        if not filename.lower().endswith(".pdf"):
            filename = os.path.basename(urlparse(url).path) or "attachment.pdf"
        pdfs.append({"url": unquote(url), "filename": filename})
    return pdfs


def _download_url(url: str, filename: str, subdir: str) -> str:
    os.makedirs(subdir, exist_ok=True)
    path = os.path.join(subdir, filename)
    resp = requests.get(url, stream=True, timeout=60)
    resp.raise_for_status()
    with open(path, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            if chunk:
                f.write(chunk)
    return path


# ─────────────────────────────────────────────────────────────────────────────
# Subject filter
# ─────────────────────────────────────────────────────────────────────────────
def is_dsm_email(subject: str) -> bool:
    """
    Return True when the subject contains DSM_SUBJECT_KEYWORD anywhere,
    case-insensitively, unless it is an alert/notification email.

    Match   : "AUTO-DSM-NIS-Booking"          ✅
    Match   : "dsm-nis-booking week 34"        ✅
    Match   : "Re: DSM-NIS-BOOKING summary"    ✅
    Match   : "FWD: Dsm-Nis-Booking"           ✅
    No-match: "DSM report"                     ❌
    No-match: "NIS weekly update"              ❌
    No-match: "[ALERT] Column Extraction..."   ❌
    """
    s = (subject or "").strip().lower()
    if s.startswith("[alert]") or "[alert]" in s or s.startswith("alert:") or "alert:" in s:
        return False
    return DSM_SUBJECT_KEYWORD.lower() in s



def has_required_tables(parsed_tables: dict) -> bool:
    """
    Return True only when BOTH the metadata table AND spv_dsm table are
    present, non-empty, and contain all mandatory columns.
    """
    return validate_mandatory_columns(parsed_tables).get("is_valid", False)



def _resolve_deleted_folder_ids(base_url: str, headers: dict) -> set:
    deleted_folder_ids = set()
    for folder_name in DELETED_MAIL_FOLDER_NAMES:
        try:
            resp = requests.get(
                f"{base_url}/mailFolders/{folder_name}",
                headers=headers,
                timeout=10,
            )
            if resp.status_code == 200:
                folder_id = str(resp.json().get("id") or "").strip()
                if folder_id:
                    deleted_folder_ids.add(folder_id)
        except Exception as exc:
            print(f"[EMAIL] Warning: Could not fetch {folder_name} folder id: {exc}")
    if deleted_folder_ids:
        print(f"[EMAIL] Deleted folder id(s) resolved: {len(deleted_folder_ids)}")
    return deleted_folder_ids


def _is_deleted_message(message: dict, deleted_folder_ids: set) -> bool:
    parent_id = str((message or {}).get("parentFolderId") or "").strip()
    return bool(parent_id and parent_id in deleted_folder_ids)


# ─────────────────────────────────────────────────────────────────────────────
# Core: process a single raw Graph API message dict
# ─────────────────────────────────────────────────────────────────────────────
def process_email(message: dict, headers: dict, base_url: str) -> dict:
    """
    Given a raw Graph API message object (already fetched), download
    attachments, parse tables, and return an enriched dict.
    """
    message_id  = message["id"]
    received_at = message.get("receivedDateTime", "")
    subject     = message.get("subject", "")
    body_html   = message.get("body", {}).get("content", "")

    # Per-message staging directory inside temp-book/.
    # Named after the (sanitised) message-id so parallel runs don't collide.
    # main.py will move the files into temp-book/<spv_name>/ once the SPV
    # name is known, upload them to SharePoint, then delete the tree.
    safe_id    = re.sub(r"[^a-zA-Z0-9_-]", "_", message_id)[:40]
    att_subdir = os.path.join(TEMP_BOOK_ROOT, safe_id)
    os.makedirs(att_subdir, exist_ok=True)

    downloaded_files = []

    # ── 1. Outlook PDF attachments (normal + inline) ──────────────────────
    try:
        att_resp = requests.get(
            f"{base_url}/messages/{message_id}/attachments",
            headers=headers,
            timeout=60,
        )
        att_resp.raise_for_status()
        for att in att_resp.json().get("value", []):
            name = att.get("name", "")
            if not name.lower().endswith(".pdf"):
                continue
            att_id    = att["id"]
            file_resp = requests.get(
                f"{base_url}/messages/{message_id}/attachments/{att_id}/$value",
                headers=headers,
                timeout=60,
            )
            if file_resp.status_code == 200:
                orig_path = os.path.join(att_subdir, name)
                with open(orig_path, "wb") as f:
                    f.write(file_resp.content)
                # Rename DSM attachments to start with DSM_
                base = os.path.basename(orig_path)
                if not base.startswith("DSM_"):
                    new_name = f"DSM_{base}"
                    new_path = os.path.join(att_subdir, new_name)
                    try:
                        os.replace(orig_path, new_path)
                        downloaded_files.append(new_path)
                    except Exception:
                        downloaded_files.append(orig_path)
                else:
                    downloaded_files.append(orig_path)
    except Exception as exc:
        print(f"  [WARN] Attachment download failed for {message_id[:20]}… : {exc}")

    # ── 2. PDF links embedded in HTML body ────────────────────────────────
    for pdf in _extract_pdf_links(body_html):
        try:
            path = _download_url(pdf["url"], pdf["filename"], att_subdir)
            # Rename DSM attachments to start with DSM_
            base = os.path.basename(path)
            if not base.startswith("DSM_"):
                new_name = f"DSM_{base}"
                new_path = os.path.join(att_subdir, new_name)
                try:
                    os.replace(path, new_path)
                    downloaded_files.append(new_path)
                except Exception:
                    downloaded_files.append(path)
            else:
                downloaded_files.append(path)
        except Exception as exc:
            print(f"  [WARN] PDF link download failed: {pdf['url']} → {exc}")

    # ── 3. Parse HTML tables ──────────────────────────────────────────────
    parsed_tables = parse_html_tables(body_html)

    return {
        "message_id":       message_id,
        "received_at":      received_at,
        "title":            subject,
        "body_html":        body_html,
        "downloaded_files": downloaded_files,
        "parsed_tables":    parsed_tables,
        "attachment_dir":   att_subdir,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Public: fetch ALL emails from the last N days
# ─────────────────────────────────────────────────────────────────────────────
def fetch_emails_last_n_days(days: int = 2) -> list:
    """
    Return a list of raw Graph API message dicts received in the last `days` days.
    Only fetches subject + body (no attachments yet).
    Pagination is followed so ALL matching emails are returned.

    IMPORTANT: Only the Inbox folder is queried (/mailFolders/inbox/messages).
    This prevents the automation's own sent alert/notification emails from
    being picked up and mistakenly re-processed.  Deleted emails inside the
    Inbox are also excluded via the parentFolderId guard.
    """
    token    = _get_access_token()
    headers  = {"Authorization": f"Bearer {token}"}
    base_url = f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}"

    deleted_folder_ids = _resolve_deleted_folder_ids(base_url, headers)

    since_dt = (datetime.now(timezone.utc) - timedelta(days=days)).strftime(
        "%Y-%m-%dT%H:%M:%SZ"
    )

    # ── Query ONLY the Inbox folder, not all mailbox folders ─────────────────
    # Using /mailFolders/inbox/messages ensures Sent Items, Drafts, Junk, and
    # any other folder are never included.  This is the safest way to avoid
    # the bot reading its own outgoing alert/notification emails.
    url = (
        f"{base_url}/mailFolders/inbox/messages"
        f"?$filter=receivedDateTime ge {since_dt}"
        "&$orderby=receivedDateTime desc"
        "&$select=id,subject,body,from,receivedDateTime,parentFolderId"
        "&$top=50"          # fetch up to 50 per page; follow @odata.nextLink
    )

    messages  = []
    seen_urls = set()
    while url:
        if url in seen_urls:
            raise RuntimeError("Graph pagination returned a repeated nextLink; stopping email fetch.")
        seen_urls.add(url)
        resp = requests.get(url, headers=headers, timeout=60)
        resp.raise_for_status()
        data = resp.json()

        for msg in data.get("value", []):
            # Belt-and-suspenders: skip any message that ended up in a
            # deleted / recoverable-deleted folder (edge case inside inbox).
            if _is_deleted_message(msg, deleted_folder_ids):
                print(f"[EMAIL] Skipping deleted email: {msg.get('subject', '')[:60]}")
                continue
            messages.append(msg)

        url = data.get("@odata.nextLink")   # None when last page

    print(f"[EMAIL] Fetched {len(messages)} inbox email(s) from the last {days} day(s).")
    return messages, headers, base_url



# ─────────────────────────────────────────────────────────────────────────────
# Public: process only the newest email (backwards-compat for direct runs)
# ─────────────────────────────────────────────────────────────────────────────
def process_latest_email() -> dict:
    messages, headers, base_url = fetch_emails_last_n_days(days=1)
    if not messages:
        return {}
    return process_email(messages[0], headers, base_url)


# ─────────────────────────────────────────────────────────────────────────────
# Public: send an email with PDF attachments via Microsoft Graph API
# ─────────────────────────────────────────────────────────────────────────────
def send_upload_email(
    subject: str,
    body_text: str,
    pdf_paths: list,
    to_email: str = USER_EMAIL,
) -> bool:
    """
    Send an email FROM the automation mailbox (USER_EMAIL) TO `to_email`
    with the given subject, plain-text body, and PDF file attachments.

    Uses the same Azure AD app-only credentials already configured for
    reading mail (Mail.Send permission required on the app registration).

    Args:
        subject   : email subject line
        body_text : plain-text email body
        pdf_paths : list of absolute paths to PDF files to attach
        to_email  : recipient email address (defaults to USER_EMAIL)

    Returns True on success, False on failure.
    """
    try:
        token   = _get_access_token()
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type":  "application/json",
        }

        # Build attachment list
        attachments = []
        for pdf_path in pdf_paths or []:
            if not pdf_path or not os.path.isfile(pdf_path):
                print(f"  [SEND-EMAIL] Skipping missing file: {pdf_path}")
                continue
            with open(pdf_path, "rb") as f:
                content_bytes = base64.b64encode(f.read()).decode("utf-8")
            attachments.append({
                "@odata.type":  "#microsoft.graph.fileAttachment",
                "name":         os.path.basename(pdf_path),
                "contentType":  "application/pdf",
                "contentBytes": content_bytes,
            })

        if not attachments:
            print("  [SEND-EMAIL] No valid PDF attachments – email not sent")
            return False

        mail_body = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "Text",
                    "content":     body_text,
                },
                "toRecipients": [
                    {"emailAddress": {"address": to_email}},
                ],
                "attachments": attachments,
            },
            "saveToSentItems": "true",
        }

        send_url = f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/sendMail"
        resp = requests.post(send_url, headers=headers, json=mail_body, timeout=60)

        if resp.status_code == 202:
            print(f"  [SEND-EMAIL] ✅ Email sent successfully to {to_email}")
            print(f"  [SEND-EMAIL]    Subject: {subject}")
            print(f"  [SEND-EMAIL]    Attachments: {[os.path.basename(p) for p in pdf_paths if p and os.path.isfile(p)]}")
            return True
        else:
            print(f"  [SEND-EMAIL] ❌ Failed to send email: HTTP {resp.status_code}")
            print(f"  [SEND-EMAIL]    Response: {resp.text[:500]}")
            return False

    except Exception as exc:
        print(f"  [SEND-EMAIL] ❌ Error sending email: {exc}")
        return False


# ─────────────────────────────────────────────────────────────────────────────
# Public: send an HTML email (no attachments) via Microsoft Graph API
# ─────────────────────────────────────────────────────────────────────────────
def send_html_email(
    subject: str,
    html_body: str,
    to_emails,
    cc_emails=None,
    pdf_attachments=None,
) -> bool:
    """
    Send an HTML email FROM the automation mailbox (USER_EMAIL) to one or more
    recipients, optionally with PDF attachments.

    Args:
        subject         : email subject line
        html_body       : HTML email body
        to_emails       : a single address (str) or a list of addresses
        cc_emails       : optional single address or list of CC addresses
        pdf_attachments : optional list of absolute paths to PDF files to attach

    Returns True on success, False on failure.
    """
    def _as_list(v):
        if not v:
            return []
        if isinstance(v, str):
            return [v]
        return [str(a).strip() for a in v if str(a).strip()]

    to_list = _as_list(to_emails)
    cc_list = _as_list(cc_emails)
    if not to_list:
        print("  [SEND-HTML] No recipients – email not sent")
        return False

    try:
        token = _get_access_token()
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type":  "application/json",
        }
        message = {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": [{"emailAddress": {"address": a}} for a in to_list],
        }
        if cc_list:
            message["ccRecipients"] = [{"emailAddress": {"address": a}} for a in cc_list]

        # Attach PDFs if provided
        if pdf_attachments:
            attachments = []
            for pdf_path in pdf_attachments:
                if not pdf_path or not os.path.isfile(pdf_path):
                    continue
                with open(pdf_path, "rb") as f:
                    content_bytes = base64.b64encode(f.read()).decode("utf-8")
                attachments.append({
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": os.path.basename(pdf_path),
                    "contentType": "application/pdf",
                    "contentBytes": content_bytes,
                })
            if attachments:
                message["attachments"] = attachments

        mail_body = {"message": message, "saveToSentItems": "true"}
        send_url = f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/sendMail"
        resp = requests.post(send_url, headers=headers, json=mail_body, timeout=60)

        if resp.status_code in (200, 202):
            att_names = [os.path.basename(p) for p in (pdf_attachments or []) if p and os.path.isfile(p)]
            print(f"  [SEND-HTML] ✅ Email sent to {', '.join(to_list)}")
            print(f"  [SEND-HTML]    Subject: {subject}")
            if att_names:
                print(f"  [SEND-HTML]    Attachments: {att_names}")
            return True
        print(f"  [SEND-HTML] ❌ Failed: HTTP {resp.status_code} {resp.text[:400]}")
        return False
    except Exception as exc:
        print(f"  [SEND-HTML] ❌ Error sending email: {exc}")
        return False


# ─────────────────────────────────────────────────────────────────────────────
# Standalone test
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    msgs, hdrs, base = fetch_emails_last_n_days(days=2)
    for m in msgs:
        result = process_email(m, hdrs, base)
        print(json.dumps({
            "message_id": result["message_id"],
            "title":      result["title"],
            "has_tables": has_required_tables(result["parsed_tables"]),
        }, indent=2))
