"""
sharepoint_upload.py
====================
Helper functions for managing files and folders on SharePoint / OneDrive
via the Microsoft Graph API.

All operations target the same SharePoint site + drive that the pipeline
already uses for vendor master data and the NIS booking tracker Excel file.

Folder structure on SharePoint:
    <FOLDER_PATH>/booking/<SPV name>/<from_date to to_date>/
        email.pdf
        DSM_*.pdf
        sap_checklist_*.pdf
        nis_checklist_*.pdf
        upload_sap_charges.pdf
        upload_nis_charges.pdf
"""

import io
import os
import time
import urllib.parse
import msal
import requests
import urllib3
import config
import log_manager as lm

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
http_session = requests.Session()
http_session.verify = False

# ─────────────────────────────────────────────────────────────────────────────
# Configuration  (reuses existing SharePoint app registration from config.py)
# ─────────────────────────────────────────────────────────────────────────────
TENANT_ID     = config.SP_TENANT_ID
CLIENT_ID     = config.SP_CLIENT_ID
CLIENT_SECRET = config.SP_CLIENT_SECRET
AUTHORITY     = config.SP_AUTHORITY
SCOPES        = config.GRAPH_SCOPE

TENANT_HOST = config.TENANT_HOST
SITE_NAME   = config.SITE_NAME
DRIVE_NAME  = config.DRIVE_NAME

# Base folder inside the SharePoint document library where booking folders
# will be created.  Final path:  <SP_BOOKING_BASE>/booking/<SPV>/<dates>/
SP_BOOKING_BASE = getattr(config, "SP_BOOKING_BASE", config.FOLDER_PATH)


# ─────────────────────────────────────────────────────────────────────────────
# Auth
# ─────────────────────────────────────────────────────────────────────────────
_cached_token: dict = {"token": "", "expires": 0.0}


def _get_token() -> str:
    """Acquire (or reuse cached) app-only bearer token."""
    now = time.time()
    if _cached_token["token"] and now < _cached_token["expires"] - 60:
        return _cached_token["token"]

    app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=SCOPES)
    if "access_token" not in result:
        raise RuntimeError(f"MSAL token error: {result}")
    _cached_token["token"]   = result["access_token"]
    _cached_token["expires"] = now + result.get("expires_in", 3600)
    return _cached_token["token"]


def _headers() -> dict:
    return {"Authorization": f"Bearer {_get_token()}"}


def _encode_sp_path(path: str) -> str:
    """URL-encode each segment of a SharePoint path (keeps '/' literal)."""
    return "/".join(
        urllib.parse.quote(seg, safe="") for seg in path.split("/") if seg
    )


# ─────────────────────────────────────────────────────────────────────────────
# Site + Drive resolution  (cached per process lifetime)
# ─────────────────────────────────────────────────────────────────────────────
_drive_id_cache: str = ""


def _get_drive_id() -> str:
    global _drive_id_cache
    if _drive_id_cache:
        return _drive_id_cache

    h = _headers()
    site = http_session.get(
        f"https://graph.microsoft.com/v1.0/sites/{TENANT_HOST}:/sites/{SITE_NAME}",
        headers=h,
    )
    site.raise_for_status()
    site_id = site.json()["id"]

    drives = http_session.get(
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives",
        headers=h,
    )
    drives.raise_for_status()
    for d in drives.json()["value"]:
        if d["name"] == DRIVE_NAME:
            _drive_id_cache = d["id"]
            return _drive_id_cache

    raise RuntimeError(f"Drive '{DRIVE_NAME}' not found on site '{SITE_NAME}'")


# ─────────────────────────────────────────────────────────────────────────────
# Core operations
# ─────────────────────────────────────────────────────────────────────────────
def _graph_drive_url() -> str:
    return f"https://graph.microsoft.com/v1.0/drives/{_get_drive_id()}"


def ensure_sp_folder(folder_path: str) -> str:
    """
    Create a folder on SharePoint if it does not already exist.
    """
    drive_url = _graph_drive_url()
    h = _headers()

    lm.log(f"SP_ENSURE_FOLDER path={folder_path}")
    # Check if folder already exists
    check_url = f"{drive_url}/root:/{_encode_sp_path(folder_path)}"
    resp = http_session.get(check_url, headers=h)
    if resp.status_code == 200:
        lm.log(f"SP_FOLDER_EXISTS path={folder_path}")
        return folder_path  # already exists

    # Create by walking path segments and creating each level
    segments = [s for s in folder_path.replace("\\", "/").split("/") if s]
    built_path = ""

    for seg in segments:
        parent_url = (
            f"{drive_url}/root:/{_encode_sp_path(built_path)}:" if built_path else f"{drive_url}/root"
        )
        body = {
            "name": seg,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "fail",  # skip if exists
        }
        create_resp = http_session.post(
            f"{parent_url}/children",
            headers={**h, "Content-Type": "application/json"},
            json=body,
        )
        if create_resp.status_code == 409:
            pass  # Already exists — continue
        elif create_resp.status_code not in (200, 201):
            msg = (f"  [SP] Folder create note for '{seg}': "
                   f"HTTP {create_resp.status_code} – {create_resp.text[:200]}")
            print(msg)
            lm.log_warn(f"SP_FOLDER_CREATE_NOTE seg={seg} "
                        f"status={create_resp.status_code} body={create_resp.text[:200]}")

        built_path = f"{built_path}/{seg}" if built_path else seg

    lm.log(f"SP_FOLDER_CREATED path={folder_path}")
    print(f"  [SP] Folder ensured: {folder_path}")
    return folder_path


def upload_file_to_sp(local_path: str, sp_folder_path: str, sp_filename: str = "") -> str:
    """
    Upload a local file to a SharePoint folder.

    Args:
        local_path:     absolute path to the local file
        sp_folder_path: SharePoint folder (relative to drive root)
        sp_filename:    filename on SharePoint (defaults to basename of local_path)

    Returns the SharePoint item path on success, or "" on failure.
    """
    if not os.path.isfile(local_path):
        msg = f"  [SP-UPLOAD] File not found: {local_path}"
        print(msg)
        lm.log_error(f"SP_UPLOAD_FAILED file_not_found local={local_path}")
        return ""

    filename  = sp_filename or os.path.basename(local_path)
    file_size = os.path.getsize(local_path)
    drive_url = _graph_drive_url()
    h = _headers()

    sp_path      = f"{sp_folder_path}/{filename}".replace("\\", "/")
    encoded_path = _encode_sp_path(sp_path)
    upload_url   = f"{drive_url}/root:/{encoded_path}:/content"

    lm.log(f"SP_UPLOAD_START local={local_path} sp={sp_path} size={file_size}")
    try:
        with open(local_path, "rb") as f:
            if file_size <= 4 * 1024 * 1024:
                resp = http_session.put(
                    upload_url,
                    headers={**h, "Content-Type": "application/octet-stream"},
                    data=f,
                    timeout=120,
                )
            else:
                # Create upload session for large files
                session_resp = http_session.post(
                    f"{drive_url}/root:/{encoded_path}:/createUploadSession",
                    headers={**h, "Content-Type": "application/json"},
                    json={"item": {"@microsoft.graph.conflictBehavior": "replace"}},
                    timeout=60,
                )
                session_resp.raise_for_status()
                upload_session_url = session_resp.json()["uploadUrl"]
                resp = _upload_large_file(f, upload_session_url, file_size)

        if resp.status_code in (200, 201):
            print(f"  [SP-UPLOAD] ✅ {filename} → {sp_folder_path}")
            lm.log(f"SP_UPLOAD_OK sp={sp_path}")
            return sp_path
        else:
            msg = f"HTTP {resp.status_code}: {resp.text[:300]}"
            print(f"  [SP-UPLOAD] ❌ {msg}")
            lm.log_error(f"SP_UPLOAD_FAILED sp={sp_path} {msg}")
            return ""

    except Exception as exc:
        print(f"  [SP-UPLOAD] ❌ Error uploading {filename}: {exc}")
        lm.log_error(f"SP_UPLOAD_EXCEPTION sp={sp_path}", exc=exc)
        return ""


def _upload_large_file(file_obj, upload_url: str, file_size: int):
    """Upload a large file in 10 MB chunks via an upload session."""
    chunk_size = 10 * 1024 * 1024
    offset = 0
    while offset < file_size:
        chunk = file_obj.read(chunk_size)
        end = offset + len(chunk) - 1
        resp = http_session.put(
            upload_url,
            headers={
                "Content-Length": str(len(chunk)),
                "Content-Range": f"bytes {offset}-{end}/{file_size}",
            },
            data=chunk,
            timeout=120,
        )
        if resp.status_code not in (200, 201, 202):
            return resp
        offset = end + 1
    return resp


def upload_bytes_to_sp(data: bytes, sp_folder_path: str, sp_filename: str) -> str:
    """
    Upload in-memory bytes to a SharePoint folder.

    Returns the SharePoint item path on success, or "" on failure.
    """
    drive_url = _graph_drive_url()
    h = _headers()
    sp_path = f"{sp_folder_path}/{sp_filename}".replace("\\", "/")
    encoded_path = _encode_sp_path(sp_path)
    upload_url = f"{drive_url}/root:/{encoded_path}:/content"

    try:
        resp = http_session.put(
            upload_url,
            headers={**h, "Content-Type": "application/octet-stream"},
            data=data,
            timeout=120,
        )
        if resp.status_code in (200, 201):
            print(f"  [SP-UPLOAD] ✅ {sp_filename} → {sp_folder_path}")
            return sp_path
        else:
            print(f"  [SP-UPLOAD] ❌ HTTP {resp.status_code}: {resp.text[:300]}")
            return ""
    except Exception as exc:
        print(f"  [SP-UPLOAD] ❌ Error uploading {sp_filename}: {exc}")
        return ""


def download_file_from_sp(sp_file_path: str, local_dest: str) -> str:
    """
    Download a file from SharePoint to a local path.

    Args:
        sp_file_path: path relative to drive root (e.g. "folder/sub/file.pdf")
        local_dest:   local file path to write to

    Returns the local path on success, or "" on failure.
    """
    drive_url = _graph_drive_url()
    h = _headers()
    url = f"{drive_url}/root:/{_encode_sp_path(sp_file_path)}:/content"

    lm.log(f"SP_DOWNLOAD_START sp={sp_file_path} dest={local_dest}")
    try:
        os.makedirs(os.path.dirname(local_dest) or ".", exist_ok=True)
        resp = http_session.get(url, headers=h, stream=True, timeout=120)
        resp.raise_for_status()
        with open(local_dest, "wb") as f:
            for chunk in resp.iter_content(8192):
                if chunk:
                    f.write(chunk)
        print(f"  [SP-DOWNLOAD] ✅ {sp_file_path} → {local_dest}")
        lm.log(f"SP_DOWNLOAD_OK sp={sp_file_path} dest={local_dest}")
        return local_dest
    except Exception as exc:
        print(f"  [SP-DOWNLOAD] ❌ Error downloading {sp_file_path}: {exc}")
        lm.log_error(f"SP_DOWNLOAD_FAILED sp={sp_file_path} dest={local_dest}", exc=exc)
        return ""


def list_sp_folder(sp_folder_path: str) -> list:
    """
    List file/folder names in a SharePoint folder.

    Returns a list of dicts:  [{"name": "file.pdf", "is_folder": False, "size": 1234}, ...]
    """
    drive_url = _graph_drive_url()
    h = _headers()
    url = f"{drive_url}/root:/{_encode_sp_path(sp_folder_path)}:/children?$select=name,size,folder"

    items = []
    try:
        while url:
            resp = http_session.get(url, headers=h)
            if resp.status_code == 404:
                lm.log_warn(f"SP_LIST_NOT_FOUND path={sp_folder_path}")
                return []
            resp.raise_for_status()
            data = resp.json()
            for item in data.get("value", []):
                items.append({
                    "name":      item["name"],
                    "is_folder": "folder" in item,
                    "size":      item.get("size", 0),
                })
            url = data.get("@odata.nextLink")
        lm.log(f"SP_LIST_OK path={sp_folder_path} count={len(items)}")
    except Exception as exc:
        print(f"  [SP-LIST] ❌ Error listing {sp_folder_path}: {exc}")
        lm.log_error(f"SP_LIST_FAILED path={sp_folder_path}", exc=exc)
    return items


def sp_file_exists(sp_file_path: str) -> bool:
    """Check if a file exists on SharePoint."""
    drive_url = _graph_drive_url()
    h = _headers()
    resp = http_session.get(f"{drive_url}/root:/{_encode_sp_path(sp_file_path)}", headers=h)
    return resp.status_code == 200


# ─────────────────────────────────────────────────────────────────────────────
# Convenience: build the booking folder path string
# ─────────────────────────────────────────────────────────────────────────────
def build_booking_folder_path(spv_name: str, from_date: str, to_date: str) -> str:
    """
    Return the SharePoint folder path for a booking:
        <SP_BOOKING_BASE>/booking/<SPV name>/<from_date to to_date>

    Characters unsafe for SharePoint folder names are replaced with underscores.
    """
    import re as _re

    def _sanitize(name: str) -> str:
        name = _re.sub(r'[\\/:*?"<>|#%]', "_", name)
        name = _re.sub(r"\s+", " ", name).strip()
        return name

    date_folder = _sanitize(f"{from_date} to {to_date}")
    base = SP_BOOKING_BASE.rstrip("/")

    if spv_name:
        spv_folder = _sanitize(spv_name)
        return f"{base}/booking/{spv_folder}/{date_folder}"
    else:
        return f"{base}/booking/{date_folder}"


def build_sap_checklist_folder_path(from_date: str, to_date: str) -> str:
    """Return the shared SharePoint folder for Debit/Credit SAP checklist PDFs:
        <SP_BOOKING_BASE>/booking/<from_date>_<to_date>

    Keyed only by the date range of the original 5-row email, so every SAP
    checklist PDF for that week (typically 8–12) lands in one folder.
    """
    import re as _re

    def _sanitize(name: str) -> str:
        name = _re.sub(r'[\\/:*?"<>|#%]', "_", name)
        name = _re.sub(r"\s+", " ", name).strip()
        return name

    base = SP_BOOKING_BASE.rstrip("/")
    date_folder = _sanitize(f"{from_date}_{to_date}")
    return f"{base}/booking/{date_folder}"
