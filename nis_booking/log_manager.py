"""
log_manager.py
===============
File-based logging helper for the DSM-NIS pipeline.

Behavior
--------
 - Base path read from config.LOG_BASE_PATH (only used when
   config.LOG_USE_CUSTOM_PATH is True; otherwise CWD is used).
 - Creates: <base>/logs/<svp_safe_name>/<from_date>_to_<to_date>.txt
 - Appends timestamped, categorised lines for every action.

Thread-local context
--------------------
Call ``set_context(svp, from_date, to_date)`` once at the start of each SPV
run.  Any module can then call the short helpers ``log(msg)`` / ``log_error()``
/ ``log_warn()`` / ``log_info()`` without repeating svp/date every time.
Modules that want explicit control can still call ``log_action(svp, ...)``
directly.
"""

import os
import re
import traceback
from datetime import datetime
from threading import local
from typing import Optional

import config

# ---------------------------------------------------------------------------
# Thread-local run context
# ---------------------------------------------------------------------------
_ctx = local()


def set_context(svp_name: str, from_date: str, to_date: str) -> None:
    """Set the current SVP + date-range context for this thread.

    Logging is silently suppressed until a valid (non-empty) SVP name and
    date range have been provided via this call.
    """
    _ctx.svp_name  = (svp_name or "").strip()
    _ctx.from_date = (from_date or "").strip()
    _ctx.to_date   = (to_date   or "").strip()


def clear_context() -> None:
    """Clear the thread-local context (call at the very end of a run)."""
    _ctx.svp_name  = ""
    _ctx.from_date = ""
    _ctx.to_date   = ""


def _ctx_or_default(svp_name, from_date, to_date):
    svp  = svp_name  if svp_name  is not None else getattr(_ctx, "svp_name",  "")
    fd   = from_date if from_date is not None else getattr(_ctx, "from_date", "")
    td   = to_date   if to_date   is not None else getattr(_ctx, "to_date",   "")
    return svp, fd, td


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------
def _sanitize(name: str) -> str:
    if not name or not name.strip():
        return ""
    name = re.sub(r'[\\/:*?"<>|]', "_", str(name))
    name = re.sub(r"\s+", " ", name).strip()
    return name


def init_log_path(svp_name: str, from_date: str, to_date: str,
                  base_path: Optional[str] = None) -> str:
    """Return full path to the log file and ensure directories exist.

    Returns an empty string when *svp_name* is missing — this tells the
    caller (``_write``) to silently skip the log line so no "unknown_spv"
    folder is ever created.

    The log file name is always ``<from_date>_to_<to_date>.txt`` which
    ensures a single, deterministic file per unique date-range.

    Resolution order for base directory:
      1. Explicit ``base_path`` argument (tests / manual calls).
      2. ``config.LOG_BASE_PATH`` — only when ``config.LOG_USE_CUSTOM_PATH`` is True.
      3. Current working directory (default).
    """
    safe_svp = _sanitize(svp_name)
    if not safe_svp:
        return ""  # no SPV yet → don't create any file

    if base_path:
        base = str(base_path)
    elif getattr(config, "LOG_USE_CUSTOM_PATH", False):
        base = getattr(config, "LOG_BASE_PATH", "") or os.getcwd()
    else:
        base = os.getcwd()

    base = str(base).rstrip("\\/")
    logs_root  = os.path.join(base, "logs")
    svp_folder = os.path.join(logs_root, safe_svp)
    os.makedirs(svp_folder, exist_ok=True)

    fd = (from_date or "").strip()
    td = (to_date   or "").strip()
    if fd and td:
        fd_safe = re.sub(r'[\\/:*?"<>|]', "_", fd)
        td_safe = re.sub(r'[\\/:*?"<>|]', "_", td)
        fname = f"{fd_safe}_to_{td_safe}.txt"
    else:
        # Without dates we still write into the SVP folder but with a
        # timestamped name so we never lose data.
        fname = datetime.now().strftime("%Y%m%d_%H%M%S.txt")

    full = os.path.join(svp_folder, fname)
    if not os.path.exists(full):
        try:
            with open(full, "a", encoding="utf-8"):
                pass
        except Exception:
            return ""
    return full


def _write(svp_name: str, from_date: str, to_date: str,
           level: str, message: str,
           base_path: Optional[str] = None) -> str:
    """Low-level: format + append one line to the log file.

    Silently skips (returns ``""``) when no valid SPV context has been set
    yet — this prevents an ``unknown_spv`` folder from being created.
    """
    try:
        path = init_log_path(svp_name, from_date, to_date, base_path=base_path)
        if not path:
            return ""  # context not ready → skip silently
        ts   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{ts}] [{level:<5}] {message}\n"
        with open(path, "a", encoding="utf-8") as fh:
            fh.write(line)
        return path
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# Public API – explicit svp/dates (always available)
# ---------------------------------------------------------------------------
def log_action(svp_name: str, from_date: str, to_date: str,
               message: str, base_path: Optional[str] = None) -> str:
    """Append a timestamped INFO line to the svp/date log file."""
    return _write(svp_name, from_date, to_date, "INFO", message, base_path)


def log_error_explicit(svp_name: str, from_date: str, to_date: str,
                       message: str, exc: Exception = None,
                       base_path: Optional[str] = None) -> str:
    """Append an ERROR line; if *exc* is provided, include the full traceback."""
    full_msg = message
    if exc is not None:
        tb = traceback.format_exc()
        full_msg = f"{message} | exception={type(exc).__name__}: {exc} | traceback:\n{tb}"
    return _write(svp_name, from_date, to_date, "ERROR", full_msg, base_path)


def log_warn_explicit(svp_name: str, from_date: str, to_date: str,
                      message: str, base_path: Optional[str] = None) -> str:
    """Append a WARN line."""
    return _write(svp_name, from_date, to_date, "WARN", message, base_path)


# ---------------------------------------------------------------------------
# Public API – context-aware (use set_context() first)
# ---------------------------------------------------------------------------
def log(message: str,
        svp_name: Optional[str] = None,
        from_date: Optional[str] = None,
        to_date: Optional[str] = None,
        base_path: Optional[str] = None) -> str:
    """Append an INFO line using context (or explicit args)."""
    svp, fd, td = _ctx_or_default(svp_name, from_date, to_date)
    return _write(svp, fd, td, "INFO", message, base_path)


def log_error(message: str,
              exc: Exception = None,
              svp_name: Optional[str] = None,
              from_date: Optional[str] = None,
              to_date: Optional[str] = None,
              base_path: Optional[str] = None) -> str:
    """Append an ERROR line with optional exception traceback."""
    svp, fd, td = _ctx_or_default(svp_name, from_date, to_date)
    full_msg = message
    if exc is not None:
        tb = traceback.format_exc()
        full_msg = f"{message} | exception={type(exc).__name__}: {exc} | traceback:\n{tb}"
    return _write(svp, fd, td, "ERROR", full_msg, base_path)


def log_warn(message: str,
             svp_name: Optional[str] = None,
             from_date: Optional[str] = None,
             to_date: Optional[str] = None,
             base_path: Optional[str] = None) -> str:
    """Append a WARN line."""
    svp, fd, td = _ctx_or_default(svp_name, from_date, to_date)
    return _write(svp, fd, td, "WARN", message, base_path)


# ---------------------------------------------------------------------------
# Compatibility aliases
# ---------------------------------------------------------------------------
write = log_action
