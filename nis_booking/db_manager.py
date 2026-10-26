"""
db_manager.py
=============
Persistence layer for the DSM-NIS booking automation pipeline.

Supports **two** backends (selected via ``DB_BACKEND`` env / config):
  • ``sqlite``  – local SQLite file (default, original behaviour)
  • ``postgres`` – remote PostgreSQL server (Azure, RDS, etc.)

When ``DB_BACKEND=postgres``, tables are created automatically on first connect
if they don't already exist.

Design
------
* ``email_runs``  – one record per email (keyed by the email ``message_id``).
* ``booking_rows`` – one record per SPV row inside an email. Each row gets a
  **unique** ``row_id`` while sharing the same ``message_id``. This is what
  makes every SPV row an *independent* booking unit: the pipeline picks one
  pending row, books it, marks it ``done`` and only then moves to the next row.
* ``run_steps`` – fine-grained step tracking (extraction, classification,
  vendor read, NIS booking, tracker update …) used for crash-resume.
* ``nis_completion_recipients`` – notification recipients.

The public function surface is a drop-in replacement for the old
`sheet_db_manager_sp` module, with additional per-row helpers
(``seed_booking_rows``, ``get_next_pending_booking_row``,
``mark_booking_row_done`` …).
"""

from __future__ import annotations

import json
import os
import sqlite3
import threading
from datetime import datetime
from pathlib import Path
from typing import Optional

def _load_env():
    _candidates = [
        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"),
        os.path.join(os.getcwd(), ".env"),
    ]
    for _path in _candidates:
        if os.path.isfile(_path):
            try:
                with open(_path, "r", encoding="utf-8") as _f:
                    for _line in _f:
                        _line = _line.strip()
                        if _line and not _line.startswith("#") and "=" in _line:
                            _k, _v = _line.split("=", 1)
                            _k = _k.strip()
                            _v = _v.strip().strip("'\"")
                            if _k not in os.environ:
                                os.environ[_k] = _v
            except Exception:
                pass

# Load environment variables from .env file
try:
    from dotenv import load_dotenv
    _root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    load_dotenv(os.path.join(_root_dir, ".env"))
    load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))
    load_dotenv()
except Exception:
    pass

_load_env()

# PostgreSQL driver — optional; only imported when backend is postgres.
try:
    import psycopg2
    import psycopg2.extras
except ImportError:
    psycopg2 = None  # type: ignore[assignment]

# ── Step-name constants ──────────────────────────────────────────────────────
STEP_READ_EMAIL           = "read_email"
STEP_EMAIL_EXTRACTION     = "email_extraction"
STEP_EMAIL_CLASSIFICATION = "email_classification"
STEP_EMAIL_VALIDATION     = "email_validation"
STEP_READ_VENDOR          = "read_vendor_master"
STEP_UPDATE_TRACKER       = "update_excel_tracker"
STEP_NIS_BOOKING          = "nis_booking"


def step_nis(field: str) -> str:
    return f"nis_booking_{field}"


def step_sap(field: str) -> str:
    return f"sap_automation_{field}"


# ── Status constants ─────────────────────────────────────────────────────────
STATUS_PENDING  = "pending"
STATUS_RUNNING  = "running"
STATUS_DONE     = "done"
STATUS_FAILED   = "failed"
STATUS_SKIPPED  = "skipped"
STATUS_NIS_DONE = "nis_done"  # NIS booking complete, SAP steps still pending

_FIXED_STEPS = [
    STEP_READ_EMAIL,
    STEP_EMAIL_EXTRACTION,
    STEP_EMAIL_CLASSIFICATION,
    STEP_EMAIL_VALIDATION,
    STEP_READ_VENDOR,
    STEP_UPDATE_TRACKER,
]

# ── Schema columns (kept explicit so reads/writes stay predictable) ──────────
_EMAIL_RUNS_COLUMNS = [
    "id", "message_id", "subject", "received_at", "spv_name", "status",
    "body_html", "metadata_json", "spv_dsm_json", "vendor_data_json",
    "positive_fields_json", "negative_fields_json", "downloaded_pdf_names",
    "nis_checklist_id", "nis_checklist_pdf_uploaded", "checklist_value",
    "sap_checklist_pdf_uploaded", "booking_folder_path", "failure_email_sent",
    "nis_completion_email_sent_at", "created_at", "started_at", "finished_at",
    "error_message",
]

_BOOKING_ROWS_COLUMNS = [
    "row_id", "run_id", "message_id", "row_index", "spv_name", "status",
    "spv_data_json", "metadata_json", "positive_fields_json",
    "negative_fields_json", "vendor_data_json", "invoice_date", "invoice_number",
    "nis_checklist_id", "nis_checklist_pdf_uploaded", "checklist_value",
    "sap_checklist_pdf_uploaded", "downloaded_pdf_names", "booking_folder_path",
    "retry_count", "error_message", "nis_completion_email_sent_at",
    "failure_email_sent",
    # ── HOD / FC approval tracking ───────────────────────────────────────────
    "hod_approval_status", "hod_approval_at", "hod_approval_message_id",
    "hod_email_sent", "hod_email_sent_at",
    "fc_approval_status", "fc_approval_at", "fc_approval_message_id",
    "fc_checklist_pdf_uploaded", "fc_email_sent", "fc_email_sent_at",
    "pdf_merge_status", "pdf_merge_at", "merged_pdf_name",
    "completion_email_sent", "completion_email_sent_at",
    # ─────────────────────────────────────────────────────────────────────────
    "created_at", "started_at", "finished_at",
]

_RUN_STEPS_COLUMNS = [
    "id", "run_id", "step_name", "status", "detail", "started_at",
    "finished_at", "error", "retry_count",
]

# Debit/Credit Note email tracking (5-row DSM table trigger). One record per
# DSM email (UNIQUE message_id): the email we send to the finance team plus the
# reply carrying the Debit/Credit Note PDFs.
_DEBIT_CREDIT_COLUMNS = [
    "id", "workflow_id", "run_id", "message_id", "spv_name", "week_no",
    "from_date", "to_date", "dsm_amount", "company_code", "due_date_cn_dn",
    "recipients", "sent_subject", "email_status", "sent_at", "reply_message_id",
    "reply_received_at", "reply_pdf_names", "reply_pdfs_uploaded",
    "booking_folder_path", "error_message", "created_at",
]

# One record per finance-reply Debit/Credit Note PDF. Tracks extraction →
# vendor resolution → SAP checklist creation → SharePoint upload.
_DEBIT_CREDIT_PDF_COLUMNS = [
    "id", "workflow_id", "dc_email_id", "run_id", "reply_message_id",
    "pdf_name", "spv_from_filename", "matched_spv_name", "vendor_code",
    "company_code", "cost_center", "plant", "ref_number", "amount",
    "amount_text", "doc_date", "note_type", "dsm_week", "extract_status",
    "sap_status", "sap_checklist_number", "sap_pdf_name", "sap_pdf_sp_path",
    "sap_folder_path", "from_date", "to_date", "error_message",
    "created_at", "updated_at",
]

_LOCK = threading.RLock()

# ─────────────────────────────────────────────────────────────────────────────
# Backend resolution (read directly from environment / .env)
# ─────────────────────────────────────────────────────────────────────────────
_DB_BACKEND = (os.getenv("DB_BACKEND") or "sqlite").strip().lower()

# SQLite path resolution (used only when backend == "sqlite")
_DEFAULT_DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "nis_booking.db")
_DB_PATH = Path(
    os.getenv("SQLITE_DB_PATH")
    or os.getenv("DSM_SHARED_DB_PATH")
    or _DEFAULT_DB_PATH
)

# PostgreSQL connection params (used only when backend == "postgres")
_PG_DSN      = os.getenv("POSTGRES_DSN")
_PG_HOST     = os.getenv("POSTGRES_HOST")
_PG_PORT     = int(os.getenv("POSTGRES_PORT")) if os.getenv("POSTGRES_PORT") else 5432
_PG_DB       = os.getenv("POSTGRES_DB")
_PG_USER     = os.getenv("POSTGRES_USER")
_PG_PASSWORD = os.getenv("POSTGRES_PASSWORD")
_PG_SSLMODE  = os.getenv("POSTGRES_SSLMODE")

_CONN = None  # Shared connection (SQLite Connection or psycopg2 connection)


# ─────────────────────────────────────────────────────────────────────────────
# Dual-backend helpers
# ─────────────────────────────────────────────────────────────────────────────
def _is_postgres() -> bool:
    return _DB_BACKEND == "postgres"


def _ph(sql: str) -> str:
    """Translate SQLite-style ``?`` placeholders to PostgreSQL ``%s``."""
    if _is_postgres():
        return sql.replace("?", "%s")
    return sql


def _now() -> str:
    return datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")


def _connect():
    """Return (and cache) the shared database connection."""
    global _CONN
    if _CONN is not None:
        if _is_postgres():
            try:
                cur = _CONN.cursor()
                cur.execute("SELECT 1")
                cur.close()
            except Exception:
                try:
                    _CONN.close()
                except Exception:
                    pass
                _CONN = None
        else:
            return _CONN

    if _CONN is not None:
        return _CONN

    if _is_postgres():
        if psycopg2 is None:
            raise ImportError(
                "psycopg2 is required when DB_BACKEND=postgres. "
                "Install with: pip install psycopg2-binary"
            )
        if _PG_DSN:
            conn = psycopg2.connect(_PG_DSN)
        else:
            conn = psycopg2.connect(
                host=_PG_HOST,
                port=_PG_PORT,
                dbname=_PG_DB,
                user=_PG_USER,
                password=_PG_PASSWORD,
                sslmode=_PG_SSLMODE,
            )
        conn.autocommit = False
    else:
        _DB_PATH.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(str(_DB_PATH), check_same_thread=False, timeout=60)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("PRAGMA foreign_keys=ON;")

    _CONN = conn
    return _CONN


class _PGConnectionWrapper:
    """Thin wrapper around a psycopg2 connection that mimics sqlite3.Connection
    usage patterns so the rest of the module needs minimal changes.

    Key translations:
      - ``?`` → ``%s`` (placeholder)
      - ``INSERT OR IGNORE`` → ``INSERT ... ON CONFLICT DO NOTHING``
      - ``with conn:`` context manager commits on success / rolls back on error
      - ``.execute()`` returns a cursor whose ``.fetchone()`` returns a dict
      - For INSERT statements, automatically appends RETURNING id to support lastrowid
    """

    def __init__(self, pg_conn):
        self._conn = pg_conn

    def _translate(self, sql: str) -> str:
        sql = sql.replace("?", "%s")
        # In psycopg2, literal % in SQL must be %% (but only outside of %s placeholders).
        # We handle this by replacing single % that are NOT followed by 's' (our placeholders).
        # Strategy: replace %s with a sentinel, escape remaining %, restore sentinel.
        _SENTINEL = "\x00PH\x00"
        sql = sql.replace("%s", _SENTINEL)
        sql = sql.replace("%", "%%")
        sql = sql.replace(_SENTINEL, "%s")
        # SQLite-specific INSERT OR IGNORE → PG ON CONFLICT DO NOTHING
        import re as _re
        if _re.search(r"INSERT\s+OR\s+IGNORE", sql, flags=_re.IGNORECASE):
            sql = _re.sub(r"INSERT\s+OR\s+IGNORE", "INSERT", sql, flags=_re.IGNORECASE)
            sql = sql.rstrip().rstrip(";")
            sql += " ON CONFLICT DO NOTHING"
        return sql

    def execute(self, sql, params=None):
        sql = self._translate(sql)
        cur = self._conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute(sql, params or ())
        return cur

    def commit(self):
        self._conn.commit()

    def rollback(self):
        self._conn.rollback()

    def close(self):
        self._conn.close()

    @property
    def rowcount(self):
        return self._conn.cursor().rowcount

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is None:
            self._conn.commit()
        else:
            self._conn.rollback()
        return False


def _get_conn():
    """Return a connection object suitable for direct use by the module.
    For PostgreSQL: returns a wrapper that translates SQL on the fly.
    For SQLite: returns the raw sqlite3 connection (unchanged behaviour)."""
    conn = _connect()
    if _is_postgres():
        # Return wrapper only if not already wrapped
        if not isinstance(conn, _PGConnectionWrapper):
            wrapped = _PGConnectionWrapper(conn)
            return wrapped
        return conn
    return conn


def _create_schema():
    conn = _connect()
    if _is_postgres():
        _create_schema_postgres(conn)
    else:
        _create_schema_sqlite(conn)


def _create_schema_sqlite(conn):
    """SQLite schema creation (original logic)."""
    with conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS email_runs (
                id                            TEXT PRIMARY KEY,
                message_id                    TEXT UNIQUE,
                subject                       TEXT DEFAULT '',
                received_at                   TEXT DEFAULT '',
                spv_name                      TEXT DEFAULT '',
                status                        TEXT DEFAULT 'pending',
                body_html                     TEXT DEFAULT '',
                metadata_json                 TEXT DEFAULT '',
                spv_dsm_json                  TEXT DEFAULT '',
                vendor_data_json              TEXT DEFAULT '',
                positive_fields_json          TEXT DEFAULT '',
                negative_fields_json          TEXT DEFAULT '',
                downloaded_pdf_names          TEXT DEFAULT '',
                nis_checklist_id              TEXT DEFAULT '',
                nis_checklist_pdf_uploaded    INTEGER DEFAULT 0,
                checklist_value               TEXT DEFAULT '',
                sap_checklist_pdf_uploaded    INTEGER DEFAULT 0,
                booking_folder_path           TEXT DEFAULT '',
                failure_email_sent            INTEGER DEFAULT 0,
                nis_completion_email_sent_at  TEXT DEFAULT '',
                created_at                    TEXT DEFAULT '',
                started_at                    TEXT DEFAULT '',
                finished_at                   TEXT DEFAULT '',
                error_message                 TEXT DEFAULT ''
            );
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS booking_rows (
                row_id                        INTEGER PRIMARY KEY AUTOINCREMENT,
                run_id                        TEXT,
                message_id                    TEXT,
                row_index                     INTEGER,
                spv_name                      TEXT DEFAULT '',
                status                        TEXT DEFAULT 'pending',
                spv_data_json                 TEXT DEFAULT '',
                metadata_json                 TEXT DEFAULT '',
                positive_fields_json          TEXT DEFAULT '',
                negative_fields_json          TEXT DEFAULT '',
                vendor_data_json              TEXT DEFAULT '',
                invoice_date                  TEXT DEFAULT '',
                invoice_number                TEXT DEFAULT '',
                nis_checklist_id              TEXT DEFAULT '',
                nis_checklist_pdf_uploaded    INTEGER DEFAULT 0,
                checklist_value               TEXT DEFAULT '',
                sap_checklist_pdf_uploaded    INTEGER DEFAULT 0,
                downloaded_pdf_names          TEXT DEFAULT '',
                booking_folder_path           TEXT DEFAULT '',
                retry_count                   INTEGER DEFAULT 0,
                error_message                 TEXT DEFAULT '',
                nis_completion_email_sent_at  TEXT DEFAULT '',
                failure_email_sent            INTEGER DEFAULT 0,
                created_at                    TEXT DEFAULT '',
                started_at                    TEXT DEFAULT '',
                finished_at                   TEXT DEFAULT '',
                UNIQUE(message_id, row_index)
            );
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS run_steps (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                run_id        TEXT,
                step_name     TEXT,
                status        TEXT DEFAULT 'pending',
                detail        TEXT DEFAULT '',
                started_at    TEXT DEFAULT '',
                finished_at   TEXT DEFAULT '',
                error         TEXT DEFAULT '',
                retry_count   INTEGER DEFAULT 0,
                UNIQUE(run_id, step_name)
            );
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS debit_credit_emails (
                id                    INTEGER PRIMARY KEY AUTOINCREMENT,
                workflow_id           TEXT DEFAULT '',
                run_id                TEXT,
                message_id            TEXT UNIQUE,
                spv_name              TEXT DEFAULT '',
                week_no               TEXT DEFAULT '',
                from_date             TEXT DEFAULT '',
                to_date               TEXT DEFAULT '',
                dsm_amount            TEXT DEFAULT '',
                company_code          TEXT DEFAULT '',
                due_date_cn_dn        TEXT DEFAULT '',
                recipients            TEXT DEFAULT '',
                sent_subject          TEXT DEFAULT '',
                email_status          TEXT DEFAULT 'pending',
                sent_at               TEXT DEFAULT '',
                reply_message_id      TEXT DEFAULT '',
                reply_received_at     TEXT DEFAULT '',
                reply_pdf_names       TEXT DEFAULT '',
                reply_pdfs_uploaded   INTEGER DEFAULT 0,
                booking_folder_path   TEXT DEFAULT '',
                error_message         TEXT DEFAULT '',
                created_at            TEXT DEFAULT ''
            );
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS debit_credit_pdfs (
                id                    INTEGER PRIMARY KEY AUTOINCREMENT,
                workflow_id           TEXT DEFAULT '',
                dc_email_id           INTEGER,
                run_id                TEXT DEFAULT '',
                reply_message_id      TEXT DEFAULT '',
                pdf_name              TEXT DEFAULT '',
                spv_from_filename     TEXT DEFAULT '',
                matched_spv_name      TEXT DEFAULT '',
                vendor_code           TEXT DEFAULT '',
                company_code          TEXT DEFAULT '',
                cost_center           TEXT DEFAULT '',
                plant                 TEXT DEFAULT '',
                ref_number            TEXT DEFAULT '',
                amount                REAL DEFAULT 0,
                amount_text           TEXT DEFAULT '',
                doc_date              TEXT DEFAULT '',
                note_type             TEXT DEFAULT '',
                dsm_week              TEXT DEFAULT '',
                extract_status        TEXT DEFAULT 'pending',
                sap_status            TEXT DEFAULT 'pending',
                sap_checklist_number  TEXT DEFAULT '',
                sap_pdf_name          TEXT DEFAULT '',
                sap_pdf_sp_path       TEXT DEFAULT '',
                sap_folder_path       TEXT DEFAULT '',
                from_date             TEXT DEFAULT '',
                to_date               TEXT DEFAULT '',
                error_message         TEXT DEFAULT '',
                created_at            TEXT DEFAULT '',
                updated_at            TEXT DEFAULT '',
                UNIQUE(dc_email_id, pdf_name)
            );
            """
        )
        conn.execute("CREATE INDEX IF NOT EXISTS idx_steps_run ON run_steps(run_id);")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_rows_run ON booking_rows(run_id);")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_dc_msg ON debit_credit_emails(message_id);")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_dc_wf ON debit_credit_emails(workflow_id);")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_dcpdf_email ON debit_credit_pdfs(dc_email_id);")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_dcpdf_wf ON debit_credit_pdfs(workflow_id);")
        _migrate_add_columns(conn)


def _migrate_add_columns(conn):
    """Add columns introduced after a DB may already exist (SQLite only).
    Each ALTER is idempotent-guarded."""
    migrations = {
        "booking_rows": [
            ("invoice_number", "TEXT DEFAULT ''"),
            ("hod_approval_status", "TEXT DEFAULT 'pending'"),
            ("hod_approval_at", "TEXT DEFAULT ''"),
            ("hod_approval_message_id", "TEXT DEFAULT ''"),
            ("hod_email_sent", "INTEGER DEFAULT 0"),
            ("hod_email_sent_at", "TEXT DEFAULT ''"),
            ("fc_approval_status", "TEXT DEFAULT 'pending'"),
            ("fc_approval_at", "TEXT DEFAULT ''"),
            ("fc_approval_message_id", "TEXT DEFAULT ''"),
            ("fc_checklist_pdf_uploaded", "INTEGER DEFAULT 0"),
            ("fc_email_sent", "INTEGER DEFAULT 0"),
            ("fc_email_sent_at", "TEXT DEFAULT ''"),
            ("pdf_merge_status", "TEXT DEFAULT 'pending'"),
            ("pdf_merge_at", "TEXT DEFAULT ''"),
            ("merged_pdf_name", "TEXT DEFAULT ''"),
            ("completion_email_sent", "INTEGER DEFAULT 0"),
            ("completion_email_sent_at", "TEXT DEFAULT ''"),
        ],
    }
    for table, cols in migrations.items():
        try:
            existing = {r[1] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()}
        except Exception:
            continue
        for col_name, col_def in cols:
            if col_name not in existing:
                try:
                    conn.execute(f"ALTER TABLE {table} ADD COLUMN {col_name} {col_def}")
                except Exception:
                    pass


def _create_schema_postgres(conn):
    """PostgreSQL schema creation. Uses SERIAL for auto-increment columns.
    Tables are created if they don't exist; columns are added if missing."""
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS email_runs (
            id                            TEXT PRIMARY KEY,
            message_id                    TEXT UNIQUE,
            subject                       TEXT DEFAULT '',
            received_at                   TEXT DEFAULT '',
            spv_name                      TEXT DEFAULT '',
            status                        TEXT DEFAULT 'pending',
            body_html                     TEXT DEFAULT '',
            metadata_json                 TEXT DEFAULT '',
            spv_dsm_json                  TEXT DEFAULT '',
            vendor_data_json              TEXT DEFAULT '',
            positive_fields_json          TEXT DEFAULT '',
            negative_fields_json          TEXT DEFAULT '',
            downloaded_pdf_names          TEXT DEFAULT '',
            nis_checklist_id              TEXT DEFAULT '',
            nis_checklist_pdf_uploaded    INTEGER DEFAULT 0,
            checklist_value               TEXT DEFAULT '',
            sap_checklist_pdf_uploaded    INTEGER DEFAULT 0,
            booking_folder_path           TEXT DEFAULT '',
            failure_email_sent            INTEGER DEFAULT 0,
            nis_completion_email_sent_at  TEXT DEFAULT '',
            created_at                    TEXT DEFAULT '',
            started_at                    TEXT DEFAULT '',
            finished_at                   TEXT DEFAULT '',
            error_message                 TEXT DEFAULT ''
        );
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS booking_rows (
            row_id                        SERIAL PRIMARY KEY,
            run_id                        TEXT,
            message_id                    TEXT,
            row_index                     INTEGER,
            spv_name                      TEXT DEFAULT '',
            status                        TEXT DEFAULT 'pending',
            spv_data_json                 TEXT DEFAULT '',
            metadata_json                 TEXT DEFAULT '',
            positive_fields_json          TEXT DEFAULT '',
            negative_fields_json          TEXT DEFAULT '',
            vendor_data_json              TEXT DEFAULT '',
            invoice_date                  TEXT DEFAULT '',
            invoice_number                TEXT DEFAULT '',
            nis_checklist_id              TEXT DEFAULT '',
            nis_checklist_pdf_uploaded    INTEGER DEFAULT 0,
            checklist_value               TEXT DEFAULT '',
            sap_checklist_pdf_uploaded    INTEGER DEFAULT 0,
            downloaded_pdf_names          TEXT DEFAULT '',
            booking_folder_path           TEXT DEFAULT '',
            retry_count                   INTEGER DEFAULT 0,
            error_message                 TEXT DEFAULT '',
            nis_completion_email_sent_at  TEXT DEFAULT '',
            failure_email_sent            INTEGER DEFAULT 0,
            hod_approval_status           TEXT DEFAULT 'pending',
            hod_approval_at               TEXT DEFAULT '',
            hod_approval_message_id       TEXT DEFAULT '',
            hod_email_sent                INTEGER DEFAULT 0,
            hod_email_sent_at             TEXT DEFAULT '',
            fc_approval_status            TEXT DEFAULT 'pending',
            fc_approval_at                TEXT DEFAULT '',
            fc_approval_message_id        TEXT DEFAULT '',
            fc_checklist_pdf_uploaded     INTEGER DEFAULT 0,
            fc_email_sent                 INTEGER DEFAULT 0,
            fc_email_sent_at              TEXT DEFAULT '',
            pdf_merge_status              TEXT DEFAULT 'pending',
            pdf_merge_at                  TEXT DEFAULT '',
            merged_pdf_name               TEXT DEFAULT '',
            completion_email_sent         INTEGER DEFAULT 0,
            completion_email_sent_at      TEXT DEFAULT '',
            created_at                    TEXT DEFAULT '',
            started_at                    TEXT DEFAULT '',
            finished_at                   TEXT DEFAULT '',
            UNIQUE(message_id, row_index)
        );
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS run_steps (
            id            SERIAL PRIMARY KEY,
            run_id        TEXT,
            step_name     TEXT,
            status        TEXT DEFAULT 'pending',
            detail        TEXT DEFAULT '',
            started_at    TEXT DEFAULT '',
            finished_at   TEXT DEFAULT '',
            error         TEXT DEFAULT '',
            retry_count   INTEGER DEFAULT 0,
            UNIQUE(run_id, step_name)
        );
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS debit_credit_emails (
            id                    SERIAL PRIMARY KEY,
            workflow_id           TEXT DEFAULT '',
            run_id                TEXT,
            message_id            TEXT UNIQUE,
            spv_name              TEXT DEFAULT '',
            week_no               TEXT DEFAULT '',
            from_date             TEXT DEFAULT '',
            to_date               TEXT DEFAULT '',
            dsm_amount            TEXT DEFAULT '',
            company_code          TEXT DEFAULT '',
            due_date_cn_dn        TEXT DEFAULT '',
            recipients            TEXT DEFAULT '',
            sent_subject          TEXT DEFAULT '',
            email_status          TEXT DEFAULT 'pending',
            sent_at               TEXT DEFAULT '',
            reply_message_id      TEXT DEFAULT '',
            reply_received_at     TEXT DEFAULT '',
            reply_pdf_names       TEXT DEFAULT '',
            reply_pdfs_uploaded   INTEGER DEFAULT 0,
            booking_folder_path   TEXT DEFAULT '',
            error_message         TEXT DEFAULT '',
            created_at            TEXT DEFAULT ''
        );
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS debit_credit_pdfs (
            id                    SERIAL PRIMARY KEY,
            workflow_id           TEXT DEFAULT '',
            dc_email_id           INTEGER,
            run_id                TEXT DEFAULT '',
            reply_message_id      TEXT DEFAULT '',
            pdf_name              TEXT DEFAULT '',
            spv_from_filename     TEXT DEFAULT '',
            matched_spv_name      TEXT DEFAULT '',
            vendor_code           TEXT DEFAULT '',
            company_code          TEXT DEFAULT '',
            cost_center           TEXT DEFAULT '',
            plant                 TEXT DEFAULT '',
            ref_number            TEXT DEFAULT '',
            amount                REAL DEFAULT 0,
            amount_text           TEXT DEFAULT '',
            doc_date              TEXT DEFAULT '',
            note_type             TEXT DEFAULT '',
            dsm_week              TEXT DEFAULT '',
            extract_status        TEXT DEFAULT 'pending',
            sap_status            TEXT DEFAULT 'pending',
            sap_checklist_number  TEXT DEFAULT '',
            sap_pdf_name          TEXT DEFAULT '',
            sap_pdf_sp_path       TEXT DEFAULT '',
            sap_folder_path       TEXT DEFAULT '',
            from_date             TEXT DEFAULT '',
            to_date               TEXT DEFAULT '',
            error_message         TEXT DEFAULT '',
            created_at            TEXT DEFAULT '',
            updated_at            TEXT DEFAULT '',
            UNIQUE(dc_email_id, pdf_name)
        );
        """
    )
    cur.execute("CREATE INDEX IF NOT EXISTS idx_steps_run ON run_steps(run_id);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_rows_run ON booking_rows(run_id);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_dc_msg ON debit_credit_emails(message_id);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_dc_wf ON debit_credit_emails(workflow_id);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_dcpdf_email ON debit_credit_pdfs(dc_email_id);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_dcpdf_wf ON debit_credit_pdfs(workflow_id);")
    _migrate_add_columns_postgres(conn)
    conn.commit()


def _migrate_add_columns_postgres(conn):
    """Add columns that may be missing on an existing PostgreSQL database.
    Each ALTER TABLE ADD COLUMN is wrapped in a savepoint so a 'duplicate
    column' error doesn't abort the entire transaction."""
    migrations = {
        "booking_rows": [
            ("invoice_number", "TEXT DEFAULT ''"),
            ("hod_approval_status", "TEXT DEFAULT 'pending'"),
            ("hod_approval_at", "TEXT DEFAULT ''"),
            ("hod_approval_message_id", "TEXT DEFAULT ''"),
            ("hod_email_sent", "INTEGER DEFAULT 0"),
            ("hod_email_sent_at", "TEXT DEFAULT ''"),
            ("fc_approval_status", "TEXT DEFAULT 'pending'"),
            ("fc_approval_at", "TEXT DEFAULT ''"),
            ("fc_approval_message_id", "TEXT DEFAULT ''"),
            ("fc_checklist_pdf_uploaded", "INTEGER DEFAULT 0"),
            ("fc_email_sent", "INTEGER DEFAULT 0"),
            ("fc_email_sent_at", "TEXT DEFAULT ''"),
            ("pdf_merge_status", "TEXT DEFAULT 'pending'"),
            ("pdf_merge_at", "TEXT DEFAULT ''"),
            ("merged_pdf_name", "TEXT DEFAULT ''"),
            ("completion_email_sent", "INTEGER DEFAULT 0"),
            ("completion_email_sent_at", "TEXT DEFAULT ''"),
        ],
    }
    cur = conn.cursor()
    for table, cols in migrations.items():
        for col_name, col_def in cols:
            try:
                cur.execute(f"SAVEPOINT sp_migrate;")
                cur.execute(f"ALTER TABLE {table} ADD COLUMN {col_name} {col_def};")
                cur.execute(f"RELEASE SAVEPOINT sp_migrate;")
            except Exception:
                cur.execute(f"ROLLBACK TO SAVEPOINT sp_migrate;")


def _row_to_dict(row) -> Optional[dict]:
    """Convert a row (sqlite3.Row or psycopg2 RealDictRow or None) to dict."""
    return dict(row) if row is not None else None


def _rows_to_dicts(rows) -> list:
    return [dict(r) for r in rows]


# ─────────────────────────────────────────────────────────────────────────────
# Init
# ─────────────────────────────────────────────────────────────────────────────
def init_db():
    with _LOCK:
        _create_schema()
    if _is_postgres():
        target = f"{_PG_HOST}:{_PG_PORT}/{_PG_DB}" if not _PG_DSN else "(DSN)"
        print(f"[DB] Initialised → PostgreSQL @ {target}")
    else:
        print(f"[DB] Initialised → SQLite @ {_DB_PATH}")


# ─────────────────────────────────────────────────────────────────────────────
# Email-run registration / lookup
# ─────────────────────────────────────────────────────────────────────────────
def is_email_registered(message_id: str) -> bool:
    run = get_run_by_message_id(message_id)
    if run is None:
        return False
    # Extraction failures are retryable. Keeping the row preserves diagnostics,
    # while returning False lets the next mailbox scan parse the email again.
    return str(run.get("status") or "") != "extraction_failed"


def register_email(message_id: str, subject: str, received_at: str, body_html: str = "") -> str:
    now = _now()
    run_id = str(message_id or "")
    with _LOCK:
        conn = _get_conn()
        existing = conn.execute(
            "SELECT id, status FROM email_runs WHERE message_id = ?", (str(message_id),)
        ).fetchone()
        if existing:
            existing_id = str(existing["id"])
            if str(existing["status"] or "") == "extraction_failed":
                with conn:
                    conn.execute(
                        """
                        UPDATE email_runs
                        SET subject = ?, received_at = ?, body_html = ?,
                            status = ?, error_message = '', finished_at = ''
                        WHERE id = ?
                        """,
                        (
                            subject,
                            received_at,
                            body_html,
                            STATUS_PENDING,
                            existing_id,
                        ),
                    )
                print(f"[DB] Retrying extraction  run_id={existing_id[:40]}")
            return existing_id

        with conn:
            conn.execute(
                """
                INSERT INTO email_runs (id, message_id, subject, received_at,
                    status, body_html, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (run_id, str(message_id), subject, received_at,
                 STATUS_PENDING, body_html, now),
            )
            for step in _FIXED_STEPS:
                conn.execute(
                    """
                    INSERT OR IGNORE INTO run_steps (run_id, step_name, status)
                    VALUES (?, ?, ?)
                    """,
                    (run_id, step, STATUS_PENDING),
                )
    print(f"[DB] Registered  run_id={run_id[:40]}  msg={str(message_id)[:40]}…")
    return run_id


def get_run_by_message_id(message_id: str) -> Optional[dict]:
    with _LOCK:
        conn = _get_conn()
        row = conn.execute(
            "SELECT * FROM email_runs WHERE message_id = ?", (str(message_id),)
        ).fetchone()
    return _row_to_dict(row)


def get_run_by_id(run_id: str) -> Optional[dict]:
    with _LOCK:
        conn = _get_conn()
        row = conn.execute(
            "SELECT * FROM email_runs WHERE id = ? OR message_id = ?",
            (str(run_id), str(run_id)),
        ).fetchone()
    return _row_to_dict(row)


def get_pending_runs() -> list:
    with _LOCK:
        conn = _get_conn()
        rows = conn.execute(
            "SELECT * FROM email_runs WHERE status IN (?, ?) ORDER BY created_at",
            (STATUS_PENDING, STATUS_RUNNING),
        ).fetchall()
    return _rows_to_dicts(rows)


def get_incomplete_runs() -> list:
    with _LOCK:
        conn = _get_conn()
        rows = conn.execute(
            "SELECT * FROM email_runs WHERE status IN (?, ?, ?) ORDER BY created_at",
            (STATUS_PENDING, STATUS_RUNNING, STATUS_FAILED),
        ).fetchall()
    return _rows_to_dicts(rows)


def _update_run(run_id: str, fields: dict):
    if not fields:
        return
    cols = [c for c in fields if c in _EMAIL_RUNS_COLUMNS]
    if not cols:
        return
    assignments = ", ".join(f"{c} = ?" for c in cols)
    params = [fields[c] for c in cols] + [str(run_id), str(run_id)]
    with _LOCK:
        conn = _get_conn()
        with conn:
            cur = conn.execute(
                f"UPDATE email_runs SET {assignments} WHERE id = ? OR message_id = ?",
                params,
            )
        if cur.rowcount == 0:
            raise ValueError(f"run_id={run_id} not found")


# ─────────────────────────────────────────────────────────────────────────────
# Email-run data setters
# ─────────────────────────────────────────────────────────────────────────────
def save_extracted_data(run_id: str, metadata_row: dict, spv_dsm_rows: list):
    _update_run(run_id, {
        "metadata_json": json.dumps(metadata_row or {}),
        "spv_dsm_json": json.dumps(spv_dsm_rows or []),
    })


def save_classification(run_id: str, positive: dict, negative: dict):
    _update_run(run_id, {
        "positive_fields_json": json.dumps(positive or {}),
        "negative_fields_json": json.dumps(negative or {}),
    })


def save_vendor_data(run_id: str, vendor_data: dict):
    _update_run(run_id, {"vendor_data_json": json.dumps(vendor_data or {})})


def save_downloaded_pdfs(run_id: str, pdf_paths: list):
    save_downloaded_pdf_name(run_id, "dsm_pdf", pdf_paths)


def _parse_downloaded_pdf_names(raw) -> dict:
    if isinstance(raw, dict):
        return dict(raw)
    try:
        parsed = json.loads(raw or "{}")
    except Exception:
        return {}
    if isinstance(parsed, dict):
        return parsed
    if isinstance(parsed, list):
        return {"dsm_pdf": parsed}
    return {}


def save_downloaded_pdf_name(run_id: str, pdf_type: str, pdf_names):
    names = [
        str(p).replace("\\", "/").split("/")[-1]
        for p in (pdf_names if isinstance(pdf_names, (list, tuple, set)) else [pdf_names])
        if str(p or "").strip()
    ]
    with _LOCK:
        row = get_run_by_id(run_id) or {}
        name_map = _parse_downloaded_pdf_names(row.get("downloaded_pdf_names"))
        name_map[str(pdf_type)] = names if len(names) != 1 else names[0]
        _update_run(run_id, {"downloaded_pdf_names": json.dumps(name_map)})


def _parse_id_collection(raw) -> list:
    raw_text = str(raw or "").strip()
    if not raw_text:
        return []
    try:
        parsed = json.loads(raw_text)
    except Exception:
        parsed = None
    if isinstance(parsed, list):
        return [str(v).strip() for v in parsed if str(v or "").strip()]
    if isinstance(parsed, dict):
        return [str(v).strip() for v in parsed.values() if str(v or "").strip()]
    return [part.strip() for part in raw_text.split(",") if part.strip()]


def save_nis_checklist_id(run_id: str, checklist_id: str, append: bool = False):
    checklist_id = str(checklist_id or "").strip()
    if not append:
        _update_run(run_id, {"nis_checklist_id": checklist_id})
        return
    with _LOCK:
        row = get_run_by_id(run_id) or {}
        ids = _parse_id_collection(row.get("nis_checklist_id"))
        if checklist_id and checklist_id not in ids:
            ids.append(checklist_id)
        _update_run(run_id, {"nis_checklist_id": ", ".join(ids)})


def _merge_id_collection(existing, new_value) -> str:
    """Merge *new_value* (one number or a comma/JSON list) into *existing*,
    preserving every pre-existing entry and de-duplicating. Returns a
    comma-joined string. Used to guarantee NIS/SAP numbers are never lost.
    """
    ids = _parse_id_collection(existing)
    for candidate in _parse_id_collection(new_value):
        if candidate and candidate not in ids:
            ids.append(candidate)
    return ", ".join(ids)


def save_checklist_value(run_id: str, checklist_value: str, append: bool = True):
    """Persist a SAP checklist number on the run.

    Preservation-safe by default: appends the new number to any existing
    value(s) and de-duplicates, so an already-stored SAP checklist number is
    never overwritten or lost. Pass ``append=False`` only to force-set.
    """
    new_value = str(checklist_value or "").strip()
    if not append:
        _update_run(run_id, {"checklist_value": new_value})
        return
    with _LOCK:
        row = get_run_by_id(run_id) or {}
        merged = _merge_id_collection(row.get("checklist_value"), new_value)
        _update_run(run_id, {"checklist_value": merged})


def save_booking_folder_path(run_id: str, folder_path: str):
    _update_run(run_id, {"booking_folder_path": str(folder_path or "")})


def get_run_by_nis_checklist_id(checklist_id: str) -> Optional[dict]:
    checklist_id = str(checklist_id or "").strip()
    with _LOCK:
        conn = _get_conn()
        rows = conn.execute("SELECT * FROM email_runs").fetchall()
    for row in _rows_to_dicts(rows):
        if checklist_id in _parse_id_collection(row.get("nis_checklist_id")):
            return row
    return None


def get_run_by_sap_checklist_value(checklist_value: str) -> Optional[dict]:
    with _LOCK:
        conn = _get_conn()
        row = conn.execute(
            "SELECT * FROM email_runs WHERE checklist_value = ?",
            (str(checklist_value),),
        ).fetchone()
    return _row_to_dict(row)


def mark_nis_checklist_pdf_uploaded(run_id: str, uploaded: bool = True):
    _update_run(run_id, {"nis_checklist_pdf_uploaded": int(bool(uploaded))})


def mark_sap_checklist_pdf_uploaded(run_id: str, uploaded: bool = True):
    _update_run(run_id, {"sap_checklist_pdf_uploaded": int(bool(uploaded))})


def save_spv_name(run_id: str, spv_name: str):
    _update_run(run_id, {"spv_name": spv_name})


def is_failure_email_sent(run_id: str) -> bool:
    row = get_run_by_id(run_id)
    return bool(row and int(row.get("failure_email_sent") or 0))


def mark_failure_email_sent(run_id: str):
    _update_run(run_id, {"failure_email_sent": 1})


def is_nis_completion_email_sent(run_id: str) -> bool:
    row = get_run_by_id(run_id)
    return bool(row and str(row.get("nis_completion_email_sent_at") or "").strip())


def mark_nis_completion_email_sent(run_id: str, sent_at: str = ""):
    _update_run(run_id, {"nis_completion_email_sent_at": sent_at or _now()})


def get_runs_pending_nis_completion_email() -> list:
    with _LOCK:
        conn = _get_conn()
        rows = conn.execute(
            "SELECT * FROM email_runs "
            "WHERE nis_completion_email_sent_at IS NULL "
            "OR TRIM(nis_completion_email_sent_at) = '' "
            "ORDER BY created_at"
        ).fetchall()
    return _rows_to_dicts(rows)


def mark_run_started(run_id: str):
    _update_run(run_id, {"status": STATUS_RUNNING, "started_at": _now()})


def mark_run_done(run_id: str):
    _update_run(run_id, {"status": STATUS_DONE, "finished_at": _now()})


def mark_run_failed(run_id: str, error: str):
    _update_run(run_id, {
        "status": STATUS_FAILED,
        "finished_at": _now(),
        "error_message": str(error)[:2000],
    })


# ─────────────────────────────────────────────────────────────────────────────
# Completion-email recipients (stored on SharePoint, NOT in SQLite)
# ─────────────────────────────────────────────────────────────────────────────
def get_nis_completion_recipients() -> list:
    """Recipients live on SharePoint (see recipients_sp.py), not in SQLite.

    Falls back to ``NIS_COMPLETION_RECIPIENTS`` / ``FAILURE_EMAIL_TO`` env var if the SharePoint module
    cannot be imported/read so notifications never silently break.
    """
    try:
        import recipients_sp
        return recipients_sp.get_recipients()
    except Exception as exc:
        print(f"[DB] Recipients unavailable from SharePoint ({exc}); using env fallback")
        seen, out = set(), []
        raw = os.getenv("NIS_COMPLETION_RECIPIENTS") or os.getenv("FAILURE_EMAIL_TO") or ""
        for addr in raw.replace(";", ",").split(","):
            email = addr.strip()
            if email and email.lower() not in seen:
                seen.add(email.lower())
                out.append(email)
        return out


# ─────────────────────────────────────────────────────────────────────────────
# Step tracking
# ─────────────────────────────────────────────────────────────────────────────
def ensure_dynamic_step(run_id: str, step_name: str):
    with _LOCK:
        conn = _get_conn()
        with conn:
            conn.execute(
                "INSERT OR IGNORE INTO run_steps (run_id, step_name, status) "
                "VALUES (?, ?, ?)",
                (str(run_id), str(step_name), STATUS_PENDING),
            )


def get_steps(run_id: str) -> list:
    with _LOCK:
        conn = _get_conn()
        rows = conn.execute(
            "SELECT * FROM run_steps WHERE run_id = ? ORDER BY id",
            (str(run_id),),
        ).fetchall()
    return _rows_to_dicts(rows)


def get_step(run_id: str, step_name: str) -> Optional[dict]:
    with _LOCK:
        conn = _get_conn()
        row = conn.execute(
            "SELECT * FROM run_steps WHERE run_id = ? AND step_name = ?",
            (str(run_id), str(step_name)),
        ).fetchone()
    return _row_to_dict(row)


def _update_step(run_id: str, step_name: str, fields: dict):
    with _LOCK:
        conn = _get_conn()
        with conn:
            conn.execute(
                "INSERT OR IGNORE INTO run_steps (run_id, step_name, status) "
                "VALUES (?, ?, ?)",
                (str(run_id), str(step_name), STATUS_PENDING),
            )
            cols = [c for c in fields if c in _RUN_STEPS_COLUMNS]
            if not cols:
                return
            assignments = ", ".join(f"{c} = ?" for c in cols)
            params = [fields[c] for c in cols] + [str(run_id), str(step_name)]
            conn.execute(
                f"UPDATE run_steps SET {assignments} "
                f"WHERE run_id = ? AND step_name = ?",
                params,
            )


def is_step_done(run_id: str, step_name: str) -> bool:
    s = get_step(run_id, step_name)
    return s is not None and s.get("status") == STATUS_DONE


def is_step_skipped(run_id: str, step_name: str) -> bool:
    s = get_step(run_id, step_name)
    return s is not None and s.get("status") == STATUS_SKIPPED


def are_all_sap_steps_complete(run_id: str) -> bool:
    rows = [r for r in get_steps(run_id)
            if str(r.get("step_name") or "").startswith("sap_automation_")]
    if not rows:
        return True
    return all(r.get("status") in (STATUS_DONE, STATUS_SKIPPED) for r in rows)


def get_all_pending_sap_steps() -> list:
    """Every SAP automation step across ALL runs that is not done/skipped.

    Returns ``run_steps`` rows whose ``step_name`` starts with
    ``sap_automation_`` and whose status is pending/running/failed. This is the
    typed SAP queue consumed by the independent sap_automation service (the NIS
    service never touches these). Ordered oldest-run first for stable draining.
    """
    with _LOCK:
        conn = _get_conn()
        rows = conn.execute(
            "SELECT * FROM run_steps "
            "WHERE step_name LIKE 'sap_automation_%' "
            "AND status NOT IN (?, ?) "
            "ORDER BY run_id, id",
            (STATUS_DONE, STATUS_SKIPPED),
        ).fetchall()
    return _rows_to_dicts(rows)


def reset_crashed_sap_steps(run_id: str, max_retries: int) -> int:
    reset_count = 0
    with _LOCK:
        conn = _get_conn()
        rows = conn.execute(
            "SELECT id, status, retry_count, step_name FROM run_steps "
            "WHERE run_id = ? AND step_name LIKE 'sap_automation_%'",
            (str(run_id),),
        ).fetchall()
        with conn:
            for r in rows:
                if r["status"] in (STATUS_RUNNING, STATUS_FAILED) and \
                        int(r["retry_count"] or 0) < max_retries:
                    conn.execute(
                        "UPDATE run_steps SET status = ?, error = '', finished_at = '' "
                        "WHERE id = ?",
                        (STATUS_PENDING, r["id"]),
                    )
                    reset_count += 1
    return reset_count


def step_start(run_id: str, step_name: str, detail: str = ""):
    _update_step(run_id, step_name, {
        "status": STATUS_RUNNING, "started_at": _now(),
        "detail": detail, "error": "", "finished_at": "",
    })


def step_done(run_id: str, step_name: str, detail: str = ""):
    _update_step(run_id, step_name, {
        "status": STATUS_DONE, "finished_at": _now(), "detail": detail,
    })


def step_failed(run_id: str, step_name: str, error: str):
    _update_step(run_id, step_name, {
        "status": STATUS_FAILED, "finished_at": _now(), "error": str(error)[:2000],
    })


def step_skip(run_id: str, step_name: str, reason: str = ""):
    _update_step(run_id, step_name, {
        "status": STATUS_SKIPPED, "finished_at": _now(), "detail": reason,
    })


def get_step_retry_count(run_id: str, step_name: str) -> int:
    row = get_step(run_id, step_name)
    return int(row.get("retry_count") or 0) if row else 0


def increment_step_retry(run_id, step_name: str, error: str) -> int:
    row = get_step(run_id, step_name)
    current = int(row.get("retry_count") or 0) if row else 0
    new_count = current + 1
    _update_step(run_id, step_name, {
        "retry_count": new_count, "status": STATUS_PENDING,
        "error": str(error)[:2000], "finished_at": _now(),
    })
    return new_count


# ─────────────────────────────────────────────────────────────────────────────
# Per-row booking units (independent SPV bookings sharing one message_id)
# ─────────────────────────────────────────────────────────────────────────────
def seed_booking_rows(run_id: str, message_id: str, spv_rows: list,
                      metadata_row: dict) -> list:
    """Create one independent booking row per SPV row in an email.

    Idempotent: re-running for the same email updates the snapshot data of
    existing rows (by ``message_id`` + ``row_index``) without resetting their
    status, and inserts any new rows. Returns all booking rows for the email.
    """
    now = _now()
    meta_json = json.dumps(metadata_row or {})
    with _LOCK:
        conn = _get_conn()
        with conn:
            for row_index, spv_row in enumerate(spv_rows or [], start=1):
                spv_name = str((spv_row or {}).get("spv_name") or "").strip()
                spv_json = json.dumps(spv_row or {})
                existing = conn.execute(
                    "SELECT row_id FROM booking_rows "
                    "WHERE message_id = ? AND row_index = ?",
                    (str(message_id), row_index),
                ).fetchone()
                if existing:
                    conn.execute(
                        "UPDATE booking_rows SET spv_name = ?, spv_data_json = ?, "
                        "metadata_json = ? WHERE row_id = ?",
                        (spv_name, spv_json, meta_json, existing["row_id"]),
                    )
                else:
                    conn.execute(
                        """
                        INSERT INTO booking_rows
                            (run_id, message_id, row_index, spv_name, status,
                             spv_data_json, metadata_json, created_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (str(run_id), str(message_id), row_index, spv_name,
                         STATUS_PENDING, spv_json, meta_json, now),
                    )
    return get_booking_rows(run_id)


def get_booking_rows(run_id: str) -> list:
    with _LOCK:
        conn = _get_conn()
        rows = conn.execute(
            "SELECT * FROM booking_rows WHERE run_id = ? OR message_id = ? "
            "ORDER BY row_index",
            (str(run_id), str(run_id)),
        ).fetchall()
    return _rows_to_dicts(rows)


def get_booking_row(row_id) -> Optional[dict]:
    with _LOCK:
        conn = _get_conn()
        row = conn.execute(
            "SELECT * FROM booking_rows WHERE row_id = ?", (int(row_id),)
        ).fetchone()
    return _row_to_dict(row)


def get_booking_row_by_index(message_id: str, row_index: int) -> Optional[dict]:
    with _LOCK:
        conn = _get_conn()
        row = conn.execute(
            "SELECT * FROM booking_rows WHERE message_id = ? AND row_index = ?",
            (str(message_id), int(row_index)),
        ).fetchone()
    return _row_to_dict(row)


def get_next_pending_booking_row(run_id: str) -> Optional[dict]:
    """Return the next booking row that still needs work (one at a time)."""
    with _LOCK:
        conn = _get_conn()
        row = conn.execute(
            "SELECT * FROM booking_rows "
            "WHERE (run_id = ? OR message_id = ?) "
            "AND status IN (?, ?, ?) ORDER BY row_index LIMIT 1",
            (str(run_id), str(run_id), STATUS_PENDING, STATUS_RUNNING, STATUS_FAILED),
        ).fetchone()
    return _row_to_dict(row)


def get_pending_booking_rows(run_id: str) -> list:
    with _LOCK:
        conn = _get_conn()
        rows = conn.execute(
            "SELECT * FROM booking_rows "
            "WHERE (run_id = ? OR message_id = ?) "
            "AND status IN (?, ?, ?) ORDER BY row_index",
            (str(run_id), str(run_id), STATUS_PENDING, STATUS_RUNNING, STATUS_FAILED),
        ).fetchall()
    return _rows_to_dicts(rows)


def get_all_pending_booking_rows(max_retries: Optional[int] = None) -> list:
    """Every incomplete booking row across ALL email runs.

    A row is incomplete when its status is pending/running/failed. Rows are
    ordered oldest-email-first then by row_index, so older work drains first.
    This is the global pool the parallel booking driver pulls from — it lets a
    free worker slot be filled with a row from a *different* email when one
    email doesn't have enough rows on its own.

    When *max_retries* is given, rows whose ``retry_count`` has reached that
    limit are excluded so an unbookable row can't be retried forever.
    """
    with _LOCK:
        conn = _get_conn()
        rows = conn.execute(
            "SELECT br.* FROM booking_rows br "
            "LEFT JOIN email_runs er ON er.id = br.run_id "
            "WHERE br.status IN (?, ?, ?) "
            "ORDER BY er.created_at, br.row_index",
            (STATUS_PENDING, STATUS_RUNNING, STATUS_FAILED),
        ).fetchall()
    result = _rows_to_dicts(rows)
    if max_retries is not None:
        result = [r for r in result if int(r.get("retry_count") or 0) < int(max_retries)]
    return result


def get_run_ids_with_pending_booking_rows() -> list:
    """Return run_ids that still have at least one incomplete booking row.

    A booking row is "incomplete" when its status is pending/running/failed
    (i.e. not done and not skipped). These runs must be finished before any
    new email is processed so no booking row is left half-done.
    Ordered by the earliest created run first so older work drains first.
    """
    with _LOCK:
        conn = _get_conn()
        rows = conn.execute(
            "SELECT br.run_id AS run_id "
            "FROM booking_rows br "
            "LEFT JOIN email_runs er ON er.id = br.run_id "
            "WHERE br.status IN (?, ?, ?) AND br.run_id IS NOT NULL "
            "GROUP BY br.run_id "
            "ORDER BY MIN(er.created_at)",
            (STATUS_PENDING, STATUS_RUNNING, STATUS_FAILED),
        ).fetchall()
    return [str(r["run_id"]) for r in rows if r["run_id"]]


def update_booking_row(row_id, fields: dict):
    if not fields:
        return
    cols = [c for c in fields if c in _BOOKING_ROWS_COLUMNS and c != "row_id"]
    if not cols:
        return
    assignments = ", ".join(f"{c} = ?" for c in cols)
    params = [fields[c] for c in cols] + [int(row_id)]
    with _LOCK:
        conn = _get_conn()
        with conn:
            conn.execute(
                f"UPDATE booking_rows SET {assignments} WHERE row_id = ?", params
            )


def mark_booking_row_started(row_id):
    update_booking_row(row_id, {"status": STATUS_RUNNING, "started_at": _now(),
                                "error_message": ""})


def mark_booking_row_done(row_id):
    update_booking_row(row_id, {"status": STATUS_DONE, "finished_at": _now()})


def mark_booking_row_nis_done(row_id):
    """Mark the row as NIS-complete but awaiting SAP automation.

    Used when a row has both positive (NIS) and negative (SAP) values.
    The SAP service will finalize it to STATUS_DONE once all SAP steps
    for this row are complete and checklist values are captured.
    """
    update_booking_row(row_id, {"status": STATUS_NIS_DONE})


def mark_booking_row_failed(row_id, error: str):
    update_booking_row(row_id, {"status": STATUS_FAILED, "finished_at": _now(),
                                "error_message": str(error)[:2000]})


def mark_booking_row_skipped(row_id, reason: str = ""):
    update_booking_row(row_id, {"status": STATUS_SKIPPED, "finished_at": _now(),
                                "error_message": str(reason)[:2000]})


def save_booking_row_classification(row_id, positive: dict, negative: dict):
    update_booking_row(row_id, {
        "positive_fields_json": json.dumps(positive or {}),
        "negative_fields_json": json.dumps(negative or {}),
    })


def save_booking_row_vendor(row_id, vendor_data: dict):
    update_booking_row(row_id, {"vendor_data_json": json.dumps(vendor_data or {})})


def save_booking_row_invoice_date(row_id, invoice_date: str):
    update_booking_row(row_id, {"invoice_date": str(invoice_date or "")})


def save_booking_row_invoice_number(row_id, invoice_number: str):
    update_booking_row(row_id, {"invoice_number": str(invoice_number or "").strip()})


def save_booking_row_nis_checklist_id(row_id, checklist_id: str, append: bool = False):
    """Persist the NIS booking number on a booking row.

    A single SPV row yields exactly ONE NIS booking, so the default overwrites
    with that number. When a row could legitimately accumulate more than one NIS
    number, pass ``append=True`` to merge + de-dup instead of overwriting (an
    existing number is never discarded)."""
    new_value = str(checklist_id or "").strip()
    if not append:
        update_booking_row(row_id, {"nis_checklist_id": new_value})
        return
    with _LOCK:
        row = get_booking_row(row_id) or {}
        merged = _merge_id_collection(row.get("nis_checklist_id"), new_value)
        update_booking_row(row_id, {"nis_checklist_id": merged})


def save_booking_row_sap_checklist_number(row_id, checklist_number: str):
    """Append a SAP checklist number to a booking row, preserving any already
    stored on it (comma-joined, de-duplicated). Never overwrites — a booking
    row may accumulate several SAP checklist numbers over time."""
    with _LOCK:
        row = get_booking_row(row_id) or {}
        merged = _merge_id_collection(row.get("checklist_value"),
                                      str(checklist_number or "").strip())
        update_booking_row(row_id, {"checklist_value": merged})


def save_booking_row_folder(row_id, folder_path: str):
    update_booking_row(row_id, {"booking_folder_path": str(folder_path or "")})


def increment_booking_row_retry(row_id, error: str) -> int:
    row = get_booking_row(row_id) or {}
    new_count = int(row.get("retry_count") or 0) + 1
    update_booking_row(row_id, {"retry_count": new_count,
                                "error_message": str(error)[:2000]})
    return new_count


def are_all_booking_rows_done(run_id: str) -> bool:
    rows = get_booking_rows(run_id)
    if not rows:
        return True
    return all(r.get("status") in (STATUS_DONE, STATUS_SKIPPED) for r in rows)


def are_all_sap_steps_done_for_row(run_id: str, row_index: int, spv_name: str) -> bool:
    """Check if ALL SAP automation steps for a specific booking row are done/skipped.

    Per-row SAP steps are named: sap_automation_<field>_row_<N>_<spv_safe>
    This function filters run_steps by the _row_<N>_ pattern and checks their status.

    If no SAP steps exist for this row AND the booking row has no negative
    fields, returns True (no SAP work needed). If the row HAS negative fields
    but SAP steps don't exist yet, returns False (steps haven't been created).
    """
    import re as _re
    steps = get_steps(run_id)
    # Match steps like: sap_automation_*_row_<row_index>_*
    pattern = f"_row_{row_index}_"
    row_sap_steps = [
        s for s in steps
        if str(s.get("step_name") or "").startswith("sap_automation_")
        and pattern in str(s.get("step_name") or "")
    ]
    if not row_sap_steps:
        # No SAP steps found — check if this row actually has negative fields.
        # If it does, steps should exist but haven't been created yet → NOT done.
        rows = get_booking_rows(run_id)
        for r in rows:
            if int(r.get("row_index") or 0) == row_index:
                neg_raw = r.get("negative_fields_json") or "{}"
                try:
                    neg = json.loads(neg_raw) if isinstance(neg_raw, str) else neg_raw
                except (json.JSONDecodeError, TypeError):
                    neg = {}
                if neg and isinstance(neg, dict) and any(
                    float(str(v).replace(",", "").strip() or 0) > 0
                    for v in neg.values()
                ):
                    # Row has negative values but no SAP steps created yet
                    return False
                break
        return True  # No negative fields → no SAP work needed
    return all(s.get("status") in (STATUS_DONE, STATUS_SKIPPED) for s in row_sap_steps)


def is_booking_row_nis_done(row_id) -> bool:
    """True if the booking row is in nis_done state (awaiting SAP completion)."""
    row = get_booking_row(row_id)
    if not row:
        return False
    return row.get("status") == STATUS_NIS_DONE


def finalize_booking_row_if_ready(run_id: str, row_index: int, spv_name: str) -> bool:
    """Called by SAP automation after completing a step. If all SAP steps for
    this row are done and the row was in nis_done state, mark it fully done.

    Returns True if the row was finalized.
    """
    # Find the booking row by run_id + row_index
    rows = get_booking_rows(run_id)
    target = None
    for r in rows:
        if int(r.get("row_index") or 0) == row_index:
            target = r
            break
    if not target:
        return False
    row_id = target.get("row_id")
    status = target.get("status")

    # If the row has no positive values (only negative), it was marked skipped
    # by NIS. Once SAP is done we should mark it done.
    if status == STATUS_NIS_DONE:
        if are_all_sap_steps_done_for_row(run_id, row_index, spv_name):
            mark_booking_row_done(row_id)
            return True
    elif status == STATUS_SKIPPED:
        # Row was skipped by NIS (no positive), but has SAP steps.
        # Check if all SAP steps done — if yes, mark it done.
        if are_all_sap_steps_done_for_row(run_id, row_index, spv_name):
            mark_booking_row_done(row_id)
            return True
    return False


# ─────────────────────────────────────────────────────────────────────────────
# HOD / FC Approval Tracking
# ─────────────────────────────────────────────────────────────────────────────
def get_booking_rows_by_nis_checklist_id(nis_id: str) -> list:
    """Find all booking rows that have this NIS checklist ID."""
    nis_id = str(nis_id or "").strip()
    if not nis_id:
        return []
    with _LOCK:
        conn = _get_conn()
        rows = conn.execute(
            "SELECT * FROM booking_rows WHERE nis_checklist_id LIKE ? ORDER BY row_id",
            (f"%{nis_id}%",),
        ).fetchall()
    return _rows_to_dicts(rows)


def mark_hod_approved(row_id, message_id: str = ""):
    """Mark a booking row as HOD approved."""
    update_booking_row(row_id, {
        "hod_approval_status": STATUS_DONE,
        "hod_approval_at": _now(),
        "hod_approval_message_id": str(message_id or ""),
    })


def is_hod_approved(row_id) -> bool:
    row = get_booking_row(row_id)
    return bool(row and row.get("hod_approval_status") == STATUS_DONE)


def mark_hod_email_sent(row_id):
    update_booking_row(row_id, {"hod_email_sent": 1, "hod_email_sent_at": _now()})


def is_hod_email_sent(row_id) -> bool:
    row = get_booking_row(row_id)
    return bool(row and int(row.get("hod_email_sent") or 0))


def mark_fc_approved(row_id, message_id: str = ""):
    """Mark a booking row as FC approved."""
    update_booking_row(row_id, {
        "fc_approval_status": STATUS_DONE,
        "fc_approval_at": _now(),
        "fc_approval_message_id": str(message_id or ""),
    })


def is_fc_approved(row_id) -> bool:
    row = get_booking_row(row_id)
    return bool(row and row.get("fc_approval_status") == STATUS_DONE)


def mark_fc_checklist_pdf_uploaded(row_id):
    update_booking_row(row_id, {"fc_checklist_pdf_uploaded": 1})


def mark_fc_email_sent(row_id):
    update_booking_row(row_id, {"fc_email_sent": 1, "fc_email_sent_at": _now()})


def is_fc_email_sent(row_id) -> bool:
    row = get_booking_row(row_id)
    return bool(row and int(row.get("fc_email_sent") or 0))


def mark_pdf_merged(row_id, merged_pdf_name: str = ""):
    update_booking_row(row_id, {
        "pdf_merge_status": STATUS_DONE,
        "pdf_merge_at": _now(),
        "merged_pdf_name": str(merged_pdf_name or ""),
    })


def is_pdf_merged(row_id) -> bool:
    row = get_booking_row(row_id)
    return bool(row and row.get("pdf_merge_status") == STATUS_DONE)


def mark_row_completion_email_sent(row_id):
    update_booking_row(row_id, {"completion_email_sent": 1, "completion_email_sent_at": _now()})


def is_row_completion_email_sent(row_id) -> bool:
    row = get_booking_row(row_id)
    return bool(row and int(row.get("completion_email_sent") or 0))


def get_rows_awaiting_hod() -> list:
    """Booking rows that are done (NIS+SAP) but not yet HOD approved."""
    with _LOCK:
        conn = _get_conn()
        rows = conn.execute(
            "SELECT * FROM booking_rows "
            "WHERE status = 'done' "
            "AND nis_checklist_id != '' "
            "AND (hod_approval_status IS NULL OR hod_approval_status = 'pending') "
            "ORDER BY row_id",
        ).fetchall()
    return _rows_to_dicts(rows)


def get_rows_pending_auto_advance() -> list:
    """Booking rows that are done but HOD/FC approval not yet granted.

    These are candidates for auto-advancing: once NIS + SAP numbers are
    captured, the workflow proceeds automatically without waiting for
    manual approval emails.
    """
    with _LOCK:
        conn = _get_conn()
        rows = conn.execute(
            "SELECT * FROM booking_rows "
            "WHERE status = 'done' "
            "AND (hod_approval_status IS NULL OR hod_approval_status = 'pending' "
            "     OR fc_approval_status IS NULL OR fc_approval_status = 'pending') "
            "ORDER BY row_id",
        ).fetchall()
    return _rows_to_dicts(rows)


def get_rows_awaiting_fc() -> list:
    """Booking rows that have HOD approval but not yet FC approval."""
    with _LOCK:
        conn = _get_conn()
        rows = conn.execute(
            "SELECT * FROM booking_rows "
            "WHERE status = 'done' "
            "AND hod_approval_status = 'done' "
            "AND (fc_approval_status IS NULL OR fc_approval_status = 'pending') "
            "ORDER BY row_id",
        ).fetchall()
    return _rows_to_dicts(rows)


def get_rows_awaiting_pdf_merge() -> list:
    """Booking rows that have FC approval but PDF not yet merged."""
    with _LOCK:
        conn = _get_conn()
        rows = conn.execute(
            "SELECT * FROM booking_rows "
            "WHERE status = 'done' "
            "AND fc_approval_status = 'done' "
            "AND (pdf_merge_status IS NULL OR pdf_merge_status = 'pending') "
            "ORDER BY row_id",
        ).fetchall()
    return _rows_to_dicts(rows)


def get_rows_awaiting_completion_email() -> list:
    """Booking rows fully done (all stages) but completion email not yet sent."""
    with _LOCK:
        conn = _get_conn()
        rows = conn.execute(
            "SELECT * FROM booking_rows "
            "WHERE status = 'done' "
            "AND hod_approval_status = 'done' "
            "AND fc_approval_status = 'done' "
            "AND pdf_merge_status = 'done' "
            "AND (completion_email_sent IS NULL OR completion_email_sent = 0) "
            "ORDER BY row_id",
        ).fetchall()
    return _rows_to_dicts(rows)


def is_row_all_stages_done(row_id) -> bool:
    """True when NIS + SAP + HOD + FC + PDF merge are all complete."""
    row = get_booking_row(row_id)
    if not row:
        return False
    return (
        row.get("status") == STATUS_DONE
        and row.get("hod_approval_status") == STATUS_DONE
        and row.get("fc_approval_status") == STATUS_DONE
        and row.get("pdf_merge_status") == STATUS_DONE
    )


# ─────────────────────────────────────────────────────────────────────────────
# Diagnostics
# ─────────────────────────────────────────────────────────────────────────────
def print_run_summary(run_id):
    run = get_run_by_id(run_id)
    if not run:
        print(f"\n[DB] Run {run_id} not found\n")
        return
    steps = get_steps(run_id)
    booking_rows = get_booking_rows(run_id)
    icons = {STATUS_DONE: "✅", STATUS_FAILED: "❌", STATUS_RUNNING: "🔄",
             STATUS_SKIPPED: "⏭ ", STATUS_PENDING: "⏳"}
    print(f"\n{'═' * 68}")
    print(f"  RUN  {run_id}  |  {run.get('subject', '')}")
    print(f"  Overall  : {run.get('status', '')}")
    print(f"  SPV      : {run.get('spv_name', '')}")
    print(f"  Started  : {run.get('started_at', '')}   Finished : {run.get('finished_at', '')}")
    if run.get("error_message"):
        print(f"  Error    : {str(run.get('error_message'))[:120]}")

    if booking_rows:
        print("  Booking rows (independent per-SPV units):")
        for r in booking_rows:
            icon = icons.get(r.get("status"), "?")
            nis = str(r.get("nis_checklist_id") or "").strip()
            nis_tag = f"  NIS={nis}" if nis else ""
            retry = int(r.get("retry_count") or 0)
            retry_tag = f"  [retries={retry}]" if retry else ""
            print(f"    {icon}  row#{r.get('row_index')} "
                  f"{str(r.get('spv_name') or ''):<28} {r.get('status', '')}{nis_tag}{retry_tag}")
            if r.get("error_message"):
                print(f"         └─ {str(r.get('error_message'))[:90]}")

    print("  Steps    :")
    for s in steps:
        icon = icons.get(s.get("status"), "?")
        dur = ""
        if s.get("started_at") and s.get("finished_at"):
            try:
                fmt = "%Y-%m-%dT%H:%M:%SZ"
                secs = int((datetime.strptime(str(s.get("finished_at")), fmt) -
                            datetime.strptime(str(s.get("started_at")), fmt)).total_seconds())
                dur = f" ({secs}s)"
            except Exception:
                pass
        retry_count = int(s.get("retry_count") or 0)
        retry_tag = f"  [retries={retry_count}]" if retry_count else ""
        print(f"    {icon}  {str(s.get('step_name', '')):<38} {s.get('status', '')}{dur}{retry_tag}")
        if s.get("error"):
            print(f"         └─ {str(s.get('error'))[:90]}")
    print(f"{'═' * 68}\n")


# ─────────────────────────────────────────────────────────────────────────────
# Debit/Credit Note email tracking (5-row DSM table trigger)
# ─────────────────────────────────────────────────────────────────────────────
def _dc_row_to_dict(row) -> Optional[dict]:
    return dict(row) if row is not None else None


def register_debit_credit_email(run_id: str, message_id: str, fields: dict) -> Optional[int]:
    """Create (idempotently) a pending debit/credit-email record for a DSM email.

    Keyed by the DSM email ``message_id`` (UNIQUE). Re-registering refreshes the
    snapshot fields but never downgrades an already-sent record back to pending.
    A stable ``workflow_id`` links the original email, the sent debit/credit
    email and the finance reply for end-to-end tracking. Returns the record id.
    """
    import hashlib
    now = _now()
    workflow_id = "WF-" + hashlib.sha1(str(message_id).encode("utf-8")).hexdigest()[:12].upper()
    allowed = {k: v for k, v in (fields or {}).items()
               if k in _DEBIT_CREDIT_COLUMNS
               and k not in ("id", "message_id", "run_id", "created_at", "workflow_id")}
    with _LOCK:
        conn = _get_conn()
        with conn:
            existing = conn.execute(
                "SELECT id, email_status, workflow_id FROM debit_credit_emails WHERE message_id = ?",
                (str(message_id),),
            ).fetchone()
            if existing:
                status = str(existing["email_status"] or "")
                # Backfill workflow_id on legacy rows, and refresh content while
                # still pending/failed (not yet sent).
                updates = dict(allowed) if status in ("", STATUS_PENDING, STATUS_FAILED) else {}
                if not str(existing["workflow_id"] or "").strip():
                    updates["workflow_id"] = workflow_id
                if updates:
                    cols = list(updates.keys())
                    assignments = ", ".join(f"{c} = ?" for c in cols)
                    params = [updates[c] for c in cols] + [existing["id"]]
                    conn.execute(
                        f"UPDATE debit_credit_emails SET {assignments} WHERE id = ?",
                        params,
                    )
                return int(existing["id"])

            cols = ["workflow_id", "run_id", "message_id", "email_status", "created_at"] + list(allowed.keys())
            vals = [workflow_id, str(run_id), str(message_id), STATUS_PENDING, now] + [allowed[c] for c in allowed]
            placeholders = ", ".join("?" for _ in cols)
            if _is_postgres():
                cur = conn.execute(
                    f"INSERT INTO debit_credit_emails ({', '.join(cols)}) VALUES ({placeholders}) RETURNING id",
                    vals,
                )
                row = cur.fetchone()
                return int(row["id"]) if row else None
            else:
                cur = conn.execute(
                    f"INSERT INTO debit_credit_emails ({', '.join(cols)}) VALUES ({placeholders})",
                    vals,
                )
                return int(cur.lastrowid)


def get_debit_credit_email_by_message_id(message_id: str) -> Optional[dict]:
    with _LOCK:
        conn = _get_conn()
        row = conn.execute(
            "SELECT * FROM debit_credit_emails WHERE message_id = ?",
            (str(message_id),),
        ).fetchone()
    return _dc_row_to_dict(row)


def get_debit_credit_email_by_id(dc_id) -> Optional[dict]:
    with _LOCK:
        conn = _get_conn()
        row = conn.execute(
            "SELECT * FROM debit_credit_emails WHERE id = ?", (dc_id,),
        ).fetchone()
    return _dc_row_to_dict(row)


def _update_debit_credit_email(dc_id, fields: dict):
    allowed = {k: v for k, v in (fields or {}).items()
               if k in _DEBIT_CREDIT_COLUMNS and k != "id"}
    if not allowed:
        return
    with _LOCK:
        conn = _get_conn()
        with conn:
            cols = list(allowed.keys())
            assignments = ", ".join(f"{c} = ?" for c in cols)
            params = [allowed[c] for c in cols] + [dc_id]
            conn.execute(
                f"UPDATE debit_credit_emails SET {assignments} WHERE id = ?", params,
            )


def get_pending_debit_credit_emails() -> list:
    """Records that still need to be SENT (status pending/failed)."""
    with _LOCK:
        conn = _get_conn()
        rows = conn.execute(
            "SELECT * FROM debit_credit_emails "
            "WHERE email_status IN (?, ?) ORDER BY id",
            (STATUS_PENDING, STATUS_FAILED),
        ).fetchall()
    return _rows_to_dicts(rows)


def get_debit_credit_emails_awaiting_reply() -> list:
    """Sent records whose finance-team reply PDFs have not yet been stored."""
    with _LOCK:
        conn = _get_conn()
        rows = conn.execute(
            "SELECT * FROM debit_credit_emails "
            "WHERE email_status = ? AND COALESCE(reply_pdfs_uploaded, 0) = 0 "
            "ORDER BY id",
            (STATUS_DONE,),
        ).fetchall()
    return _rows_to_dicts(rows)


def mark_debit_credit_email_sent(dc_id, sent_subject: str, recipients, sent_at: str = ""):
    rec = recipients if isinstance(recipients, str) else ", ".join(recipients or [])
    _update_debit_credit_email(dc_id, {
        "email_status": STATUS_DONE,
        "sent_subject": str(sent_subject or ""),
        "recipients": rec,
        "sent_at": sent_at or _now(),
        "error_message": "",
    })


def mark_debit_credit_email_failed(dc_id, error: str):
    _update_debit_credit_email(dc_id, {
        "email_status": STATUS_FAILED,
        "error_message": str(error or "")[:2000],
    })


def mark_debit_credit_reply_received(dc_id, reply_message_id: str, pdf_names,
                                     uploaded: bool = True, received_at: str = ""):
    names = pdf_names if isinstance(pdf_names, str) else ", ".join(pdf_names or [])
    _update_debit_credit_email(dc_id, {
        "reply_message_id": str(reply_message_id or ""),
        "reply_pdf_names": names,
        "reply_pdfs_uploaded": int(bool(uploaded)),
        "reply_received_at": received_at or _now(),
    })


# ─────────────────────────────────────────────────────────────────────────────
# Debit/Credit reply PDF tracking (extraction → vendor → SAP checklist → upload)
# ─────────────────────────────────────────────────────────────────────────────
def upsert_debit_credit_pdf(dc_email_id, pdf_name: str, fields: dict) -> Optional[int]:
    """Create/refresh a debit_credit_pdfs row (UNIQUE dc_email_id + pdf_name).

    Idempotent per (email, pdf). Existing SAP results are never overwritten by a
    re-extraction unless explicitly provided. Returns the row id.
    """
    now = _now()
    allowed = {k: v for k, v in (fields or {}).items()
               if k in _DEBIT_CREDIT_PDF_COLUMNS
               and k not in ("id", "dc_email_id", "pdf_name", "created_at")}
    with _LOCK:
        conn = _get_conn()
        with conn:
            existing = conn.execute(
                "SELECT id FROM debit_credit_pdfs WHERE dc_email_id = ? AND pdf_name = ?",
                (dc_email_id, str(pdf_name)),
            ).fetchone()
            if existing:
                allowed["updated_at"] = now
                cols = list(allowed.keys())
                if cols:
                    assignments = ", ".join(f"{c} = ?" for c in cols)
                    params = [allowed[c] for c in cols] + [existing["id"]]
                    conn.execute(
                        f"UPDATE debit_credit_pdfs SET {assignments} WHERE id = ?", params,
                    )
                return int(existing["id"])

            cols = ["dc_email_id", "pdf_name", "created_at", "updated_at"] + list(allowed.keys())
            vals = [dc_email_id, str(pdf_name), now, now] + [allowed[c] for c in allowed]
            placeholders = ", ".join("?" for _ in cols)
            if _is_postgres():
                cur = conn.execute(
                    f"INSERT INTO debit_credit_pdfs ({', '.join(cols)}) VALUES ({placeholders}) RETURNING id",
                    vals,
                )
                row = cur.fetchone()
                return int(row["id"]) if row else None
            else:
                cur = conn.execute(
                    f"INSERT INTO debit_credit_pdfs ({', '.join(cols)}) VALUES ({placeholders})",
                    vals,
                )
                return int(cur.lastrowid)


def update_debit_credit_pdf(pdf_id, fields: dict):
    allowed = {k: v for k, v in (fields or {}).items()
               if k in _DEBIT_CREDIT_PDF_COLUMNS and k != "id"}
    if not allowed:
        return
    allowed["updated_at"] = _now()
    with _LOCK:
        conn = _get_conn()
        with conn:
            cols = list(allowed.keys())
            assignments = ", ".join(f"{c} = ?" for c in cols)
            params = [allowed[c] for c in cols] + [pdf_id]
            conn.execute(
                f"UPDATE debit_credit_pdfs SET {assignments} WHERE id = ?", params,
            )


def get_debit_credit_pdf_by_id(pdf_id) -> Optional[dict]:
    with _LOCK:
        conn = _get_conn()
        row = conn.execute(
            "SELECT * FROM debit_credit_pdfs WHERE id = ?", (pdf_id,),
        ).fetchone()
    return _dc_row_to_dict(row)


def get_debit_credit_pdfs_for_email(dc_email_id) -> list:
    with _LOCK:
        conn = _get_conn()
        rows = conn.execute(
            "SELECT * FROM debit_credit_pdfs WHERE dc_email_id = ? ORDER BY id",
            (dc_email_id,),
        ).fetchall()
    return _rows_to_dicts(rows)


def get_pending_sap_debit_credit_pdfs() -> list:
    """Debit/Credit PDFs that have their extracted values but no SAP checklist
    done yet (ready for SAP checklist automation)."""
    with _LOCK:
        conn = _get_conn()
        rows = conn.execute(
            "SELECT * FROM debit_credit_pdfs "
            "WHERE sap_status IN (?, ?) ORDER BY id",
            (STATUS_PENDING, STATUS_FAILED),
        ).fetchall()
    return _rows_to_dicts(rows)


def mark_debit_credit_pdf_sap_done(pdf_id, checklist_number: str,
                                   sap_pdf_name: str = "", sap_pdf_sp_path: str = ""):
    update_debit_credit_pdf(pdf_id, {
        "sap_status": STATUS_DONE,
        "sap_checklist_number": str(checklist_number or ""),
        "sap_pdf_name": str(sap_pdf_name or ""),
        "sap_pdf_sp_path": str(sap_pdf_sp_path or ""),
        "error_message": "",
    })


def mark_debit_credit_pdf_sap_failed(pdf_id, error: str):
    update_debit_credit_pdf(pdf_id, {
        "sap_status": STATUS_FAILED,
        "error_message": str(error or "")[:2000],
    })


def are_all_debit_credit_pdfs_sap_done(dc_email_id) -> bool:
    pdfs = get_debit_credit_pdfs_for_email(dc_email_id)
    if not pdfs:
        return False
    return all(p.get("sap_status") in (STATUS_DONE, STATUS_SKIPPED) for p in pdfs)
