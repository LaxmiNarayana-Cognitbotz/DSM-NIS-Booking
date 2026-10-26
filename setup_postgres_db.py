"""
setup_postgres_db.py
====================
One-time script to create the 'dsm_nis_booking' database on the Azure
PostgreSQL server and initialize all required tables.

Run once:
    python setup_postgres_db.py

After this script succeeds, set DB_BACKEND=postgres in your .env (or env vars)
and the application will use this PostgreSQL database.
"""

import sys

try:
    import psycopg2
    from psycopg2 import sql
    from psycopg2.extensions import ISOLATION_LEVEL_AUTOCOMMIT
except ImportError:
    sys.exit("psycopg2 is not installed. Run: pip install psycopg2-binary")

# ─── Connection credentials ──────────────────────────────────────────────────
HOST     = "az10psqldmrcbtp01.postgres.database.azure.com"
PORT     = 5432
USER     = "psqladmin"
PASSWORD = "1k8h02grUu+qJ2uHZb<{lB3LF%+Yj-Ar"
SSLMODE  = "require"

# The database we want to create
DB_NAME  = "dsm_nis_booking"


def create_database():
    """Connect to the default 'postgres' database and CREATE the target DB."""
    print(f"[1/3] Connecting to {HOST}:{PORT} as '{USER}' ...")
    conn = psycopg2.connect(
        host=HOST,
        port=PORT,
        user=USER,
        password=PASSWORD,
        dbname="postgres",
        sslmode=SSLMODE,
    )
    conn.set_isolation_level(ISOLATION_LEVEL_AUTOCOMMIT)
    cur = conn.cursor()

    # Check if database already exists
    cur.execute(
        "SELECT 1 FROM pg_database WHERE datname = %s", (DB_NAME,)
    )
    if cur.fetchone():
        print(f"       Database '{DB_NAME}' already exists. Skipping creation.")
    else:
        print(f"       Creating database '{DB_NAME}' ...")
        cur.execute(sql.SQL("CREATE DATABASE {}").format(sql.Identifier(DB_NAME)))
        print(f"       Database '{DB_NAME}' created successfully.")

    cur.close()
    conn.close()


def create_tables():
    """Connect to the new database and create all application tables."""
    print(f"[2/3] Connecting to database '{DB_NAME}' to create tables ...")
    conn = psycopg2.connect(
        host=HOST,
        port=PORT,
        user=USER,
        password=PASSWORD,
        dbname=DB_NAME,
        sslmode=SSLMODE,
    )
    cur = conn.cursor()

    # ── email_runs ───────────────────────────────────────────────────────────
    cur.execute("""
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
    """)

    # ── booking_rows ─────────────────────────────────────────────────────────
    cur.execute("""
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
    """)

    # ── run_steps ────────────────────────────────────────────────────────────
    cur.execute("""
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
    """)

    # ── debit_credit_emails ──────────────────────────────────────────────────
    cur.execute("""
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
    """)

    # ── debit_credit_pdfs ────────────────────────────────────────────────────
    cur.execute("""
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
    """)

    # ── Indexes ──────────────────────────────────────────────────────────────
    cur.execute("CREATE INDEX IF NOT EXISTS idx_steps_run ON run_steps(run_id);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_rows_run ON booking_rows(run_id);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_dc_msg ON debit_credit_emails(message_id);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_dc_wf ON debit_credit_emails(workflow_id);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_dcpdf_email ON debit_credit_pdfs(dc_email_id);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_dcpdf_wf ON debit_credit_pdfs(workflow_id);")

    conn.commit()
    cur.close()
    conn.close()
    print("       All tables and indexes created successfully.")


def verify_connection():
    """Quick verification that the app can connect to the new database."""
    print(f"[3/3] Verifying connection to '{DB_NAME}' ...")
    conn = psycopg2.connect(
        host=HOST,
        port=PORT,
        user=USER,
        password=PASSWORD,
        dbname=DB_NAME,
        sslmode=SSLMODE,
    )
    cur = conn.cursor()
    cur.execute("SELECT table_name FROM information_schema.tables WHERE table_schema = 'public' ORDER BY table_name;")
    tables = [row[0] for row in cur.fetchall()]
    cur.close()
    conn.close()

    print(f"       Tables found: {', '.join(tables)}")
    print()
    print("=" * 60)
    print("  SUCCESS! Database is ready.")
    print()
    print("  To use PostgreSQL, set these in your .env:")
    print(f"    DB_BACKEND=postgres")
    print(f"    POSTGRES_HOST={HOST}")
    print(f"    POSTGRES_PORT={PORT}")
    print(f"    POSTGRES_DB={DB_NAME}")
    print(f"    POSTGRES_USER={USER}")
    print(f"    POSTGRES_PASSWORD={PASSWORD}")
    print(f"    POSTGRES_SSLMODE={SSLMODE}")
    print("=" * 60)


if __name__ == "__main__":
    try:
        create_database()
        create_tables()
        verify_connection()
    except Exception as e:
        print(f"\nERROR: {e}")
        sys.exit(1)
