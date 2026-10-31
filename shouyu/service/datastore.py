"""SQLite tables that make SQLite (not Excel) the source of truth for
today's tasks, the backlog, pomodoros and clipboard notes.

Design reference: docs/design/planning/20261031-01-tasks-to-sqlite.md (§3).
The tables live in the same file as the durable queue (`message_queue.py`),
so a data write and the "mirror this to Excel" queue message it produces can
commit in one transaction, and `note` can reference queue `attachments`.
"""
from __future__ import annotations

import contextlib
import logging
import sqlite3
import threading
from typing import Dict, List, Optional

from shouyu.service import message_queue

_BUSY_TIMEOUT_MS = 5000

META_SCHEMA_VERSION = 'schema_version'
META_IMPORT_FINISHED_AT = 'excel_import_finished_at'
META_IMPORT_SOURCE = 'excel_import_source'

_META_DDL = """
CREATE TABLE IF NOT EXISTS meta (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
)
"""

# One entry per schema version; each is applied exactly once, in order.
_MIGRATIONS: Dict[int, List[str]] = {
    1: [
        """
        CREATE TABLE IF NOT EXISTS task (
            id               INTEGER PRIMARY KEY,
            text             TEXT    NOT NULL,
            category         TEXT    NOT NULL DEFAULT 'work' CHECK (category IN ('work', 'life')),
            priority         TEXT    NOT NULL DEFAULT ''     CHECK (priority IN ('', 'P1', 'P2', 'P3')),
            duration_minutes INTEGER NOT NULL DEFAULT 0      CHECK (duration_minutes >= 0),
            reflection       TEXT    NOT NULL DEFAULT '',
            created_at       TEXT    NOT NULL,
            updated_at       TEXT    NOT NULL,
            deleted_at       TEXT,
            source           TEXT    NOT NULL DEFAULT 'ui'   CHECK (source IN ('ui', 'excel_import')),
            source_ref       TEXT
        )
        """,
        # Re-running the Excel import must not create duplicates.
        "CREATE UNIQUE INDEX IF NOT EXISTS ux_task_source_ref ON task(source, source_ref) "
        "WHERE source_ref IS NOT NULL",
        """
        CREATE TABLE IF NOT EXISTS day_plan_entry (
            id         INTEGER PRIMARY KEY,
            plan_date  TEXT    NOT NULL,
            task_id    INTEGER NOT NULL REFERENCES task(id),
            position   INTEGER NOT NULL,
            status     TEXT    NOT NULL CHECK (status IN ('pending', 'in_progress', 'suspended', 'done')),
            created_at TEXT    NOT NULL,
            updated_at TEXT    NOT NULL,
            UNIQUE (plan_date, task_id)
        )
        """,
        "CREATE INDEX IF NOT EXISTS ix_day_plan_entry_date ON day_plan_entry(plan_date, position)",
        "CREATE INDEX IF NOT EXISTS ix_day_plan_entry_task ON day_plan_entry(task_id)",
        """
        CREATE TABLE IF NOT EXISTS backlog_entry (
            task_id    INTEGER PRIMARY KEY REFERENCES task(id),
            position   INTEGER NOT NULL,
            entered_on TEXT    NOT NULL,
            created_at TEXT    NOT NULL
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS task_status_log (
            id          INTEGER PRIMARY KEY,
            task_id     INTEGER NOT NULL REFERENCES task(id),
            plan_date   TEXT    NOT NULL,
            from_status TEXT,
            to_status   TEXT    NOT NULL,
            changed_at  TEXT    NOT NULL
        )
        """,
        "CREATE INDEX IF NOT EXISTS ix_task_status_log_task ON task_status_log(task_id, changed_at)",
        """
        CREATE TABLE IF NOT EXISTS pomodoro (
            id              INTEGER PRIMARY KEY,
            task_id         INTEGER REFERENCES task(id),
            task_text       TEXT    NOT NULL DEFAULT '',
            plan_date       TEXT    NOT NULL,
            kind            TEXT    NOT NULL DEFAULT 'work' CHECK (kind IN ('work', 'planning')),
            mode            TEXT    CHECK (mode IN ('classic', 'deep')),
            started_at      TEXT,
            ended_at        TEXT,
            planned_minutes INTEGER,
            actual_minutes  INTEGER,
            ended_early     INTEGER NOT NULL DEFAULT 0 CHECK (ended_early IN (0, 1)),
            source          TEXT    NOT NULL DEFAULT 'ui' CHECK (source IN ('ui', 'excel_import')),
            source_ref      TEXT
        )
        """,
        "CREATE INDEX IF NOT EXISTS ix_pomodoro_date ON pomodoro(plan_date, kind)",
        "CREATE UNIQUE INDEX IF NOT EXISTS ux_pomodoro_source_ref ON pomodoro(source, source_ref) "
        "WHERE source_ref IS NOT NULL",
        # `attachments` is created by message_queue.init_db(), which must run first.
        """
        CREATE TABLE IF NOT EXISTS note (
            id            INTEGER PRIMARY KEY,
            plan_date     TEXT    NOT NULL,
            kind          TEXT    NOT NULL CHECK (kind IN ('text', 'image')),
            text          TEXT    NOT NULL DEFAULT '',
            attachment_id INTEGER REFERENCES attachments(id),
            source        TEXT    NOT NULL CHECK (source IN ('clipboard', 'ocr')),
            excel_column  TEXT    CHECK (excel_column IN ('A', 'B')),
            created_at    TEXT    NOT NULL
        )
        """,
        "CREATE INDEX IF NOT EXISTS ix_note_date ON note(plan_date)",
    ],
}

SCHEMA_VERSION = max(_MIGRATIONS)

_init_lock = threading.Lock()
_initialized_paths: set = set()


def _open() -> sqlite3.Connection:
    # isolation_level=None: no implicit transactions, so `transaction()` can
    # issue BEGIN IMMEDIATE itself (a deferred read-then-write transaction in
    # WAL mode fails with SQLITE_BUSY_SNAPSHOT instead of waiting).
    conn = sqlite3.connect(
        message_queue.db_path(), timeout=_BUSY_TIMEOUT_MS / 1000, isolation_level=None
    )
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA journal_mode=WAL')
    conn.execute(f'PRAGMA busy_timeout={_BUSY_TIMEOUT_MS}')
    conn.execute('PRAGMA foreign_keys=ON')
    return conn


@contextlib.contextmanager
def transaction():
    """Short-lived write transaction. Takes the write lock up front."""
    ensure_schema()
    conn = _open()
    try:
        conn.execute('BEGIN IMMEDIATE')
        try:
            yield conn
            conn.execute('COMMIT')
        except BaseException:
            conn.execute('ROLLBACK')
            raise
    finally:
        conn.close()


@contextlib.contextmanager
def reader():
    """Short-lived read-only snapshot, consistent across several SELECTs."""
    ensure_schema()
    conn = _open()
    try:
        conn.execute('BEGIN')
        try:
            yield conn
        finally:
            conn.execute('COMMIT')
    finally:
        conn.close()


def init_schema() -> None:
    """Create / upgrade the tables. Safe to call on every startup."""
    message_queue.init_db()
    conn = _open()
    try:
        conn.execute('BEGIN IMMEDIATE')
        try:
            conn.execute(_META_DDL)
            row = conn.execute(
                "SELECT value FROM meta WHERE key = ?", (META_SCHEMA_VERSION,)
            ).fetchone()
            current = int(row['value']) if row else 0
            for version in sorted(v for v in _MIGRATIONS if v > current):
                for statement in _MIGRATIONS[version]:
                    conn.execute(statement)
                set_meta(conn, META_SCHEMA_VERSION, str(version))
                logging.info(f'datastore schema upgraded to version {version}')
            conn.execute('COMMIT')
        except BaseException:
            conn.execute('ROLLBACK')
            raise
    finally:
        conn.close()
    with _init_lock:
        _initialized_paths.add(message_queue.db_path())


def ensure_schema() -> None:
    """Lazily run init_schema() once per database path in this process."""
    path = message_queue.db_path()
    with _init_lock:
        if path in _initialized_paths:
            return
    init_schema()


def get_meta(conn: sqlite3.Connection, key: str) -> Optional[str]:
    row = conn.execute("SELECT value FROM meta WHERE key = ?", (key,)).fetchone()
    return row['value'] if row else None


def set_meta(conn: sqlite3.Connection, key: str, value: str) -> None:
    conn.execute(
        "INSERT INTO meta (key, value) VALUES (?, ?) "
        "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
        (key, value),
    )
