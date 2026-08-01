"""Durable local message queue backing Excel-save resilience.

Design reference: docs/excel-save-resilience.md (§8). This module owns two
sqlite tables in `shouyu.db` (co-located with kb.ini):

  queue        - append-only ledger of writes waiting to be dispatched to
                 their eventual home (today: kb.xlsx). Rows are never
                 deleted automatically; a consumer's progress is tracked by
                 `queue_cursor`, not by row deletion.
  attachments  - binary/text payloads (screenshots, future pdf/xlsx/zip)
                 referenced by a queue row's `attachment_id`. Images are
                 always stored inline (BLOB); other types spill to an
                 external `attachments/<YYYY-MM-DD>/` file above
                 NON_IMAGE_INLINE_MAX_BYTES.

Every write path (habit_dialog plan-save, todo_panel plan-save, clipboard
save) should call `enqueue()` and return immediately; the actual Excel
write happens later on the shared background dispatcher (see dispatch.py).
"""
from __future__ import annotations

import contextlib
import hashlib
import json
import logging
import os
import re
import sqlite3
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any, List, Optional

from shouyu.config import Config

DB_FILE_NAME = 'shouyu.db'

# Images are always inlined regardless of size, up to this hard safety-valve
# cap; beyond it (an outlier, not the common case) they spill to an external
# file just like an oversized non-image attachment would.
IMAGE_INLINE_HARD_CAP_BYTES = 20 * 1024 * 1024
# Non-image attachments (future pdf/xlsx/zip/txt) inline up to this size;
# above it they spill to an external, day-bucketed file.
NON_IMAGE_INLINE_MAX_BYTES = 5 * 1024 * 1024

MAX_ATTEMPTS_BEFORE_DEAD = 10
# A row claimed 'processing' whose claimed_at is older than this is assumed
# to be orphaned by a crash mid-dispatch, and is reset back to 'pending'.
STALE_PROCESSING_MINUTES = 5

_BUSY_TIMEOUT_MS = 5000


def _db_path() -> str:
    base_dir = os.path.dirname(os.path.abspath(Config.FILE_NAME)) or '.'
    return os.path.join(base_dir, DB_FILE_NAME)


def _attachments_root() -> str:
    base_dir = os.path.dirname(os.path.abspath(Config.FILE_NAME)) or '.'
    return os.path.join(base_dir, 'attachments')


def _utcnow_iso() -> str:
    return datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3] + 'Z'


@contextlib.contextmanager
def _connect():
    """Short-lived connection per call, per sqlite's recommended multi-thread
    pattern: WAL mode + a busy_timeout so writers block-and-retry instead of
    raising `database is locked`."""
    conn = sqlite3.connect(_db_path(), timeout=_BUSY_TIMEOUT_MS / 1000)
    conn.row_factory = sqlite3.Row
    try:
        conn.execute('PRAGMA journal_mode=WAL')
        conn.execute(f'PRAGMA busy_timeout={_BUSY_TIMEOUT_MS}')
        conn.execute('PRAGMA foreign_keys=ON')
        with conn:
            yield conn
    finally:
        conn.close()


def init_db() -> None:
    """Create tables/indexes if they don't exist yet. Safe to call on every startup."""
    with _connect() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS attachments (
                id         INTEGER PRIMARY KEY,
                created_at TEXT NOT NULL,
                sha256     TEXT NOT NULL,
                mime_type  TEXT,
                byte_size  INTEGER NOT NULL,
                storage    TEXT NOT NULL CHECK(storage IN ('inline', 'file')),
                data       BLOB,
                file_path  TEXT
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS queue (
                id             INTEGER PRIMARY KEY,
                created_at     TEXT NOT NULL,
                updated_at     TEXT NOT NULL,
                kind           TEXT NOT NULL,
                payload        TEXT NOT NULL CHECK(json_valid(payload)),
                attachment_id  INTEGER REFERENCES attachments(id),
                status         TEXT NOT NULL DEFAULT 'pending'
                                CHECK(status IN ('pending', 'processing', 'done', 'dead')),
                attempts       INTEGER NOT NULL DEFAULT 0,
                claimed_at     TEXT,
                last_error     TEXT
            )
            """
        )
        conn.execute('CREATE INDEX IF NOT EXISTS idx_queue_status_created ON queue(status, created_at)')
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS queue_cursor (
                consumer   TEXT PRIMARY KEY,
                last_id    INTEGER NOT NULL DEFAULT 0,
                updated_at TEXT NOT NULL
            )
            """
        )


# ---------- attachments ----------

_UNSAFE_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')


def _sanitize_filename(name: str) -> str:
    name = (name or '').strip()
    name = _UNSAFE_FILENAME_CHARS.sub('_', name)
    return name[:120] or 'attachment'


def store_attachment(data: bytes, mime_type: str = '', filename: str = '') -> int:
    """Persist a binary/text payload and return its attachment id.

    Images are always inlined (up to a hard safety-valve cap); other types
    inline up to NON_IMAGE_INLINE_MAX_BYTES and spill to an external,
    day-bucketed file above it. See docs/excel-save-resilience.md §8.2.
    """
    digest = hashlib.sha256(data).hexdigest()
    byte_size = len(data)
    is_image = (mime_type or '').startswith('image/')
    inline_cap = IMAGE_INLINE_HARD_CAP_BYTES if is_image else NON_IMAGE_INLINE_MAX_BYTES

    if byte_size <= inline_cap:
        storage, blob, file_path = 'inline', data, None
    else:
        storage, blob = 'file', None
        file_path = _write_external_file(data, digest, filename)

    with _connect() as conn:
        cur = conn.execute(
            "INSERT INTO attachments (created_at, sha256, mime_type, byte_size, storage, data, file_path) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (_utcnow_iso(), digest, mime_type, byte_size, storage, blob, file_path),
        )
        return cur.lastrowid


def _write_external_file(data: bytes, digest: str, filename: str) -> str:
    root = _attachments_root()
    # Bucketed by *local* calendar day (this is for human browsing/partial
    # migration, not for the UTC timestamps stored in the DB columns).
    day_dir = os.path.join(root, time.strftime('%Y-%m-%d'))
    os.makedirs(day_dir, exist_ok=True)
    safe_name = _sanitize_filename(filename)
    relative_path = os.path.join(time.strftime('%Y-%m-%d'), f'{digest[:16]}_{safe_name}')
    target = os.path.join(root, relative_path)
    if not os.path.exists(target):
        tmp = target + f'.tmp_{os.getpid()}'
        with open(tmp, 'wb') as f:
            f.write(data)
        os.replace(tmp, target)
    return relative_path


def get_attachment_bytes(attachment_id: int) -> bytes:
    with _connect() as conn:
        row = conn.execute(
            "SELECT storage, data, file_path FROM attachments WHERE id = ?", (attachment_id,)
        ).fetchone()
    if row is None:
        raise KeyError(f'attachment {attachment_id} not found')
    if row['storage'] == 'inline':
        return row['data']
    with open(os.path.join(_attachments_root(), row['file_path']), 'rb') as f:
        return f.read()


# ---------- queue ----------

@dataclass
class QueueMessage:
    id: int
    created_at: str
    kind: str
    payload: Any
    attachment_id: Optional[int]
    attempts: int


def enqueue(kind: str, payload: Any, attachment_id: Optional[int] = None) -> int:
    now = _utcnow_iso()
    body = json.dumps(payload, ensure_ascii=False)
    with _connect() as conn:
        cur = conn.execute(
            "INSERT INTO queue (created_at, updated_at, kind, payload, attachment_id, status, attempts) "
            "VALUES (?, ?, ?, ?, ?, 'pending', 0)",
            (now, now, kind, body, attachment_id),
        )
        return cur.lastrowid


def get_cursor(consumer: str) -> int:
    with _connect() as conn:
        row = conn.execute(
            "SELECT last_id FROM queue_cursor WHERE consumer = ?", (consumer,)
        ).fetchone()
    return row['last_id'] if row else 0


def fetch_pending_batch(consumer: str, limit: int = 200) -> List[QueueMessage]:
    """Return the next batch of pending messages after this consumer's cursor.

    Also resets any 'processing' rows whose claimed_at is stale (crash
    mid-dispatch) back to 'pending' so they get picked up again.
    """
    now = _utcnow_iso()
    stale_cutoff = datetime.now(timezone.utc).timestamp() - STALE_PROCESSING_MINUTES * 60
    with _connect() as conn:
        stuck = conn.execute(
            "SELECT id, claimed_at FROM queue WHERE status = 'processing'"
        ).fetchall()
        for row in stuck:
            claimed_at = row['claimed_at']
            if not claimed_at or _parse_iso(claimed_at) < stale_cutoff:
                conn.execute(
                    "UPDATE queue SET status = 'pending', updated_at = ? WHERE id = ?",
                    (now, row['id']),
                )

        cursor = conn.execute(
            "SELECT last_id FROM queue_cursor WHERE consumer = ?", (consumer,)
        ).fetchone()
        last_id = cursor['last_id'] if cursor else 0
        rows = conn.execute(
            "SELECT id, created_at, kind, payload, attachment_id, attempts FROM queue "
            "WHERE id > ? AND status = 'pending' ORDER BY id LIMIT ?",
            (last_id, limit),
        ).fetchall()
    return [
        QueueMessage(
            id=r['id'],
            created_at=r['created_at'],
            kind=r['kind'],
            payload=json.loads(r['payload']),
            attachment_id=r['attachment_id'],
            attempts=r['attempts'],
        )
        for r in rows
    ]


def _parse_iso(value: str) -> float:
    """Parse the ISO-8601 UTC strings written by `_utcnow_iso()`."""
    try:
        return datetime.strptime(value, '%Y-%m-%dT%H:%M:%S.%fZ').replace(tzinfo=timezone.utc).timestamp()
    except Exception:
        try:
            return datetime.fromisoformat(value.replace('Z', '+00:00')).timestamp()
        except Exception:
            return 0.0


def mark_processing(ids: List[int]) -> None:
    if not ids:
        return
    now = _utcnow_iso()
    with _connect() as conn:
        conn.executemany(
            "UPDATE queue SET status = 'processing', claimed_at = ?, updated_at = ? WHERE id = ?",
            [(now, now, i) for i in ids],
        )


def mark_done(ids: List[int]) -> None:
    if not ids:
        return
    now = _utcnow_iso()
    with _connect() as conn:
        conn.executemany(
            "UPDATE queue SET status = 'done', updated_at = ? WHERE id = ?",
            [(now, i) for i in ids],
        )


def mark_pending_retry(ids: List[int], error: str) -> List[int]:
    """Bump attempts and put messages back to 'pending' (or 'dead' once
    MAX_ATTEMPTS_BEFORE_DEAD is exceeded). Returns the ids that became dead."""
    if not ids:
        return []
    now = _utcnow_iso()
    dead: List[int] = []
    with _connect() as conn:
        for i in ids:
            row = conn.execute("SELECT attempts FROM queue WHERE id = ?", (i,)).fetchone()
            attempts = (row['attempts'] if row else 0) + 1
            status = 'dead' if attempts >= MAX_ATTEMPTS_BEFORE_DEAD else 'pending'
            if status == 'dead':
                dead.append(i)
            conn.execute(
                "UPDATE queue SET status = ?, attempts = ?, updated_at = ?, last_error = ? WHERE id = ?",
                (status, attempts, now, str(error)[:2000], i),
            )
    return dead


def mark_dead(message_id: int, error: str) -> None:
    now = _utcnow_iso()
    with _connect() as conn:
        conn.execute(
            "UPDATE queue SET status = 'dead', updated_at = ?, last_error = ? WHERE id = ?",
            (now, str(error)[:2000], message_id),
        )


def advance_cursor(consumer: str) -> None:
    """Move the consumer's cursor forward over every contiguous
    done/dead row starting right after its current position."""
    now = _utcnow_iso()
    with _connect() as conn:
        row = conn.execute(
            "SELECT last_id FROM queue_cursor WHERE consumer = ?", (consumer,)
        ).fetchone()
        last_id = row['last_id'] if row else 0
        while True:
            nxt = conn.execute(
                "SELECT id, status FROM queue WHERE id = ?", (last_id + 1,)
            ).fetchone()
            if nxt is None or nxt['status'] not in ('done', 'dead'):
                break
            last_id += 1
        conn.execute(
            "INSERT INTO queue_cursor (consumer, last_id, updated_at) VALUES (?, ?, ?) "
            "ON CONFLICT(consumer) DO UPDATE SET last_id = excluded.last_id, updated_at = excluded.updated_at",
            (consumer, last_id, now),
        )


def stuck_info(consumer: str, threshold_minutes: int) -> Optional[dict]:
    """Return {'count': N, 'oldest_created_at': ts} for pending messages older
    than `threshold_minutes`, or None if nothing is stuck that long."""
    cutoff = datetime.now(timezone.utc).timestamp() - threshold_minutes * 60
    with _connect() as conn:
        cursor = conn.execute(
            "SELECT last_id FROM queue_cursor WHERE consumer = ?", (consumer,)
        ).fetchone()
        last_id = cursor['last_id'] if cursor else 0
        rows = conn.execute(
            "SELECT created_at FROM queue WHERE id > ? AND status = 'pending' ORDER BY id",
            (last_id,),
        ).fetchall()
    if not rows:
        return None
    oldest = rows[0]['created_at']
    if _parse_iso(oldest) > cutoff:
        return None
    return {'count': len(rows), 'oldest_created_at': oldest}


def purge_resolved(older_than_days: int) -> int:
    """Manual cleanup: delete 'done'/'dead' rows (and their attachments,
    including any external files) older than `older_than_days`. Never
    touches 'pending'/'processing' rows. Returns the number of rows removed."""
    cutoff_ts = datetime.now(timezone.utc).timestamp() - older_than_days * 86400
    removed = 0
    with _connect() as conn:
        rows = conn.execute(
            "SELECT id, created_at, attachment_id FROM queue WHERE status IN ('done', 'dead')"
        ).fetchall()
        for row in rows:
            if _parse_iso(row['created_at']) >= cutoff_ts:
                continue
            attachment_id = row['attachment_id']
            conn.execute("DELETE FROM queue WHERE id = ?", (row['id'],))
            removed += 1
            if attachment_id is not None:
                still_referenced = conn.execute(
                    "SELECT 1 FROM queue WHERE attachment_id = ? LIMIT 1", (attachment_id,)
                ).fetchone()
                if still_referenced is None:
                    att = conn.execute(
                        "SELECT storage, file_path FROM attachments WHERE id = ?", (attachment_id,)
                    ).fetchone()
                    if att is not None:
                        if att['storage'] == 'file' and att['file_path']:
                            _safe_remove(os.path.join(_attachments_root(), att['file_path']))
                        conn.execute("DELETE FROM attachments WHERE id = ?", (attachment_id,))
    if removed:
        # VACUUM cannot run inside a transaction; do it on its own connection
        # after the delete transaction above has committed and closed.
        vacuum_conn = sqlite3.connect(_db_path(), timeout=_BUSY_TIMEOUT_MS / 1000)
        try:
            vacuum_conn.execute('VACUUM')
        finally:
            vacuum_conn.close()
    return removed


def _safe_remove(path: str) -> None:
    try:
        if os.path.exists(path):
            os.remove(path)
    except Exception:
        logging.exception(f'failed to remove attachment file: {path}')


def dump_human_readable(limit: int = 200) -> str:
    """Debug helper: a plain-text dump of the most recent queue rows."""
    with _connect() as conn:
        rows = conn.execute(
            "SELECT id, created_at, kind, status, attempts, last_error, attachment_id "
            "FROM queue ORDER BY id DESC LIMIT ?",
            (limit,),
        ).fetchall()
        cursors = conn.execute("SELECT consumer, last_id, updated_at FROM queue_cursor").fetchall()
    lines = [f'shouyu 待同步队列 - {_db_path()}', '']
    lines.append('游标:')
    for c in cursors:
        lines.append(f"  {c['consumer']}: last_id={c['last_id']} ({c['updated_at']})")
    lines.append('')
    lines.append(f'最近 {len(rows)} 条 (新→旧):')
    for r in rows:
        attach = f" attachment={r['attachment_id']}" if r['attachment_id'] else ''
        err = f" error={r['last_error']}" if r['last_error'] else ''
        lines.append(
            f"  #{r['id']:>6}  {r['created_at']}  {r['kind']:<16}  {r['status']:<10}  "
            f"attempts={r['attempts']}{attach}{err}"
        )
    return '\n'.join(lines)
