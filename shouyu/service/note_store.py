"""Clipboard / screenshot / OCR records, stored in SQLite first.

Design references:
- docs/design/planning/20261031-01-tasks-to-sqlite.md (§4.5): saving.
- docs/design/retrieval/20261031-01-notes-panel-and-day-view.md (§3.2, §4.4):
  paging, search, edit and delete for the notes panel and the day view.

A note row is the record; the Excel "other" area only receives a mirrored
copy via the existing `clipboard_append` queue message when a note is first
saved. Later edits / deletes / manual notes stay in SQLite only. Images live
in the queue's `attachments` table, which the note and the mirror message
share.
"""
from __future__ import annotations

import logging
import sqlite3
import time
from dataclasses import dataclass, field
from typing import List, Optional

from shouyu.config import Config
from shouyu.service import datastore, message_queue

KIND_TEXT = 'text'
KIND_IMAGE = 'image'
SOURCE_CLIPBOARD = 'clipboard'
SOURCE_OCR = 'ocr'
SOURCE_MANUAL = 'manual'
SOURCE_EXCEL_IMPORT = 'excel_import'

PAGE_SIZE = 200
SEARCH_LIMIT = 300

_COLUMNS = "id, plan_date, kind, text, attachment_id, source, excel_column, created_at, updated_at"


@dataclass
class Note:
    id: int
    plan_date: str
    kind: str
    text: str
    attachment_id: Optional[int]
    source: str
    excel_column: Optional[str]
    created_at: str
    updated_at: Optional[str] = None


@dataclass
class NoteQuery:
    # Every term must appear in the note text (case-insensitive for ASCII).
    terms: List[str] = field(default_factory=list)
    # Inclusive local dates (YYYY-MM-DD); None means unbounded.
    date_from: Optional[str] = None
    date_to: Optional[str] = None
    limit: int = SEARCH_LIMIT


def parse_terms(text: str) -> List[str]:
    return [t for t in (text or '').split() if t]


def _escape_like(term: str) -> str:
    return term.replace('\\', '\\\\').replace('%', '\\%').replace('_', '\\_')


def _row_to_note(row: sqlite3.Row) -> Note:
    return Note(
        id=row['id'],
        plan_date=row['plan_date'],
        kind=row['kind'],
        text=row['text'],
        attachment_id=row['attachment_id'],
        source=row['source'],
        excel_column=row['excel_column'],
        created_at=row['created_at'],
        updated_at=row['updated_at'],
    )


def _kick_dispatcher() -> None:
    try:
        from shouyu.service import dispatch

        dispatch.kick()
    except Exception:
        logging.exception('failed to kick the Excel dispatcher')


# ---------- writes ----------

def save(
    text: Optional[str] = None,
    image_png: Optional[bytes] = None,
    excel_column: Optional[str] = None,
    source: str = SOURCE_CLIPBOARD,
) -> Note:
    """Store one record (text or a PNG image) and queue its Excel mirror,
    all in one transaction."""
    if image_png is None and not text:
        raise ValueError('nothing to save')
    column = excel_column if excel_column in ('A', 'B') else None
    kind = KIND_IMAGE if image_png is not None else KIND_TEXT
    plan_date = time.strftime('%Y-%m-%d')
    now = message_queue.utcnow_iso()
    enqueued = False
    with datastore.transaction() as conn:
        attachment_id = None
        if image_png is not None:
            attachment_id = message_queue.store_attachment(
                image_png, mime_type='image/png', filename='clipboard.png', conn=conn
            )
        note_text = '' if image_png is not None else text
        note_id = conn.execute(
            "INSERT INTO note (plan_date, kind, text, attachment_id, source, excel_column, created_at) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (plan_date, kind, note_text, attachment_id, source, column, now),
        ).lastrowid
        if Config.excel_mirror():
            payload = {'column': excel_column, 'note_id': note_id}
            if image_png is None:
                payload['text'] = text
            message_queue.enqueue('clipboard_append', payload, attachment_id=attachment_id, conn=conn)
            enqueued = True
    if enqueued:
        _kick_dispatcher()
    return Note(
        id=note_id,
        plan_date=plan_date,
        kind=kind,
        text=note_text,
        attachment_id=attachment_id,
        source=source,
        excel_column=column,
        created_at=now,
    )


def add_manual(plan_date: str, text: str) -> Note:
    """A text note typed in the day view for any date. Not mirrored to Excel."""
    text = (text or '').strip()
    if not text:
        raise ValueError('nothing to save')
    now = message_queue.utcnow_iso()
    with datastore.transaction() as conn:
        note_id = conn.execute(
            "INSERT INTO note (plan_date, kind, text, source, created_at) VALUES (?, ?, ?, ?, ?)",
            (plan_date, KIND_TEXT, text, SOURCE_MANUAL, now),
        ).lastrowid
    return Note(
        id=note_id,
        plan_date=plan_date,
        kind=KIND_TEXT,
        text=text,
        attachment_id=None,
        source=SOURCE_MANUAL,
        excel_column=None,
        created_at=now,
    )


def update_text(note_id: int, text: str) -> bool:
    """Change a text note's content. Image notes have no text to edit."""
    text = (text or '').strip()
    if not text:
        raise ValueError('text must not be empty')
    with datastore.transaction() as conn:
        cur = conn.execute(
            "UPDATE note SET text = ?, updated_at = ? WHERE id = ? AND kind = ? AND deleted_at IS NULL",
            (text, message_queue.utcnow_iso(), note_id, KIND_TEXT),
        )
        return cur.rowcount > 0


def delete(note_id: int) -> bool:
    """Soft delete: the row (and its attachment) stays, every read skips it."""
    with datastore.transaction() as conn:
        cur = conn.execute(
            "UPDATE note SET deleted_at = ? WHERE id = ? AND deleted_at IS NULL",
            (message_queue.utcnow_iso(), note_id),
        )
        return cur.rowcount > 0


# ---------- reads ----------

def last() -> Optional[Note]:
    with datastore.reader() as conn:
        row = conn.execute(
            f"SELECT {_COLUMNS} FROM note WHERE deleted_at IS NULL ORDER BY id DESC LIMIT 1"
        ).fetchone()
    return _row_to_note(row) if row is not None else None


def get(note_id: int) -> Optional[Note]:
    with datastore.reader() as conn:
        row = conn.execute(
            f"SELECT {_COLUMNS} FROM note WHERE id = ? AND deleted_at IS NULL", (note_id,)
        ).fetchone()
    return _row_to_note(row) if row is not None else None


def page_before(plan_date: str, note_id: int, limit: int = PAGE_SIZE) -> List[Note]:
    """Up to `limit` notes strictly before (plan_date, note_id), returned
    oldest first. Start with (tomorrow, 0) to get the newest page."""
    with datastore.reader() as conn:
        rows = conn.execute(
            f"SELECT {_COLUMNS} FROM note WHERE deleted_at IS NULL AND (plan_date, id) < (?, ?) "
            "ORDER BY plan_date DESC, id DESC LIMIT ?",
            (plan_date, note_id, limit),
        ).fetchall()
    return [_row_to_note(r) for r in reversed(rows)]


def list_day(plan_date: str) -> List[Note]:
    with datastore.reader() as conn:
        rows = conn.execute(
            f"SELECT {_COLUMNS} FROM note WHERE deleted_at IS NULL AND plan_date = ? ORDER BY id",
            (plan_date,),
        ).fetchall()
    return [_row_to_note(r) for r in rows]


def day_page(plan_date: str, before_id: Optional[int] = None, limit: int = PAGE_SIZE) -> List[Note]:
    """The newest `limit` notes of one day older than `before_id` (all when
    None), oldest first. Some imported days hold tens of thousands of cells."""
    with datastore.reader() as conn:
        rows = conn.execute(
            f"SELECT {_COLUMNS} FROM note WHERE deleted_at IS NULL AND plan_date = ? AND id < ? "
            "ORDER BY id DESC LIMIT ?",
            (plan_date, before_id if before_id is not None else 2 ** 62, limit),
        ).fetchall()
    return [_row_to_note(r) for r in reversed(rows)]


def count_day(plan_date: str) -> int:
    with datastore.reader() as conn:
        return conn.execute(
            "SELECT COUNT(*) FROM note WHERE deleted_at IS NULL AND plan_date = ?", (plan_date,)
        ).fetchone()[0]


def search(query: NoteQuery, conn: Optional[sqlite3.Connection] = None) -> List[Note]:
    """Notes matching every term within the date range, newest first.
    With no terms this is plain browsing (image notes included); with terms
    only text matches (image notes have no text)."""
    sql = [f"SELECT {_COLUMNS} FROM note WHERE deleted_at IS NULL"]
    args: list = []
    if query.date_from:
        sql.append("AND plan_date >= ?")
        args.append(query.date_from)
    if query.date_to:
        sql.append("AND plan_date <= ?")
        args.append(query.date_to)
    for term in query.terms:
        sql.append("AND text LIKE ? ESCAPE '\\'")
        args.append(f'%{_escape_like(term)}%')
    sql.append("ORDER BY plan_date DESC, id DESC LIMIT ?")
    args.append(max(1, int(query.limit)))
    statement = ' '.join(sql)
    if conn is not None:
        return [_row_to_note(r) for r in conn.execute(statement, args).fetchall()]
    with datastore.reader() as own:
        return [_row_to_note(r) for r in own.execute(statement, args).fetchall()]


def image_bytes(note: Note) -> Optional[bytes]:
    if note.attachment_id is None:
        return None
    return message_queue.get_attachment_bytes(note.attachment_id)
