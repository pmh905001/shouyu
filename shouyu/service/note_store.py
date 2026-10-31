"""Clipboard / screenshot / OCR records, stored in SQLite first.

Design reference: docs/design/planning/20261031-01-tasks-to-sqlite.md (§4.5).
A note row is the record; the Excel "other" area only receives a mirrored
copy via the existing `clipboard_append` queue message. Images live in the
queue's `attachments` table, which the note and the mirror message share.
"""
from __future__ import annotations

import logging
import time
from dataclasses import dataclass
from typing import Optional

from shouyu.config import Config
from shouyu.service import datastore, message_queue

KIND_TEXT = 'text'
KIND_IMAGE = 'image'
SOURCE_CLIPBOARD = 'clipboard'
SOURCE_OCR = 'ocr'


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
        try:
            from shouyu.service import dispatch

            dispatch.kick()
        except Exception:
            logging.exception('failed to kick the Excel dispatcher')
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


def last() -> Optional[Note]:
    with datastore.reader() as conn:
        row = conn.execute("SELECT * FROM note ORDER BY id DESC LIMIT 1").fetchone()
    if row is None:
        return None
    return Note(
        id=row['id'],
        plan_date=row['plan_date'],
        kind=row['kind'],
        text=row['text'],
        attachment_id=row['attachment_id'],
        source=row['source'],
        excel_column=row['excel_column'],
        created_at=row['created_at'],
    )


def image_bytes(note: Note) -> Optional[bytes]:
    if note.attachment_id is None:
        return None
    return message_queue.get_attachment_bytes(note.attachment_id)
