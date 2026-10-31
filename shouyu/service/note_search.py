"""Background note search for the notes panel.

Design reference: docs/design/retrieval/20261031-01-notes-panel-and-day-view.md
(§3.2). `LIKE '%term%'` over years of notes can take a few hundred ms for a
rare term, so searches run on one worker thread with its own read-only
connection:

1. The last three days of the requested range are searched first and
   reported right away (milliseconds, via the plan_date index).
2. The rest of the range is searched next and reported as an append.

A newer submit interrupts the running query (`sqlite3.Connection.interrupt`
is safe to call from another thread) and only the newest request is kept,
so typing never waits for an outdated search. Callbacks run on the worker
thread; the UI marshals them to its own thread and drops results whose
generation is not the latest.
"""
from __future__ import annotations

import logging
import sqlite3
import threading
import time
from dataclasses import replace
from datetime import datetime, timedelta
from typing import Callable, List, Optional, Tuple

from shouyu.service import datastore, note_store
from shouyu.service.note_store import Note, NoteQuery

NEAR_DAYS = 3

# (generation, notes, appended, finished)
SearchCallback = Callable[[int, List[Note], bool, bool], None]


def _shift(date_str: str, days: int) -> str:
    return (datetime.strptime(date_str, '%Y-%m-%d') + timedelta(days=days)).strftime('%Y-%m-%d')


def split_range(query: NoteQuery, today: str) -> List[NoteQuery]:
    """The query split into [near part, far part], skipping empty parts.
    "Near" is the last NEAR_DAYS days up to today."""
    near_from = _shift(today, -(NEAR_DAYS - 1))
    lo, hi = query.date_from, query.date_to
    parts: List[NoteQuery] = []
    # Near part: [max(lo, near_from), hi]
    n_lo = max(lo, near_from) if lo else near_from
    if hi is None or n_lo <= hi:
        parts.append(replace(query, date_from=n_lo, date_to=hi))
    # Far part: [lo, min(hi, near_from - 1)]
    f_hi = _shift(near_from, -1)
    if hi is not None and hi < f_hi:
        f_hi = hi
    if lo is None or lo <= f_hi:
        parts.append(replace(query, date_from=lo, date_to=f_hi))
    return parts


class NoteSearcher:
    def __init__(self, today_fn: Callable[[], str] = lambda: time.strftime('%Y-%m-%d')) -> None:
        self._today_fn = today_fn
        self._cond = threading.Condition()
        self._pending: Optional[Tuple[int, NoteQuery, SearchCallback]] = None
        self._running_generation: Optional[int] = None
        self._conn: Optional[sqlite3.Connection] = None
        self._thread: Optional[threading.Thread] = None
        self._closed = False

    def submit(self, generation: int, query: NoteQuery, callback: SearchCallback) -> None:
        with self._cond:
            if self._closed:
                return
            self._pending = (generation, query, callback)
            if self._running_generation is not None and self._conn is not None:
                try:
                    self._conn.interrupt()
                except Exception:
                    logging.exception('failed to interrupt the running note search')
            if self._thread is None:
                self._thread = threading.Thread(target=self._run, name='shouyu-note-search', daemon=True)
                self._thread.start()
            self._cond.notify()

    def close(self) -> None:
        with self._cond:
            self._closed = True
            self._pending = None
            if self._conn is not None and self._running_generation is not None:
                try:
                    self._conn.interrupt()
                except Exception:
                    pass
            self._cond.notify()

    def _has_newer(self, generation: int) -> bool:
        with self._cond:
            return self._closed or (self._pending is not None and self._pending[0] != generation)

    def _run(self) -> None:
        try:
            conn = datastore.open_connection()
        except Exception:
            logging.exception('note search: failed to open the database')
            with self._cond:
                self._thread = None
            return
        with self._cond:
            self._conn = conn
        try:
            while True:
                with self._cond:
                    while self._pending is None and not self._closed:
                        self._cond.wait()
                    if self._closed:
                        return
                    generation, query, callback = self._pending
                    self._pending = None
                    self._running_generation = generation
                try:
                    self._execute(conn, generation, query, callback)
                except sqlite3.OperationalError as e:
                    if 'interrupt' not in str(e).lower():
                        logging.exception('note search failed')
                except Exception:
                    logging.exception('note search failed')
                finally:
                    with self._cond:
                        self._running_generation = None
        finally:
            with self._cond:
                self._conn = None
            conn.close()

    def _execute(self, conn: sqlite3.Connection, generation: int, query: NoteQuery, callback: SearchCallback) -> None:
        parts = split_range(query, self._today_fn())
        if not parts:
            callback(generation, [], False, True)
            return
        remaining = query.limit
        for index, part in enumerate(parts):
            if self._has_newer(generation):
                return
            notes = note_store.search(replace(part, limit=remaining), conn=conn) if remaining > 0 else []
            remaining -= len(notes)
            finished = index == len(parts) - 1 or remaining <= 0
            callback(generation, notes, index > 0, finished)
            if finished:
                return
