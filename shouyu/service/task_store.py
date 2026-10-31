"""Today's plan, the work / life backlog and pomodoro records, stored in SQLite.

Design reference: docs/design/planning/20261031-01-tasks-to-sqlite.md (§4).
SQLite is the only source of truth: the habit dialog, the todo panel and the
pomodoro service read and write through this module, and Excel only gets a
mirrored copy through the durable queue (never read back).

The UI keeps working the way it always has: load once, edit the in-memory
`PlanTask` lists (with undo / redo), then hand the whole lists back to
`save_plan()`, which works out the differences and writes them in a single
transaction together with the Excel mirror messages.
"""
from __future__ import annotations

import logging
import sqlite3
import time
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from typing import Dict, Iterable, List, Optional, Set

from shouyu.config import Config
from shouyu.service import datastore, message_queue
from shouyu.service.plan import PlanTask, TaskCategory, TaskPriority, TaskStatus

HISTORY_LOOKBACK_DAYS = 30
KIND_WORK = 'work'
KIND_PLANNING = 'planning'

_TASK_COLUMNS = (
    "t.id AS task_id, t.text, t.category, t.priority, t.duration_minutes, t.reflection"
)


@dataclass
class DaySnapshot:
    date: str
    unfinished: List[PlanTask] = field(default_factory=list)
    done: int = 0
    total: int = 0
    pomodoros: int = 0

    def as_dict(self) -> dict:
        return {
            "date": self.date,
            "unfinished": list(self.unfinished),
            "done": self.done,
            "total": self.total,
            "pomodoros": self.pomodoros,
        }


@dataclass
class DialogData:
    today_tasks: List[PlanTask]
    work_tasks: List[PlanTask]
    life_tasks: List[PlanTask]
    today_pomodoros: int
    # Latest earlier day (within HISTORY_LOOKBACK_DAYS) that has tasks, or
    # yesterday when there is none.
    history: DaySnapshot


@dataclass
class PlanSaveRequest:
    # Local date (YYYY-MM-DD) the window was showing when it was opened.
    plan_date: str
    today_tasks: List[PlanTask]
    original_in_progress: Optional[str] = None
    # None means "this window doesn't edit that backlog section": leave it as is.
    work_tasks: Optional[List[PlanTask]] = None
    life_tasks: Optional[List[PlanTask]] = None
    # Earlier-day tasks the user ticked as done from the carry-over card.
    history_date: Optional[str] = None
    history_done_task_ids: List[int] = field(default_factory=list)


@dataclass
class SaveResult:
    changed: bool


@dataclass
class PomodoroRecord:
    plan_date: str
    task_text: str = ''
    kind: str = KIND_WORK
    mode: Optional[str] = None
    started_at: Optional[float] = None  # epoch seconds
    ended_at: Optional[float] = None
    planned_minutes: Optional[int] = None
    actual_minutes: Optional[int] = None
    ended_early: bool = False


# ---------- helpers ----------

def _epoch_to_iso(ts: float) -> str:
    return datetime.fromtimestamp(ts, timezone.utc).strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3] + 'Z'


def _days_before(date_str: str, days: int) -> str:
    return (datetime.strptime(date_str, '%Y-%m-%d') - timedelta(days=days)).strftime('%Y-%m-%d')


def _non_empty(tasks: Optional[Iterable[PlanTask]]) -> List[PlanTask]:
    return [t for t in (tasks or []) if (t.text or '').strip()]


def _row_to_task(row: sqlite3.Row, status: TaskStatus = TaskStatus.PENDING, created_date: str = '') -> PlanTask:
    return PlanTask(
        text=row['text'],
        status=status,
        duration_minutes=int(row['duration_minutes'] or 0),
        priority=TaskPriority.from_value(row['priority']),
        reflection=row['reflection'] or '',
        category=TaskCategory.from_value(row['category']),
        created_date=created_date,
        task_id=int(row['task_id']),
    )


def _read_day_tasks(conn: sqlite3.Connection, plan_date: str) -> List[PlanTask]:
    rows = conn.execute(
        f"SELECT {_TASK_COLUMNS}, e.status FROM day_plan_entry e "
        "JOIN task t ON t.id = e.task_id WHERE e.plan_date = ? ORDER BY e.position, e.id",
        (plan_date,),
    ).fetchall()
    return [_row_to_task(r, status=TaskStatus(r['status'])) for r in rows]


def _read_backlog(conn: sqlite3.Connection, category: TaskCategory) -> List[PlanTask]:
    rows = conn.execute(
        f"SELECT {_TASK_COLUMNS}, b.entered_on FROM backlog_entry b "
        "JOIN task t ON t.id = b.task_id WHERE t.category = ? ORDER BY b.position, b.task_id",
        (category.value,),
    ).fetchall()
    return [_row_to_task(r, created_date=r['entered_on'] or '') for r in rows]


def _count_pomodoros(conn: sqlite3.Connection, plan_date: str) -> int:
    return conn.execute(
        "SELECT COUNT(*) FROM pomodoro WHERE plan_date = ? AND kind = ?",
        (plan_date, KIND_WORK),
    ).fetchone()[0]


def _read_day(conn: sqlite3.Connection, plan_date: str) -> DaySnapshot:
    tasks = _read_day_tasks(conn, plan_date)
    unfinished = []
    for t in tasks:
        if t.status == TaskStatus.DONE:
            continue
        t.status = TaskStatus.PENDING
        unfinished.append(t)
    return DaySnapshot(
        date=plan_date,
        unfinished=unfinished,
        done=sum(1 for t in tasks if t.status == TaskStatus.DONE),
        total=len(tasks),
        pomodoros=_count_pomodoros(conn, plan_date),
    )


def _latest_history_date(conn: sqlite3.Connection, today: str) -> Optional[str]:
    row = conn.execute(
        "SELECT MAX(plan_date) FROM day_plan_entry WHERE plan_date < ? AND plan_date >= ?",
        (today, _days_before(today, HISTORY_LOOKBACK_DAYS)),
    ).fetchone()
    return row[0] if row and row[0] else None


def _in_progress_row(conn: sqlite3.Connection, plan_date: str) -> Optional[sqlite3.Row]:
    return conn.execute(
        f"SELECT {_TASK_COLUMNS} FROM day_plan_entry e JOIN task t ON t.id = e.task_id "
        "WHERE e.plan_date = ? AND e.status = ? ORDER BY e.position LIMIT 1",
        (plan_date, TaskStatus.IN_PROGRESS.value),
    ).fetchone()


def _log_status(conn, task_id: int, plan_date: str, from_status: Optional[str], to_status: str, now: str) -> None:
    conn.execute(
        "INSERT INTO task_status_log (task_id, plan_date, from_status, to_status, changed_at) "
        "VALUES (?, ?, ?, ?, ?)",
        (task_id, plan_date, from_status, to_status, now),
    )


def _kick_dispatcher() -> None:
    try:
        from shouyu.service import dispatch

        dispatch.kick()
    except Exception:
        logging.exception('failed to kick the Excel dispatcher')


# ---------- reads ----------

def load_dialog(today: str) -> DialogData:
    with datastore.reader() as conn:
        history_date = _latest_history_date(conn, today) or _days_before(today, 1)
        return DialogData(
            today_tasks=_read_day_tasks(conn, today),
            work_tasks=_read_backlog(conn, TaskCategory.WORK),
            life_tasks=_read_backlog(conn, TaskCategory.LIFE),
            today_pomodoros=_count_pomodoros(conn, today),
            history=_read_day(conn, history_date),
        )


def load_today(today: str) -> List[PlanTask]:
    with datastore.reader() as conn:
        return _read_day_tasks(conn, today)


def load_day(plan_date: str) -> DaySnapshot:
    with datastore.reader() as conn:
        return _read_day(conn, plan_date)


def count_pomodoros(plan_date: str) -> int:
    with datastore.reader() as conn:
        return _count_pomodoros(conn, plan_date)


def current_in_progress(today: str) -> Optional[PlanTask]:
    with datastore.reader() as conn:
        row = _in_progress_row(conn, today)
    return _row_to_task(row, status=TaskStatus.IN_PROGRESS) if row else None


# ---------- save ----------

class _Saver:
    """One `save_plan()` call: diff the UI's lists against the database."""

    def __init__(self, conn: sqlite3.Connection, req: PlanSaveRequest) -> None:
        self.conn = conn
        self.req = req
        self.now = message_queue.utcnow_iso()
        self.changed = False
        self.claimed: Set[int] = set()

        self.prev_today: Dict[int, sqlite3.Row] = {
            r['task_id']: r
            for r in conn.execute(
                "SELECT id, task_id, position, status FROM day_plan_entry WHERE plan_date = ?",
                (req.plan_date,),
            )
        }
        self.prev_backlog: Dict[int, sqlite3.Row] = {
            r['task_id']: r
            for r in conn.execute(
                "SELECT b.task_id, b.position, b.entered_on, t.category "
                "FROM backlog_entry b JOIN task t ON t.id = b.task_id"
            )
        }

    def run(self) -> None:
        req = self.req
        today = _non_empty(req.today_tasks)
        rewrite_backlog = req.work_tasks is not None or req.life_tasks is not None
        work = _non_empty(req.work_tasks) if req.work_tasks is not None else None
        life = _non_empty(req.life_tasks) if req.life_tasks is not None else None
        all_tasks = today + (work or []) + (life or [])

        self._build_reuse_pool(all_tasks)
        for task in all_tasks:
            self._upsert_task(task)

        self._save_today(today)
        backlog_ids = self._save_backlog(work, life) if rewrite_backlog else set(self.prev_backlog)
        self._soft_delete_dropped({t.task_id for t in today} | backlog_ids)
        self.done_texts = self._mark_history_done()
        self._enqueue_mirror(today, work, life)

    # A task saved earlier in this session may come back with task_id == 0
    # (e.g. restored from an undo snapshot taken before that save). Match it
    # by text to the same row it already had today / in the backlog instead
    # of inserting a duplicate.
    def _build_reuse_pool(self, all_tasks: List[PlanTask]) -> None:
        explicit = {t.task_id for t in all_tasks if t.task_id}
        candidates = sorted(self.prev_today.items(), key=lambda kv: kv[1]['position'])
        ids = [tid for tid, _ in candidates] + list(self.prev_backlog)
        self.reuse_pool: Dict[str, List[int]] = {}
        if not ids:
            return
        placeholders = ','.join('?' * len(ids))
        texts = {
            r['id']: r['text']
            for r in self.conn.execute(f"SELECT id, text FROM task WHERE id IN ({placeholders})", ids)
        }
        for tid in ids:
            if tid in explicit or tid not in texts:
                continue
            self.reuse_pool.setdefault(texts[tid].strip(), []).append(tid)

    def _take_reusable(self, text: str) -> int:
        pool = self.reuse_pool.get(text.strip()) or []
        while pool:
            tid = pool.pop(0)
            if tid not in self.claimed:
                return tid
        return 0

    def _upsert_task(self, task: PlanTask) -> None:
        conn = self.conn
        row = None
        tid = task.task_id if task.task_id and task.task_id not in self.claimed else 0
        if tid:
            row = conn.execute("SELECT * FROM task WHERE id = ?", (tid,)).fetchone()
        if row is None:
            tid = self._take_reusable(task.text)
            if tid:
                row = conn.execute("SELECT * FROM task WHERE id = ?", (tid,)).fetchone()
        values = (
            task.text,
            task.category.value,
            task.priority.value,
            max(0, int(task.duration_minutes or 0)),
            task.reflection or '',
        )
        if row is None:
            tid = conn.execute(
                "INSERT INTO task (text, category, priority, duration_minutes, reflection, created_at, updated_at) "
                "VALUES (?, ?, ?, ?, ?, ?, ?)",
                values + (self.now, self.now),
            ).lastrowid
            self.changed = True
        else:
            current = (row['text'], row['category'], row['priority'], row['duration_minutes'], row['reflection'])
            if current != values or row['deleted_at'] is not None:
                conn.execute(
                    "UPDATE task SET text = ?, category = ?, priority = ?, duration_minutes = ?, "
                    "reflection = ?, deleted_at = NULL, updated_at = ? WHERE id = ?",
                    values + (self.now, tid),
                )
                self.changed = True
        self.claimed.add(tid)
        task.task_id = tid

    def _save_today(self, today: List[PlanTask]) -> None:
        conn, plan_date = self.conn, self.req.plan_date
        keep = {t.task_id for t in today}
        for tid, prev in self.prev_today.items():
            if tid not in keep:
                conn.execute("DELETE FROM day_plan_entry WHERE id = ?", (prev['id'],))
                self.changed = True
        for position, task in enumerate(today):
            status = task.status.value
            prev = self.prev_today.get(task.task_id)
            if prev is None:
                conn.execute(
                    "INSERT INTO day_plan_entry (plan_date, task_id, position, status, created_at, updated_at) "
                    "VALUES (?, ?, ?, ?, ?, ?)",
                    (plan_date, task.task_id, position, status, self.now, self.now),
                )
                _log_status(conn, task.task_id, plan_date, None, status, self.now)
                self.changed = True
                continue
            if prev['position'] != position or prev['status'] != status:
                conn.execute(
                    "UPDATE day_plan_entry SET position = ?, status = ?, updated_at = ? WHERE id = ?",
                    (position, status, self.now, prev['id']),
                )
                self.changed = True
            if prev['status'] != status:
                _log_status(conn, task.task_id, plan_date, prev['status'], status, self.now)

    def _save_backlog(self, work: Optional[List[PlanTask]], life: Optional[List[PlanTask]]) -> Set[int]:
        conn = self.conn
        today_local = time.strftime('%Y-%m-%d')
        desired: Dict[int, tuple] = {}
        for category, tasks in ((TaskCategory.WORK, work), (TaskCategory.LIFE, life)):
            if tasks is None:
                # Section not edited by this window: keep it exactly as it was.
                for tid, prev in self.prev_backlog.items():
                    if prev['category'] == category.value and tid not in desired:
                        desired[tid] = (prev['position'], prev['entered_on'])
                continue
            for position, task in enumerate(tasks):
                if task.task_id in desired:
                    continue
                prev = self.prev_backlog.get(task.task_id)
                entered_on = (task.created_date or '').strip() or (prev['entered_on'] if prev else '') or today_local
                task.created_date = entered_on
                desired[task.task_id] = (position, entered_on)

        for tid in self.prev_backlog:
            if tid not in desired:
                conn.execute("DELETE FROM backlog_entry WHERE task_id = ?", (tid,))
                self.changed = True
        for tid, (position, entered_on) in desired.items():
            prev = self.prev_backlog.get(tid)
            if prev is None:
                conn.execute(
                    "INSERT INTO backlog_entry (task_id, position, entered_on, created_at) VALUES (?, ?, ?, ?)",
                    (tid, position, entered_on, self.now),
                )
                self.changed = True
            elif prev['position'] != position or prev['entered_on'] != entered_on:
                conn.execute(
                    "UPDATE backlog_entry SET position = ?, entered_on = ? WHERE task_id = ?",
                    (position, entered_on, tid),
                )
                self.changed = True
        return set(desired)

    def _soft_delete_dropped(self, still_placed: Set[int]) -> None:
        """A task that was on today's plan or in the backlog before this save
        and is now nowhere is a deletion - unless it still appears on another
        day, where it is history and must stay readable."""
        conn = self.conn
        for tid in (set(self.prev_today) | set(self.prev_backlog)) - still_placed:
            elsewhere = conn.execute(
                "SELECT 1 FROM day_plan_entry WHERE task_id = ? LIMIT 1", (tid,)
            ).fetchone()
            if elsewhere is None:
                conn.execute(
                    "UPDATE task SET deleted_at = ?, updated_at = ? WHERE id = ? AND deleted_at IS NULL",
                    (self.now, self.now, tid),
                )
                self.changed = True

    def _mark_history_done(self) -> List[str]:
        req, conn = self.req, self.conn
        texts: List[str] = []
        if not req.history_date or not req.history_done_task_ids:
            return texts
        for tid in dict.fromkeys(req.history_done_task_ids):
            row = conn.execute(
                "SELECT e.id, e.status, t.text FROM day_plan_entry e JOIN task t ON t.id = e.task_id "
                "WHERE e.plan_date = ? AND e.task_id = ?",
                (req.history_date, tid),
            ).fetchone()
            if row is None or row['status'] == TaskStatus.DONE.value:
                continue
            conn.execute(
                "UPDATE day_plan_entry SET status = ?, updated_at = ? WHERE id = ?",
                (TaskStatus.DONE.value, self.now, row['id']),
            )
            _log_status(conn, tid, req.history_date, row['status'], TaskStatus.DONE.value, self.now)
            texts.append(row['text'])
            self.changed = True
        return texts

    def _enqueue_mirror(self, today, work, life) -> None:
        self.enqueued = False
        if not self.changed or not Config.excel_mirror():
            return
        req = self.req
        message_queue.enqueue(
            'plan_save',
            {
                'date': req.plan_date,
                'tasks': [t.to_dict() for t in today],
                'original_in_progress': req.original_in_progress,
                'work_tasks': [t.to_dict() for t in work] if work is not None else None,
                'life_tasks': [t.to_dict() for t in life] if life is not None else None,
            },
            conn=self.conn,
        )
        if self.done_texts:
            message_queue.enqueue(
                'plan_mark_done_by_text',
                {'date': req.history_date, 'texts': self.done_texts},
                conn=self.conn,
            )
        self.enqueued = True


def save_plan(req: PlanSaveRequest) -> SaveResult:
    """Persist the UI's lists. Mutates the given PlanTask objects in place:
    new tasks get their `task_id`, backlog tasks get their `created_date`."""
    with datastore.transaction() as conn:
        saver = _Saver(conn, req)
        saver.run()
    if saver.enqueued:
        _kick_dispatcher()
    return SaveResult(changed=saver.changed)


# ---------- pomodoro ----------

def pomodoro_label(record: PomodoroRecord) -> str:
    """Same text v1 wrote into Excel: `🍅 50min @09:00-09:50`."""
    started = time.strftime('%H:%M', time.localtime(record.started_at)) if record.started_at else '--:--'
    ended = time.strftime('%H:%M', time.localtime(record.ended_at)) if record.ended_at else '--:--'
    return f"🍅 {record.actual_minutes or 0}min @{started}-{ended}"


def record_pomodoro(record: PomodoroRecord) -> int:
    """Insert one pomodoro row (and its Excel mirror message for work
    pomodoros). Attributed to today's in-progress task when its text matches
    the task the timer was showing. Returns the new row id."""
    enqueued = False
    with datastore.transaction() as conn:
        task_id = None
        wanted = (record.task_text or '').strip()
        row = _in_progress_row(conn, record.plan_date)
        if row is not None and wanted and row['text'].strip() == wanted:
            task_id = row['task_id']
        pomodoro_id = conn.execute(
            "INSERT INTO pomodoro (task_id, task_text, plan_date, kind, mode, started_at, ended_at, "
            "planned_minutes, actual_minutes, ended_early) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (
                task_id,
                record.task_text or '',
                record.plan_date,
                record.kind,
                record.mode,
                _epoch_to_iso(record.started_at) if record.started_at else None,
                _epoch_to_iso(record.ended_at) if record.ended_at else None,
                record.planned_minutes,
                record.actual_minutes,
                1 if record.ended_early else 0,
            ),
        ).lastrowid
        if record.kind == KIND_WORK and Config.excel_mirror():
            message_queue.enqueue(
                'pomodoro_log',
                {'date': record.plan_date, 'label': pomodoro_label(record), 'pomodoro_id': pomodoro_id},
                conn=conn,
            )
            enqueued = True
    if enqueued:
        _kick_dispatcher()
    return pomodoro_id
