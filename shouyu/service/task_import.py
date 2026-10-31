"""One-time import of the Excel task history into SQLite.

Design reference: docs/design/planning/20261031-01-tasks-to-sqlite.md (§4.7).
Parsing reuses `PlanService` / `BacklogService`, so imported tasks read
exactly the way the v1 UI shows them (font-color status, `50m` durations,
legacy placeholder rows...). The workbook is parsed completely before the
write transaction starts, so the SQLite write lock is held only for the
inserts themselves.

Every imported row records where it came from (`source_ref`, e.g.
`2026-10-30!B5`), with unique indexes + INSERT OR IGNORE, so running the
import again never duplicates anything.
"""
from __future__ import annotations

import logging
import os
import re
import threading
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from typing import List, Optional

from shouyu.config import Config
from shouyu.service import datastore, message_queue
from shouyu.service.plan import (
    BACKLOG_SHEET_LIFE,
    BACKLOG_SHEET_WORK,
    POMODORO_LOG_PREFIX,
    BacklogService,
    PlanService,
    PlanTask,
    TaskCategory,
)
from shouyu.service.task_store import KIND_WORK

_DATE_SHEET = re.compile(r'^\d{4}-\d{2}-\d{2}$')
_POMODORO_LINE = re.compile(r'🍅\s*(\d+)min\s*@(\d{1,2}:\d{2})-(\d{1,2}:\d{2})')

_running = threading.Event()
_lock = threading.Lock()


@dataclass
class _ParsedPomodoro:
    row: int
    task_text: str
    minutes: Optional[int]
    started_at: Optional[str]
    ended_at: Optional[str]


@dataclass
class _ParsedDay:
    date: str
    tasks: List[PlanTask]
    pomodoros: List[_ParsedPomodoro]


@dataclass
class ImportResult:
    days: int = 0
    tasks: int = 0
    backlog: int = 0
    pomodoros: int = 0
    skipped_reason: str = ''
    errors: List[str] = field(default_factory=list)


def _local_to_utc_iso(date_str: str, hhmm: str = '00:00', add_days: int = 0) -> str:
    local = datetime.strptime(f'{date_str} {hhmm}', '%Y-%m-%d %H:%M') + timedelta(days=add_days)
    # A naive datetime is interpreted as local time by astimezone().
    return local.astimezone(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3] + 'Z'


def _parse_pomodoros(date_str: str, plan: PlanService) -> List[_ParsedPomodoro]:
    """Same rows `PlanService.count_pomodoros_logged()` counts (first cell per
    row starting with 🍅, from the active area to the end of the sheet), so
    the imported count matches what v1 displayed for that day."""
    ws = plan.ws
    start = plan.active_area_start_row()
    end = max(ws.max_row, start)
    max_col = max(ws.max_column, 3)
    titles = {e.row: e.text for e in plan.list_active_entries()}
    title_rows = sorted(titles)
    out: List[_ParsedPomodoro] = []
    for offset, values in enumerate(
        ws.iter_rows(min_row=start, max_row=end, max_col=max_col, values_only=True)
    ):
        row = start + offset
        line = next(
            (v for v in values if isinstance(v, str) and v.strip().startswith(POMODORO_LOG_PREFIX)),
            None,
        )
        if line is None:
            continue
        owner = [r for r in title_rows if r < row]
        task_text = titles[owner[-1]] if owner else ''
        minutes = started = ended = None
        m = _POMODORO_LINE.search(line)
        if m:
            minutes = int(m.group(1))
            try:
                started = _local_to_utc_iso(date_str, m.group(2))
                crosses_midnight = m.group(3).zfill(5) < m.group(2).zfill(5)
                ended = _local_to_utc_iso(date_str, m.group(3), add_days=1 if crosses_midnight else 0)
            except ValueError:
                started = ended = None
        out.append(_ParsedPomodoro(row=row, task_text=task_text, minutes=minutes, started_at=started, ended_at=ended))
    return out


def _parse_workbook(path: str):
    import openpyxl

    workbook = openpyxl.load_workbook(path)
    days: List[_ParsedDay] = []
    for name in sorted(n for n in workbook.sheetnames if _DATE_SHEET.match(n)):
        plan = PlanService(workbook[name])
        # In-memory only (the workbook is never saved): drop the old
        # task 1/2/3 placeholders exactly like opening the day in v1 does.
        plan.remove_legacy_default_tasks()
        days.append(_ParsedDay(date=name, tasks=plan.read_plan_tasks(), pomodoros=_parse_pomodoros(name, plan)))
    backlog = []
    for sheet, category in ((BACKLOG_SHEET_WORK, TaskCategory.WORK), (BACKLOG_SHEET_LIFE, TaskCategory.LIFE)):
        if sheet in workbook.sheetnames:
            backlog.append((sheet, category, BacklogService(workbook[sheet], category).read()))
    return days, backlog


def _insert_task(conn, task: PlanTask, source_ref: str, created_at: str) -> int:
    conn.execute(
        "INSERT OR IGNORE INTO task (text, category, priority, duration_minutes, reflection, "
        "created_at, updated_at, source, source_ref) VALUES (?, ?, ?, ?, ?, ?, ?, 'excel_import', ?)",
        (
            task.text,
            task.category.value,
            task.priority.value,
            max(0, int(task.duration_minutes or 0)),
            task.reflection or '',
            created_at,
            created_at,
            source_ref,
        ),
    )
    return conn.execute(
        "SELECT id FROM task WHERE source = 'excel_import' AND source_ref = ?", (source_ref,)
    ).fetchone()['id']


def run_import(excel_path: Optional[str] = None) -> ImportResult:
    """Import every dated sheet, both backlog sheets and the 🍅 lines.
    Marks the import finished in `meta` only when it succeeds."""
    path = excel_path or Config.excel_path()
    result = ImportResult()
    if not path or not os.path.exists(path):
        result.skipped_reason = f'Excel file not found: {path}'
        logging.warning(f'task import skipped: {result.skipped_reason}')
        return result

    days, backlog = _parse_workbook(path)
    with datastore.transaction() as conn:
        for day in days:
            created_at = _local_to_utc_iso(day.date)
            ids_by_text = {}
            for position, task in enumerate(day.tasks):
                task_id = _insert_task(conn, task, f'{day.date}!B{task.row}', created_at)
                conn.execute(
                    "INSERT OR IGNORE INTO day_plan_entry (plan_date, task_id, position, status, created_at, updated_at) "
                    "VALUES (?, ?, ?, ?, ?, ?)",
                    (day.date, task_id, position, task.status.value, created_at, created_at),
                )
                ids_by_text.setdefault(task.text.strip(), task_id)
                result.tasks += 1
            for p in day.pomodoros:
                conn.execute(
                    "INSERT OR IGNORE INTO pomodoro (task_id, task_text, plan_date, kind, started_at, ended_at, "
                    "actual_minutes, source, source_ref) VALUES (?, ?, ?, ?, ?, ?, ?, 'excel_import', ?)",
                    (
                        ids_by_text.get(p.task_text.strip()),
                        p.task_text,
                        day.date,
                        KIND_WORK,
                        p.started_at,
                        p.ended_at,
                        p.minutes,
                        f'{day.date}!R{p.row}',
                    ),
                )
                result.pomodoros += 1
            result.days += 1

        now = message_queue.utcnow_iso()
        today = datetime.now().strftime('%Y-%m-%d')
        for sheet, category, tasks in backlog:
            for position, task in enumerate(tasks):
                task.category = category
                task_id = _insert_task(conn, task, f'{sheet}!A{task.row}', now)
                conn.execute(
                    "INSERT OR IGNORE INTO backlog_entry (task_id, position, entered_on, created_at) "
                    "VALUES (?, ?, ?, ?)",
                    (task_id, position, (task.created_date or '').strip() or today, now),
                )
                result.backlog += 1

        datastore.set_meta(conn, datastore.META_IMPORT_FINISHED_AT, now)
        datastore.set_meta(conn, datastore.META_IMPORT_SOURCE, path)
    logging.info(
        f'task import finished from {path}: {result.days} days, {result.tasks} tasks, '
        f'{result.backlog} backlog, {result.pomodoros} pomodoros'
    )
    return result


def needs_import() -> bool:
    with datastore.reader() as conn:
        return datastore.get_meta(conn, datastore.META_IMPORT_FINISHED_AT) is None


def is_running() -> bool:
    return _running.is_set()


def start_in_background() -> bool:
    """Start the import on a daemon thread if it has never completed.
    Returns True when an import was started."""
    with _lock:
        if _running.is_set():
            return False
        try:
            if not needs_import():
                return False
        except Exception:
            logging.exception('failed to check whether the task import is needed')
            return False
        _running.set()

    def _work() -> None:
        try:
            run_import()
        except Exception:
            logging.exception('task import from Excel failed; it will be retried on next startup')
        finally:
            _running.clear()

    threading.Thread(target=_work, name='shouyu-task-import', daemon=True).start()
    return True
