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
_notes_running = threading.Event()
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


def _load_workbook(path: str):
    import openpyxl

    return openpyxl.load_workbook(path)


def _parse_workbook(path: str, workbook=None):
    workbook = workbook or _load_workbook(path)
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


def run_import(excel_path: Optional[str] = None, workbook=None) -> ImportResult:
    """Import every dated sheet, both backlog sheets and the 🍅 lines.
    Marks the import finished in `meta` only when it succeeds."""
    path = excel_path or Config.excel_path()
    result = ImportResult()
    if not path or not os.path.exists(path):
        result.skipped_reason = f'Excel file not found: {path}'
        logging.warning(f'task import skipped: {result.skipped_reason}')
        return result

    days, backlog = _parse_workbook(path, workbook)
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


@dataclass
class _ParsedNote:
    date: str
    cell: str
    column: str
    text: str


@dataclass
class NotesImportResult:
    notes: int = 0
    skipped_pomodoro_lines: int = 0
    skipped_duplicates: int = 0
    sheets_without_header: List[str] = field(default_factory=list)
    skipped_reason: str = ''


def _cell_text(value) -> str:
    if value is None:
        return ''
    if hasattr(value, 'strftime'):
        try:
            return value.strftime('%Y-%m-%d %H:%M' if getattr(value, 'hour', 0) or getattr(value, 'minute', 0) else '%Y-%m-%d')
        except Exception:
            return str(value)
    return str(value)


def _parse_notes(workbook, result: NotesImportResult) -> List[_ParsedNote]:
    """Every non-empty cell below the `other` header, row by row and left to
    right in a row: one cell = one note (R-Q10). Images stay in Excel (R-Q5).
    🍅 lines that ended up below the header are pomodoro logs, already
    imported as pomodoros, so they are not notes."""
    from openpyxl.utils import get_column_letter

    notes: List[_ParsedNote] = []
    for name in sorted(n for n in workbook.sheetnames if _DATE_SHEET.match(n)):
        ws = workbook[name]
        header = PlanService(ws)._other_header_row()
        if not header:
            result.sheets_without_header.append(name)
            continue
        for row_offset, values in enumerate(
            ws.iter_rows(min_row=header + 1, max_row=max(ws.max_row, header + 1), values_only=True)
        ):
            for col_offset, value in enumerate(values):
                text = _cell_text(value)
                if not text.strip():
                    continue
                if text.strip().startswith(POMODORO_LOG_PREFIX):
                    result.skipped_pomodoro_lines += 1
                    continue
                column = get_column_letter(col_offset + 1)
                notes.append(_ParsedNote(
                    date=name, cell=f'{column}{header + 1 + row_offset}', column=column, text=text,
                ))
    return notes


def _switch_day(conn) -> Optional[str]:
    """Local date of the day v2 took over (the task import finished). Excel's
    `other` area of that day holds both v1's earlier notes and v2's mirrored
    ones."""
    finished = datastore.get_meta(conn, datastore.META_IMPORT_FINISHED_AT)
    if not finished:
        return None
    try:
        moment = datetime.strptime(finished, '%Y-%m-%dT%H:%M:%S.%fZ').replace(tzinfo=timezone.utc)
    except ValueError:
        return None
    return moment.astimezone().strftime('%Y-%m-%d')


def run_notes_import(excel_path: Optional[str] = None, workbook=None) -> NotesImportResult:
    """Import the historical text of the `other` areas as notes (REC-18,
    design §4.7). Idempotent through `source_ref`; never queues Excel sync
    messages, the text is already in Excel."""
    path = excel_path or Config.excel_path()
    result = NotesImportResult()
    if not path or not os.path.exists(path):
        result.skipped_reason = f'Excel file not found: {path}'
        logging.warning(f'notes import skipped: {result.skipped_reason}')
        return result

    parsed = _parse_notes(workbook or _load_workbook(path), result)
    with datastore.transaction() as conn:
        switch_day = _switch_day(conn)
        existing = set()
        if switch_day:
            existing = {
                r['text'].strip() for r in conn.execute(
                    "SELECT text FROM note WHERE plan_date = ? AND kind = 'text' AND source != 'excel_import' "
                    "AND deleted_at IS NULL", (switch_day,),
                )
            }
        created_at = {}
        for note in parsed:
            if note.date == switch_day and note.text.strip() in existing:
                result.skipped_duplicates += 1
                continue
            if note.date not in created_at:
                created_at[note.date] = _local_to_utc_iso(note.date)
            cursor = conn.execute(
                "INSERT OR IGNORE INTO note (plan_date, kind, text, source, source_ref, excel_column, created_at) "
                "VALUES (?, 'text', ?, 'excel_import', ?, ?, ?)",
                (note.date, note.text, f'{note.date}!{note.cell}',
                 note.column if note.column in ('A', 'B') else None, created_at[note.date]),
            )
            result.notes += cursor.rowcount
        datastore.set_meta(conn, datastore.META_NOTES_IMPORT_FINISHED_AT, message_queue.utcnow_iso())
    logging.info(
        f'notes import finished from {path}: {result.notes} notes, '
        f'{result.skipped_duplicates} already in v2, {result.skipped_pomodoro_lines} 🍅 lines skipped; '
        f'sheets without an "other" header: {", ".join(result.sheets_without_header) or "none"}'
    )
    return result


def needs_import() -> bool:
    with datastore.reader() as conn:
        return datastore.get_meta(conn, datastore.META_IMPORT_FINISHED_AT) is None


def needs_notes_import() -> bool:
    with datastore.reader() as conn:
        return datastore.get_meta(conn, datastore.META_NOTES_IMPORT_FINISHED_AT) is None


def is_running() -> bool:
    """True while the *task* import runs (the habit dialog waits for it).
    The notes import has its own flag and never blocks the dialog."""
    return _running.is_set()


def is_notes_import_running() -> bool:
    return _notes_running.is_set()


def start_in_background() -> bool:
    """Start whichever imports have never completed (tasks first, then the
    notes) on one daemon thread; the workbook is loaded once. Returns True
    when something was started."""
    with _lock:
        if _running.is_set() or _notes_running.is_set():
            return False
        try:
            tasks_needed = needs_import()
            notes_needed = needs_notes_import()
        except Exception:
            logging.exception('failed to check whether the Excel import is needed')
            return False
        if not (tasks_needed or notes_needed):
            return False
        if tasks_needed:
            _running.set()
        if notes_needed:
            _notes_running.set()

    def _work() -> None:
        workbook = None
        path = Config.excel_path()
        try:
            if path and os.path.exists(path):
                workbook = _load_workbook(path)
        except Exception:
            logging.exception('failed to open the Excel file for import; it will be retried on next startup')
        try:
            if tasks_needed:
                try:
                    run_import(path, workbook)
                except Exception:
                    logging.exception('task import from Excel failed; it will be retried on next startup')
                finally:
                    _running.clear()
            if notes_needed:
                try:
                    run_notes_import(path, workbook)
                except Exception:
                    logging.exception('notes import from Excel failed; it will be retried on next startup')
        finally:
            _running.clear()
            _notes_running.clear()

    threading.Thread(target=_work, name='shouyu-task-import', daemon=True).start()
    return True
