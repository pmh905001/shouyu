"""Import of the historical `other` area text into SQLite notes (REC-18,
design §4.7)."""
import time

import openpyxl

from conftest import queue_rows
from shouyu.service import datastore, note_store, task_import
from shouyu.service.plan import OTHER_HEADER_TEXT, PlanService, PlanTask

DAY1, DAY2, NO_HEADER, SWITCH = '2026-10-28', '2026-10-29', '2026-10-25', '2026-10-31'


def _build_workbook(path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet(DAY1)
    PlanService(ws).write_plan_tasks([PlanTask('写周报')])
    header = PlanService(ws).ensure_other_header()
    ws.cell(row=header + 1, column=2, value='第一条 B')
    ws.cell(row=header + 2, column=1, value='A 列的')
    ws.cell(row=header + 2, column=2, value='同一行 B 列的')
    ws.cell(row=header + 3, column=2, value='🍅 50min @17:12-18:02')  # pomodoro log, not a note
    ws.cell(row=header + 4, column=2, value='   ')  # blank
    ws.cell(row=header + 5, column=3, value='C 列手写的')
    ws.cell(row=header + 6, column=2, value=42)

    ws2 = wb.create_sheet(DAY2)
    header2 = PlanService(ws2).ensure_other_header()
    ws2.cell(row=header2 + 1, column=2, value='多行\n文字')

    # A sheet without the `other` header: skipped and reported.
    wb.create_sheet(NO_HEADER)['B2'] = '计划里的任务'

    ws3 = wb.create_sheet(SWITCH)
    header3 = PlanService(ws3).ensure_other_header()
    ws3.cell(row=header3 + 1, column=2, value='v1 早上存的')
    ws3.cell(row=header3 + 2, column=2, value='v2 同步过去的')

    wb.create_sheet('accounts')['A5'] = OTHER_HEADER_TEXT
    wb.save(path)


def _set_switch_day(conn, local_date):
    utc = task_import._local_to_utc_iso(local_date, '12:00')
    datastore.set_meta(conn, datastore.META_IMPORT_FINISHED_AT, utc)


def test_import_other_area_text(db, v2env):
    _build_workbook(v2env['excel_path'])
    _set_switch_day(db, SWITCH)
    db.execute(
        "INSERT INTO note (plan_date, kind, text, source, excel_column, created_at) "
        "VALUES (?, 'text', 'v2 同步过去的', 'clipboard', 'B', '2026-10-31T02:00:00.000Z')", (SWITCH,),
    )
    assert task_import.needs_notes_import()

    result = task_import.run_notes_import()
    assert result.notes == 7
    assert result.skipped_pomodoro_lines == 1
    assert result.skipped_duplicates == 1
    assert result.sheets_without_header == [NO_HEADER]
    assert not task_import.needs_notes_import()

    day1 = note_store.list_day(DAY1)
    rows = db.execute(
        "SELECT text, source_ref, excel_column FROM note WHERE plan_date = ? ORDER BY id", (DAY1,)
    ).fetchall()
    assert [tuple(r) for r in rows] == [
        ('第一条 B', f'{DAY1}!B{_header(v2env, DAY1) + 1}', 'B'),
        ('A 列的', f'{DAY1}!A{_header(v2env, DAY1) + 2}', 'A'),
        ('同一行 B 列的', f'{DAY1}!B{_header(v2env, DAY1) + 2}', 'B'),
        ('C 列手写的', f'{DAY1}!C{_header(v2env, DAY1) + 5}', None),
        ('42', f'{DAY1}!B{_header(v2env, DAY1) + 6}', 'B'),
    ]
    assert all(n.source == note_store.SOURCE_EXCEL_IMPORT for n in day1)
    assert day1[0].created_at == task_import._local_to_utc_iso(DAY1)
    assert [n.text for n in note_store.list_day(DAY2)] == ['多行\n文字']
    assert [n.text for n in note_store.list_day(SWITCH)] == ['v2 同步过去的', 'v1 早上存的']
    assert note_store.list_day(NO_HEADER) == []

    # Already in Excel: nothing queued for the mirror.
    assert queue_rows(db) == []
    # Searchable right away.
    found = note_store.search(note_store.NoteQuery(terms=['手写']))
    assert [n.text for n in found] == ['C 列手写的']


def _header(v2env, sheet):
    wb = openpyxl.load_workbook(v2env['excel_path'])
    return PlanService(wb[sheet])._other_header_row()


def test_import_is_idempotent(db, v2env):
    _build_workbook(v2env['excel_path'])
    first = task_import.run_notes_import()
    assert first.notes == 8  # no switch day yet: nothing to dedup
    second = task_import.run_notes_import()
    assert second.notes == 0
    assert db.execute("SELECT COUNT(*) FROM note").fetchone()[0] == 8


def test_missing_excel_leaves_the_import_pending(db, v2env):
    result = task_import.run_notes_import()
    assert result.skipped_reason
    assert task_import.needs_notes_import()


def test_background_runs_the_notes_import_after_the_task_import(db, v2env):
    _build_workbook(v2env['excel_path'])
    assert task_import.start_in_background()
    deadline = time.time() + 30
    while (task_import.is_running() or task_import.is_notes_import_running()) and time.time() < deadline:
        time.sleep(0.05)
    assert not task_import.needs_import()
    assert not task_import.needs_notes_import()
    # The task import finished today, so today is the switch day; the other
    # days are imported in full.
    assert [n.text for n in note_store.list_day(DAY2)] == ['多行\n文字']
    assert not task_import.start_in_background()


def test_existing_database_still_gets_the_notes_import(db, v2env):
    """A DB that finished the task import before this feature existed."""
    _build_workbook(v2env['excel_path'])
    _set_switch_day(db, SWITCH)
    assert not task_import.needs_import()
    assert task_import.start_in_background()
    deadline = time.time() + 30
    while task_import.is_notes_import_running() and time.time() < deadline:
        time.sleep(0.05)
    assert not task_import.is_running()  # the habit dialog was never blocked
    assert not task_import.needs_notes_import()
    assert len(note_store.list_day(DAY1)) == 5
