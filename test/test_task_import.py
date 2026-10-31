import openpyxl

from shouyu.service import datastore, task_import, task_store
from shouyu.service.plan import (
    BACKLOG_SHEET_LIFE,
    BACKLOG_SHEET_WORK,
    BacklogService,
    PlanService,
    PlanTask,
    TaskCategory,
    TaskPriority,
    TaskStatus,
)

DAY1, DAY2 = '2026-10-29', '2026-10-30'


def _build_workbook(path):
    """A v1-style workbook written with v1's own writers."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws1 = wb.create_sheet(DAY1)
    p1 = PlanService(ws1)
    p1.write_plan_tasks([
        PlanTask('写周报', status=TaskStatus.DONE, duration_minutes=30, priority=TaskPriority.P1, reflection='顺利'),
        PlanTask('修 bug', status=TaskStatus.IN_PROGRESS, duration_minutes=50),
        PlanTask('跑步', category=TaskCategory.LIFE),
    ])
    in_progress = p1.read_plan_tasks()[1]
    title_row = p1.switch_in_progress(in_progress)
    ws1.cell(row=title_row + 1, column=2, value='🍅 50min @09:00-09:50')
    ws1.cell(row=title_row + 2, column=2, value='🍅 25min @10:00-10:25')
    ws1.cell(row=title_row + 3, column=2, value='🍅 手写的番茄行')  # unparseable, still counted
    ws1['C4'] = '50m'  # v1 also accepts "50m" durations

    ws2 = wb.create_sheet(DAY2)
    p2 = PlanService(ws2)
    p2.write_plan_tasks([PlanTask('评审')])
    # Older versions wrote bare placeholders with no status/metadata.
    for offset, text in enumerate(('task 1', 'task 2', 'task 3')):
        ws2.cell(row=3 + offset, column=2, value=text)

    BacklogService(wb.create_sheet(BACKLOG_SHEET_WORK), TaskCategory.WORK).write(
        [PlanTask('学习 SQLite', created_date='2026-10-01'), PlanTask('读论文', duration_minutes=90)]
    )
    BacklogService(wb.create_sheet(BACKLOG_SHEET_LIFE), TaskCategory.LIFE).write(
        [PlanTask('修自行车', created_date='2026-09-15')]
    )
    wb.create_sheet('accounts')['A1'] = 'not a dated sheet'
    wb.save(path)


def _counts(conn):
    return {
        table: conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
        for table in ('task', 'day_plan_entry', 'backlog_entry', 'pomodoro')
    }


def test_import_matches_what_v1_shows_and_is_idempotent(db, v2env):
    path = v2env['excel_path']
    _build_workbook(path)
    wb = openpyxl.load_workbook(path)
    expected_day1 = PlanService(wb[DAY1]).read_plan_tasks()
    expected_pomodoros = PlanService(wb[DAY1]).count_pomodoros_logged()
    assert expected_pomodoros == 3

    assert task_import.needs_import()
    result = task_import.run_import()
    assert (result.days, result.tasks, result.backlog, result.pomodoros) == (2, 4, 3, 3)
    assert not task_import.needs_import()

    day1 = task_store.load_day(DAY1)
    assert (day1.done, day1.total, day1.pomodoros) == (1, 3, expected_pomodoros)
    with datastore.reader() as conn:
        imported = task_store._read_day_tasks(conn, DAY1)
    assert [
        (t.text, t.status, t.duration_minutes, t.priority, t.reflection, t.category) for t in imported
    ] == [
        (t.text, t.status, t.duration_minutes, t.priority, t.reflection, t.category) for t in expected_day1
    ]

    # Legacy task 1/2/3 placeholders are dropped, as v1 does when opening the day.
    assert [t.text for t in task_store.load_today(DAY2)] == ['评审']

    data = task_store.load_dialog('2026-10-31')
    assert [(t.text, t.created_date) for t in data.work_tasks][0] == ('学习 SQLite', '2026-10-01')
    assert data.work_tasks[1].duration_minutes == 90
    assert [(t.text, t.created_date) for t in data.life_tasks] == [('修自行车', '2026-09-15')]
    assert data.history.date == DAY2

    pomodoros = db.execute(
        "SELECT task_id, task_text, actual_minutes, started_at, ended_at, source FROM pomodoro ORDER BY id"
    ).fetchall()
    bug_id = next(t.task_id for t in imported if t.text == '修 bug')
    assert [p['task_id'] for p in pomodoros] == [bug_id] * 3
    assert pomodoros[0]['actual_minutes'] == 50 and pomodoros[0]['started_at'].endswith('Z')
    assert pomodoros[2]['actual_minutes'] is None and pomodoros[2]['started_at'] is None
    assert {p['source'] for p in pomodoros} == {'excel_import'}

    # The import never writes into the queue (Excel already has this data).
    assert db.execute("SELECT COUNT(*) FROM queue").fetchone()[0] == 0

    before = _counts(db)
    task_import.run_import()
    assert _counts(db) == before


def test_missing_excel_is_retried_later(db, v2env):
    result = task_import.run_import()
    assert result.skipped_reason
    assert task_import.needs_import()


def test_background_import(db, v2env):
    import time

    _build_workbook(v2env['excel_path'])
    assert task_import.start_in_background()
    deadline = time.time() + 30
    while task_import.is_running() and time.time() < deadline:
        time.sleep(0.05)
    assert not task_import.is_running()
    assert not task_import.needs_import()
    assert not task_import.start_in_background()  # already done
