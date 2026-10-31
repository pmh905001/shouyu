"""PomodoroService's SQLite hooks (PLN-05), without starting timers or Qt."""
import threading
import time

from conftest import queue_rows
from shouyu.service import task_store
from shouyu.service.plan import PlanTask, TaskStatus
from shouyu.service.pomodoro import PomodoroService
from shouyu.service.task_store import PlanSaveRequest


def _service(mode=PomodoroService.MODE_CLASSIC):
    svc = object.__new__(PomodoroService)
    svc._state_lock = threading.RLock()
    svc._mode = mode
    svc._current_task_text = ''
    svc._current_task_started_at = None
    svc._planning_started_at = None
    svc._task_text_override = None
    return svc


def _today():
    return time.strftime('%Y-%m-%d')


def test_current_task_comes_from_sqlite(db, v2env):
    svc = _service()
    svc._refresh_current_task_text_locked()
    assert svc._current_task_text == ''

    task_store.save_plan(PlanSaveRequest(
        plan_date=_today(), today_tasks=[PlanTask('a'), PlanTask('专注任务', status=TaskStatus.IN_PROGRESS)],
    ))
    svc._refresh_current_task_text_locked()
    assert svc._current_task_text == '专注任务'

    svc._task_text_override = '手动指定'
    svc._refresh_current_task_text_locked()
    assert svc._current_task_text == '手动指定'


def test_work_pomodoro_is_recorded_and_counted(db, v2env):
    focus = PlanTask('专注任务', status=TaskStatus.IN_PROGRESS)
    task_store.save_plan(PlanSaveRequest(plan_date=_today(), today_tasks=[focus]))
    svc = _service(PomodoroService.MODE_DEEP)
    svc._refresh_current_task_text_locked()
    svc._current_task_started_at = time.time() - 600

    svc._record_pomodoro_completion()
    svc._record_pomodoro_completion(actual_minutes=7)

    rows = db.execute("SELECT * FROM pomodoro ORDER BY id").fetchall()
    assert [(r['task_id'], r['kind'], r['mode'], r['ended_early'], r['source']) for r in rows] == [
        (focus.task_id, 'work', 'deep', 0, 'ui'),
        (focus.task_id, 'work', 'deep', 1, 'ui'),
    ]
    assert rows[1]['actual_minutes'] == 7
    assert rows[0]['started_at'] < rows[0]['ended_at']
    assert task_store.count_pomodoros(_today()) == 2
    assert task_store.load_dialog(_today()).today_pomodoros == 2
    assert len(queue_rows(db, 'pomodoro_log')) == 2


def test_planning_session_is_stored_but_not_counted(db, v2env):
    svc = _service()
    svc._planning_started_at = time.time() - 120
    svc._record_planning_completion(ended_early=True)

    row = db.execute("SELECT * FROM pomodoro").fetchone()
    assert (row['kind'], row['ended_early'], row['actual_minutes']) == ('planning', 1, 2)
    assert svc._planning_started_at is None
    assert task_store.count_pomodoros(_today()) == 0
    assert queue_rows(db, 'pomodoro_log') == []


def test_pomodoro_survives_missing_excel(db, v2env):
    """PLN-05 acceptance: no Excel file at all, recording still works."""
    import os

    assert not os.path.exists(v2env['excel_path'])
    svc = _service()
    svc._record_pomodoro_completion()
    assert task_store.count_pomodoros(_today()) == 1
