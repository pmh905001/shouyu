"""End to end: save to SQLite, run the real queue dispatcher once, and check
the Excel display copy (PLN-09)."""
import time

import openpyxl

from conftest import queue_rows
from shouyu.service import dispatch, note_store, task_store
from shouyu.service.plan import (
    BACKLOG_SHEET_LIFE,
    BACKLOG_SHEET_WORK,
    BacklogService,
    PlanService,
    PlanTask,
    TaskCategory,
    TaskStatus,
)
from shouyu.service.task_store import PlanSaveRequest, PomodoroRecord


def _today():
    return time.strftime('%Y-%m-%d')


def _yesterday():
    return time.strftime('%Y-%m-%d', time.localtime(time.time() - 86400))


def _drain(db):
    dispatch._drain_once()
    statuses = {r['status'] for r in queue_rows(db)}
    assert statuses <= {'done'}, queue_rows(db)


def test_plan_backlog_pomodoro_and_note_reach_excel(db, v2env):
    today = _today()
    focus = PlanTask('写代码', status=TaskStatus.IN_PROGRESS, duration_minutes=50)
    task_store.save_plan(PlanSaveRequest(
        plan_date=today,
        today_tasks=[focus, PlanTask('买菜', category=TaskCategory.LIFE, status=TaskStatus.DONE)],
        work_tasks=[PlanTask('整理需求')],
        life_tasks=[PlanTask('换灯泡', category=TaskCategory.LIFE)],
    ))
    start = time.time() - 50 * 60
    task_store.record_pomodoro(PomodoroRecord(
        plan_date=today, task_text='写代码', started_at=start, ended_at=time.time(), actual_minutes=50,
    ))
    note_store.save(text='一段剪贴板文字', excel_column='B')

    _drain(db)

    wb = openpyxl.load_workbook(v2env['excel_path'])
    plan = PlanService(wb[today])
    tasks = plan.read_plan_tasks()
    assert [(t.text, t.status) for t in tasks] == [('写代码', TaskStatus.IN_PROGRESS), ('买菜', TaskStatus.DONE)]
    assert plan.current_in_progress_entry().text == '写代码'
    assert plan.count_pomodoros_logged() == 1
    assert [t.text for t in BacklogService(wb[BACKLOG_SHEET_WORK], TaskCategory.WORK).read()] == ['整理需求']
    assert [t.text for t in BacklogService(wb[BACKLOG_SHEET_LIFE], TaskCategory.LIFE).read()] == ['换灯泡']
    values = [c.value for row in wb[today].iter_rows() for c in row if c.value]
    assert '一段剪贴板文字' in values


def test_history_done_mark_reaches_excel(db, v2env):
    yesterday = _yesterday()
    unfinished = PlanTask('昨天没勾')
    task_store.save_plan(PlanSaveRequest(plan_date=yesterday, today_tasks=[unfinished]))
    _drain(db)

    task_store.save_plan(PlanSaveRequest(
        plan_date=_today(), today_tasks=[], work_tasks=[], life_tasks=[],
        history_date=yesterday, history_done_task_ids=[unfinished.task_id],
    ))
    _drain(db)

    wb = openpyxl.load_workbook(v2env['excel_path'])
    assert [t.status for t in PlanService(wb[yesterday]).read_plan_tasks()] == [TaskStatus.DONE]


def test_excel_edits_are_not_read_back(db, v2env):
    """PLN-03: SQLite is the source of truth; Excel is display only."""
    today = _today()
    task_store.save_plan(PlanSaveRequest(plan_date=today, today_tasks=[PlanTask('原文')]))
    _drain(db)

    wb = openpyxl.load_workbook(v2env['excel_path'])
    wb[today]['B2'] = '在 Excel 里改过'
    wb.save(v2env['excel_path'])

    assert [t.text for t in task_store.load_today(today)] == ['原文']
