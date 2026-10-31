"""task_store additions for the day view (design §3.4): backlog-only saves,
version checks, pomodoro reads."""
import time

import pytest

from conftest import queue_rows
from shouyu.service import task_store
from shouyu.service.plan import PlanTask, TaskCategory, TaskStatus
from shouyu.service.task_store import ConflictError, PlanSaveRequest, PomodoroRecord

DAY = '2026-10-29'


def _texts(tasks):
    return [t.text for t in tasks]


def _seed(v2env):
    v2env['settings']['excel_mirror'] = False
    task_store.save_plan(PlanSaveRequest(
        plan_date=DAY,
        today_tasks=[PlanTask('写周报', status=TaskStatus.IN_PROGRESS), PlanTask('评审')],
        work_tasks=[PlanTask('学习 SQLite', created_date='2026-10-01')],
        life_tasks=[PlanTask('修自行车', category=TaskCategory.LIFE)],
    ))


def test_backlog_only_save_leaves_the_day_untouched(db, v2env):
    _seed(v2env)
    day_before = task_store.day_version(DAY)
    work, life, version = task_store.load_backlog_for_edit()
    work.append(PlanTask('读论文'))
    life[0].text = '修自行车链条'
    task_store.save_plan(PlanSaveRequest(
        plan_date=DAY, today_tasks=None, work_tasks=work, life_tasks=life, expected_backlog_version=version,
    ))
    assert _texts(task_store.load_today(DAY)) == ['写周报', '评审']
    assert task_store.day_version(DAY) == day_before
    work, life, _ = task_store.load_backlog_for_edit()
    assert _texts(work) == ['学习 SQLite', '读论文'] and _texts(life) == ['修自行车链条']


def test_backlog_only_save_does_not_steal_a_day_task_with_the_same_text(db, v2env):
    _seed(v2env)
    work, life, _ = task_store.load_backlog_for_edit()
    work.append(PlanTask('评审'))  # same text as a task on the day, but a new backlog row
    task_store.save_plan(PlanSaveRequest(plan_date=DAY, today_tasks=None, work_tasks=work, life_tasks=life))
    day = task_store.load_today(DAY)
    work, _life, _ = task_store.load_backlog_for_edit()
    assert _texts(day) == ['写周报', '评审']
    assert day[1].task_id != work[-1].task_id


def test_backlog_only_mirror_message_does_not_carry_day_tasks(db, v2env):
    _seed(v2env)
    v2env['settings']['excel_mirror'] = True
    work, life, _ = task_store.load_backlog_for_edit()
    work.append(PlanTask('读论文'))
    task_store.save_plan(PlanSaveRequest(plan_date=DAY, today_tasks=None, work_tasks=work, life_tasks=life))
    payload = queue_rows(db, 'plan_save')[-1]['payload']
    assert payload['tasks'] is None
    assert [t['text'] for t in payload['work_tasks']] == ['学习 SQLite', '读论文']


@pytest.mark.parametrize('change', ['rename', 'status', 'reorder', 'delete', 'add', 'other_day_rename'])
def test_day_version_detects_every_kind_of_change(db, v2env, change):
    _seed(v2env)
    tasks, version = task_store.load_day_for_edit(DAY)
    time.sleep(0.002)  # updated_at has millisecond resolution
    other = task_store.load_today(DAY)
    if change == 'rename':
        other[1].text = '评审 v2'
    elif change == 'status':
        other[1].status = TaskStatus.DONE
    elif change == 'reorder':
        other.reverse()
    elif change == 'delete':
        other.pop()
    elif change == 'add':
        other.append(PlanTask('新任务'))
    if change == 'other_day_rename':
        # Renaming the same task while it sits on another day changes it here too.
        moved = task_store.load_today(DAY)[1]
        moved.text = '评审（改名）'
        task_store.save_plan(PlanSaveRequest(plan_date='2026-10-30', today_tasks=[moved]))
    else:
        task_store.save_plan(PlanSaveRequest(plan_date=DAY, today_tasks=other))
    assert task_store.day_version(DAY) != version

    tasks.append(PlanTask('日视图里加的'))
    with pytest.raises(ConflictError) as info:
        task_store.save_plan(PlanSaveRequest(plan_date=DAY, today_tasks=tasks, expected_day_version=version))
    assert info.value.what == 'day'
    assert '日视图里加的' not in _texts(task_store.load_today(DAY))  # nothing written


def test_matching_version_saves_and_unrelated_changes_do_not_conflict(db, v2env):
    _seed(v2env)
    tasks, version = task_store.load_day_for_edit(DAY)
    # Another day and the backlog changing do not matter for this day.
    task_store.save_plan(PlanSaveRequest(plan_date='2026-10-30', today_tasks=[PlanTask('别的日子')]))
    work, life, _ = task_store.load_backlog_for_edit()
    task_store.save_plan(PlanSaveRequest(plan_date=DAY, today_tasks=None, work_tasks=work + [PlanTask('x')], life_tasks=life))
    tasks[0].status = TaskStatus.DONE
    task_store.save_plan(PlanSaveRequest(plan_date=DAY, today_tasks=tasks, expected_day_version=version))
    assert task_store.load_today(DAY)[0].status == TaskStatus.DONE


def test_backlog_version_conflict(db, v2env):
    _seed(v2env)
    work, life, version = task_store.load_backlog_for_edit()
    other_work, other_life, _ = task_store.load_backlog_for_edit()
    other_work.reverse()
    other_work.append(PlanTask('别处加的'))
    task_store.save_plan(PlanSaveRequest(plan_date=DAY, today_tasks=None, work_tasks=other_work, life_tasks=other_life))
    with pytest.raises(ConflictError) as info:
        task_store.save_plan(PlanSaveRequest(
            plan_date=DAY, today_tasks=None, work_tasks=work, life_tasks=life, expected_backlog_version=version,
        ))
    assert info.value.what == 'backlog'


def test_backlog_reorder_changes_the_version(db, v2env):
    _seed(v2env)
    task_store.save_plan(PlanSaveRequest(
        plan_date=DAY, today_tasks=None, work_tasks=[PlanTask('a'), PlanTask('b')], life_tasks=[],
    ))
    work, life, version = task_store.load_backlog_for_edit()
    task_store.save_plan(PlanSaveRequest(plan_date=DAY, today_tasks=None, work_tasks=list(reversed(work)), life_tasks=life))
    assert task_store.backlog_version() != version


def test_load_pomodoros(db, v2env):
    v2env['settings']['excel_mirror'] = False
    base = time.mktime(time.strptime(f'{DAY} 09:00', '%Y-%m-%d %H:%M'))
    task_store.record_pomodoro(PomodoroRecord(
        plan_date=DAY, task_text='写周报', started_at=base + 3600, ended_at=base + 3600 + 1500, actual_minutes=25,
        mode='classic', ended_early=True,
    ))
    task_store.record_pomodoro(PomodoroRecord(
        plan_date=DAY, task_text='', kind=task_store.KIND_PLANNING, started_at=base, ended_at=base + 600,
        actual_minutes=10,
    ))
    task_store.record_pomodoro(PomodoroRecord(plan_date='2026-10-30', task_text='别的日子'))
    rows = task_store.load_pomodoros(DAY)
    assert [(r.kind, r.actual_minutes) for r in rows] == [('planning', 10), ('work', 25)]
    assert rows[1].started_at == pytest.approx(base + 3600, abs=0.01)
    assert rows[1].ended_early and rows[1].mode == 'classic' and rows[1].pomodoro_id
    assert task_store.count_pomodoros(DAY) == 1  # planning ones are not counted
