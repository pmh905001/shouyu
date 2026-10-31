import time

from conftest import queue_rows
from shouyu.service import task_store
from shouyu.service.plan import PlanTask, TaskCategory, TaskPriority, TaskStatus
from shouyu.service.task_store import PlanSaveRequest, PomodoroRecord

TODAY = '2026-10-31'
YESTERDAY = '2026-10-30'


def _save(today, work=None, life=None, plan_date=TODAY, **kw):
    return task_store.save_plan(
        PlanSaveRequest(plan_date=plan_date, today_tasks=today, work_tasks=work, life_tasks=life, **kw)
    )


def _texts(tasks):
    return [t.text for t in tasks]


def test_save_and_reload_round_trip(db):
    today = [
        PlanTask('写设计', status=TaskStatus.IN_PROGRESS, duration_minutes=50, priority=TaskPriority.P1),
        PlanTask('买菜', category=TaskCategory.LIFE, reflection='早点去'),
        PlanTask('   '),  # empty rows (auto-extend) are never saved
    ]
    work = [PlanTask('整理需求', category=TaskCategory.WORK)]
    life = [PlanTask('换灯泡', category=TaskCategory.LIFE, created_date='2026-10-20')]

    result = _save(today, work, life)

    assert result.changed
    assert all(t.task_id for t in today[:2] + work + life)
    assert today[2].task_id == 0
    # Backlog tasks without an entry date get today's date stamped, like v1.
    assert work[0].created_date == time.strftime('%Y-%m-%d')
    assert life[0].created_date == '2026-10-20'

    data = task_store.load_dialog(TODAY)
    assert _texts(data.today_tasks) == ['写设计', '买菜']
    first, second = data.today_tasks
    assert (first.status, first.duration_minutes, first.priority) == (TaskStatus.IN_PROGRESS, 50, TaskPriority.P1)
    assert (second.category, second.reflection) == (TaskCategory.LIFE, '早点去')
    assert _texts(data.work_tasks) == ['整理需求']
    assert _texts(data.life_tasks) == ['换灯泡']
    assert data.life_tasks[0].created_date == '2026-10-20'


def test_status_changes_are_logged(db):
    a, b = PlanTask('A'), PlanTask('B')
    _save([a, b])
    a.status = TaskStatus.IN_PROGRESS
    _save([a, b])
    a.status = TaskStatus.DONE
    b.status = TaskStatus.IN_PROGRESS
    _save([b, a])

    log = db.execute(
        "SELECT task_id, from_status, to_status FROM task_status_log ORDER BY id"
    ).fetchall()
    assert [tuple(r) for r in log] == [
        (a.task_id, None, 'pending'),
        (b.task_id, None, 'pending'),
        (a.task_id, 'pending', 'in_progress'),
        # one save logs in list order: [b, a]
        (b.task_id, 'pending', 'in_progress'),
        (a.task_id, 'in_progress', 'done'),
    ]
    assert _texts(task_store.load_today(TODAY)) == ['B', 'A']


def test_move_between_today_and_backlog_keeps_identity(db):
    task = PlanTask('跨列表移动')
    _save([task], [], [])
    task_id = task.task_id

    # today -> work backlog
    task.created_date = '2026-10-31'
    _save([], [task], [])
    data = task_store.load_dialog(TODAY)
    assert data.today_tasks == []
    assert [t.task_id for t in data.work_tasks] == [task_id]

    # work -> life keeps the entry date, then back to today
    task.category = TaskCategory.LIFE
    _save([], [], [task])
    data = task_store.load_dialog(TODAY)
    assert [t.task_id for t in data.life_tasks] == [task_id]
    assert data.life_tasks[0].created_date == '2026-10-31'

    task.created_date = ''
    _save([task], [], [])
    data = task_store.load_dialog(TODAY)
    assert [t.task_id for t in data.today_tasks] == [task_id]
    assert data.work_tasks == data.life_tasks == []
    assert db.execute("SELECT COUNT(*) FROM task").fetchone()[0] == 1
    assert db.execute("SELECT deleted_at FROM task").fetchone()[0] is None


def test_delete_soft_deletes_unless_task_has_history(db):
    old = PlanTask('昨天的')
    _save([old], plan_date=YESTERDAY)
    carried = PlanTask('昨天的', task_id=old.task_id)
    fresh = PlanTask('今天新建')
    _save([carried, fresh])

    _save([])  # both removed from today

    rows = dict(db.execute("SELECT id, deleted_at FROM task").fetchall())
    assert rows[old.task_id] is None  # still on yesterday's plan: history
    assert rows[fresh.task_id] is not None


def test_carry_over_and_mark_history_done(db):
    y1, y2, y3 = PlanTask('Y1'), PlanTask('Y2'), PlanTask('Y3', status=TaskStatus.DONE)
    _save([y1, y2, y3], plan_date=YESTERDAY)

    snap = task_store.load_day(YESTERDAY)
    assert (snap.done, snap.total) == (1, 3)
    assert [(t.text, t.task_id) for t in snap.unfinished] == [('Y1', y1.task_id), ('Y2', y2.task_id)]
    assert all(t.status == TaskStatus.PENDING for t in snap.unfinished)

    carried = PlanTask('Y1', task_id=y1.task_id)
    _save([carried], history_date=YESTERDAY, history_done_task_ids=[y2.task_id])

    assert carried.task_id == y1.task_id
    entries = db.execute(
        "SELECT plan_date, status FROM day_plan_entry WHERE task_id = ? ORDER BY plan_date", (y1.task_id,)
    ).fetchall()
    assert [tuple(e) for e in entries] == [(YESTERDAY, 'pending'), (TODAY, 'pending')]
    assert task_store.load_day(YESTERDAY).done == 2
    done_log = db.execute(
        "SELECT plan_date, from_status, to_status FROM task_status_log WHERE task_id = ? ORDER BY id DESC LIMIT 1",
        (y2.task_id,),
    ).fetchone()
    assert tuple(done_log) == (YESTERDAY, 'pending', 'done')
    mirror = queue_rows(db, 'plan_mark_done_by_text')
    assert mirror[-1]['payload'] == {'date': YESTERDAY, 'texts': ['Y2']}


def test_unsaved_copy_of_saved_task_is_not_duplicated(db):
    """An undo snapshot taken before a save holds task_id == 0 copies; saving
    it again must reuse the rows created by the first save."""
    live = [PlanTask('A'), PlanTask('B')]
    undo_snapshot = [PlanTask('A'), PlanTask('B')]
    _save(live)
    _save(undo_snapshot)
    assert db.execute("SELECT COUNT(*) FROM task").fetchone()[0] == 2
    assert [t.task_id for t in undo_snapshot] == [t.task_id for t in live]


def test_same_task_twice_in_one_save_gets_two_rows(db):
    a = PlanTask('A')
    _save([a])
    dup = PlanTask('A', task_id=a.task_id)
    _save([a, dup])
    assert a.task_id != dup.task_id
    assert len(task_store.load_today(TODAY)) == 2


def test_todo_panel_save_leaves_backlog_alone(db):
    _save([PlanTask('today')], [PlanTask('w')], [PlanTask('l', category=TaskCategory.LIFE)])
    today = task_store.load_today(TODAY)
    today.append(PlanTask('panel added'))
    _save(today)  # work/life None, like TodoPanel
    data = task_store.load_dialog(TODAY)
    assert _texts(data.today_tasks) == ['today', 'panel added']
    assert _texts(data.work_tasks) == ['w']
    assert _texts(data.life_tasks) == ['l']


def test_mirror_messages(db, v2env):
    task = PlanTask('A', status=TaskStatus.IN_PROGRESS)
    _save([task], [PlanTask('w')], [], original_in_progress=None)
    msgs = queue_rows(db, 'plan_save')
    assert len(msgs) == 1
    payload = msgs[0]['payload']
    assert payload['date'] == TODAY
    assert [t['text'] for t in payload['tasks']] == ['A']
    assert payload['tasks'][0]['task_id'] == task.task_id
    assert [t['text'] for t in payload['work_tasks']] == ['w']
    assert payload['life_tasks'] == []

    # Nothing changed -> nothing to mirror.
    assert not _save([task], [PlanTask('w', task_id=msgs[0]['payload']['work_tasks'][0]['task_id'])], []).changed
    assert len(queue_rows(db, 'plan_save')) == 1

    # Todo-panel save mirrors without backlog fields.
    task.text = 'A2'
    _save([task])
    assert queue_rows(db, 'plan_save')[-1]['payload']['work_tasks'] is None

    v2env['settings']['excel_mirror'] = False
    task.text = 'A3'
    assert _save([task]).changed
    assert len(queue_rows(db, 'plan_save')) == 2


def test_history_day_lookup_and_counts(db):
    _save([PlanTask('old')], plan_date='2026-09-01')  # older than 30 days
    _save([PlanTask('x', status=TaskStatus.DONE), PlanTask('y')], plan_date='2026-10-28')
    data = task_store.load_dialog(TODAY)
    assert data.history.date == '2026-10-28'
    assert (data.history.done, data.history.total) == (1, 2)

    data = task_store.load_dialog('2026-10-01')  # exactly 30 days later: still found
    assert data.history.date == '2026-09-01'
    data = task_store.load_dialog('2026-10-02')
    assert data.history.date == '2026-10-01'  # nothing within 30 days -> yesterday
    assert data.history.total == 0


def test_pomodoro_records_and_counts(db):
    focus = PlanTask('写代码', status=TaskStatus.IN_PROGRESS)
    _save([PlanTask('别的'), focus])
    assert task_store.current_in_progress(TODAY).task_id == focus.task_id

    start = time.mktime(time.strptime('2026-10-31 09:00', '%Y-%m-%d %H:%M'))
    pid = task_store.record_pomodoro(PomodoroRecord(
        plan_date=TODAY, task_text='写代码', mode='deep', started_at=start, ended_at=start + 50 * 60,
        planned_minutes=50, actual_minutes=50,
    ))
    task_store.record_pomodoro(PomodoroRecord(
        plan_date=TODAY, task_text='改名后的任务', started_at=start, ended_at=start + 600,
        planned_minutes=25, actual_minutes=10, ended_early=True,
    ))
    task_store.record_pomodoro(PomodoroRecord(
        plan_date=TODAY, kind=task_store.KIND_PLANNING, started_at=start, ended_at=start + 600,
        planned_minutes=10, actual_minutes=10,
    ))

    rows = db.execute("SELECT id, task_id, kind, ended_early, started_at FROM pomodoro ORDER BY id").fetchall()
    assert rows[0]['id'] == pid and rows[0]['task_id'] == focus.task_id
    assert rows[1]['task_id'] is None and rows[1]['ended_early'] == 1
    assert rows[2]['kind'] == 'planning'
    assert rows[0]['started_at'].endswith('Z')
    assert task_store.count_pomodoros(TODAY) == 2
    assert task_store.load_dialog(TODAY).today_pomodoros == 2

    logs = queue_rows(db, 'pomodoro_log')
    assert len(logs) == 2  # planning is not mirrored
    assert logs[0]['payload']['label'] == '🍅 50min @09:00-09:50'
    assert logs[0]['payload']['date'] == TODAY
