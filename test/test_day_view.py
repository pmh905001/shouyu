"""Day view, driven offscreen against a temp database (DAY-01 .. DAY-13)."""
import os
import time

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

import pytest
from PySide6.QtCore import QItemSelectionModel, Qt
from PySide6.QtGui import QGuiApplication
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QMessageBox

from shouyu.service import note_store, task_store
from shouyu.service.plan import PlanTask, TaskCategory, TaskPriority, TaskStatus
from shouyu.service.task_store import PlanSaveRequest, PomodoroRecord
from shouyu.view import task_table
from shouyu.view.msgbox import MessageBox

DAY = '2026-10-29'


def _texts(tasks):
    return [t.text for t in tasks]


@pytest.fixture
def view(db, v2env, monkeypatch):
    app = QApplication.instance() or QApplication([])
    v2env['settings']['excel_mirror'] = False
    monkeypatch.setattr(MessageBox, 'pop_up_message', classmethod(lambda cls, *a, **kw: None))
    task_store.save_plan(PlanSaveRequest(
        plan_date=DAY,
        today_tasks=[PlanTask('写周报', status=TaskStatus.IN_PROGRESS, duration_minutes=30), PlanTask('评审')],
        work_tasks=[PlanTask('学习 SQLite', created_date='2026-10-01')],
        life_tasks=[PlanTask('修自行车', category=TaskCategory.LIFE)],
    ))
    base = time.mktime(time.strptime(f'{DAY} 09:00', '%Y-%m-%d %H:%M'))
    task_store.record_pomodoro(PomodoroRecord(
        plan_date=DAY, task_text='', kind=task_store.KIND_PLANNING, started_at=base, ended_at=base + 600,
        actual_minutes=10,
    ))
    task_store.record_pomodoro(PomodoroRecord(
        plan_date=DAY, task_text='写周报', started_at=base + 900, ended_at=base + 900 + 3000, actual_minutes=50,
        mode='deep',
    ))
    db.execute(
        "INSERT INTO note (plan_date, kind, text, source, created_at) VALUES (?, 'text', '那天的记录', 'clipboard', ?)",
        (DAY, '2026-10-29T02:00:00.000Z'),
    )
    from shouyu.view.day_view import DayView

    DayView._instance = None
    v = DayView.get_or_create()
    v.open_date(DAY)
    yield v
    v._backlog_undo.clear()
    v.hide()
    v.deleteLater()
    DayView._instance = None
    app.processEvents()


def test_open_shows_tasks_pomodoros_and_notes(view):
    assert view.plan_date == DAY
    assert view.date_label.text() == '2026-10-29 星期四'
    assert _texts(view.day_model.tasks) == ['写周报', '评审']
    assert view.tasks_title.text() == '任务 · 0/2 完成'
    assert view.pomodoro_table.rowCount() == 2
    assert view.pomodoro_title.text() == '番茄（只读）· 1 个'  # planning is not counted
    assert view.pomodoro_table.item(0, 4).text() == '规划'
    assert view.pomodoro_table.item(1, 0).text() == '09:15–10:05'
    assert view.pomodoro_table.item(1, 3).text() == '深度'
    assert [n.text for n in view.notes_view.note_model.notes()] == ['那天的记录']
    assert view.notes_title.text() == '记录（1）'
    assert _texts(view.work_model.tasks) == ['学习 SQLite']
    assert not view.is_dirty() and view.dirty_label.isHidden()


def test_edit_and_save_with_one_in_progress(view):
    model = view.day_model
    status_col = model.columns.index(task_table.COL_STATUS)
    model.setData(model.index(1, status_col), TaskStatus.IN_PROGRESS)
    assert [t.status for t in model.tasks] == [TaskStatus.PENDING, TaskStatus.IN_PROGRESS]
    model.setData(model.index(1, model.columns.index(task_table.COL_PRIORITY)), TaskPriority.P1)
    model.setData(model.index(1, model.columns.index(task_table.COL_DURATION)), 45)
    model.setData(model.index(0, 0), '写周报（改）')
    assert view.is_dirty() and not view.dirty_label.isHidden()

    assert view.save()
    assert not view.is_dirty()
    saved = task_store.load_today(DAY)
    assert [(t.text, t.status, t.priority, t.duration_minutes) for t in saved] == [
        ('写周报（改）', TaskStatus.PENDING, TaskPriority.NONE, 30),
        ('评审', TaskStatus.IN_PROGRESS, TaskPriority.P1, 45),
    ]
    # The version was refreshed: a second edit + save does not conflict.
    model.setData(model.index(0, status_col), TaskStatus.DONE)
    assert view.save()
    assert task_store.load_today(DAY)[0].status == TaskStatus.DONE
    assert view.tasks_title.text() == '任务 · 1/2 完成'


def _change_elsewhere():
    other = task_store.load_today(DAY)
    other.append(PlanTask('晨间仪式里加的'))
    task_store.save_plan(PlanSaveRequest(plan_date=DAY, today_tasks=other))


@pytest.mark.parametrize('choice', ['reload', 'force', 'cancel'])
def test_conflict(view, monkeypatch, choice):
    view.day_model.setData(view.day_model.index(1, 0), '日视图里改的')
    _change_elsewhere()
    asked = []
    monkeypatch.setattr(view, '_ask_conflict', lambda what: asked.append(what) or choice)
    result = view.save()
    assert asked == [f'{DAY} 的任务']
    saved = _texts(task_store.load_today(DAY))
    if choice == 'reload':
        assert result and saved == ['写周报', '评审', '晨间仪式里加的']
        assert _texts(view.day_model.tasks) == saved and not view.is_dirty()
    elif choice == 'force':
        assert result and saved == ['写周报', '日视图里改的']
        assert not view.is_dirty()
    else:
        assert not result and saved == ['写周报', '评审', '晨间仪式里加的']
        assert view.is_dirty()


def test_backlog_page_saves_only_the_backlog(view):
    view.tabs.setCurrentIndex(1)
    view.work_table.setCurrentIndex(view.work_model.index(0, 0))
    view.work_table.add_row()
    view.work_model.setData(view.work_model.index(1, 0), '读论文')
    view.work_table.setCurrentIndex(view.work_model.index(0, 0))
    view.move_backlog_task(view.work_table, TaskCategory.LIFE)
    assert _texts(view.work_model.tasks) == ['读论文']
    assert _texts(view.life_model.tasks) == ['修自行车', '学习 SQLite']
    assert view.backlog_dirty() and not view.day_dirty()
    shelved = view.life_model.index(1, view.life_model.columns.index(task_table.COL_SHELVED)).data()
    assert shelved.endswith(' 天')

    _change_elsewhere()  # the day changed elsewhere: irrelevant for a backlog-only save
    assert view.save()
    work, life, _ = task_store.load_backlog_for_edit()
    assert _texts(work) == ['读论文']
    assert _texts(life) == ['修自行车', '学习 SQLite']
    assert life[1].category == TaskCategory.LIFE
    assert life[1].created_date == '2026-10-01'  # keeps how long it has been shelved
    assert _texts(task_store.load_today(DAY)) == ['写周报', '评审', '晨间仪式里加的']


def test_undo_redo(view):
    model = view.day_model
    model.setData(model.index(0, 0), '改一次')
    view.day_table.setCurrentIndex(model.index(1, 0))
    view.day_table.add_row()
    assert len(model.tasks) == 3
    view.undo()
    assert _texts(model.tasks) == ['改一次', '评审']
    view.undo()
    assert _texts(model.tasks) == ['写周报', '评审']
    assert not view.is_dirty()
    view.redo()
    assert _texts(model.tasks) == ['改一次', '评审']


def test_keyboard_add_move_delete_copy(view, monkeypatch):
    table, model = view.day_table, view.day_model
    table.setCurrentIndex(model.index(0, 0))
    QTest.keyClick(table, Qt.Key_Return, Qt.ControlModifier)
    assert len(model.tasks) == 3 and model.tasks[1].text == ''
    QTest.keyClick(QApplication.focusWidget() or table, Qt.Key_Escape)

    table.setCurrentIndex(model.index(2, 0))
    QTest.keyClick(table, Qt.Key_Up, Qt.AltModifier)
    assert _texts(model.tasks) == ['写周报', '评审', '']
    assert table.currentIndex().row() == 1

    table.setCurrentIndex(model.index(2, 0))
    QTest.keyClick(table, Qt.Key_Delete)  # empty row: no question asked
    assert _texts(model.tasks) == ['写周报', '评审']
    monkeypatch.setattr(QMessageBox, 'question', staticmethod(lambda *a, **kw: QMessageBox.No))
    table.setCurrentIndex(model.index(1, 0))
    QTest.keyClick(table, Qt.Key_Delete)
    assert _texts(model.tasks) == ['写周报', '评审']  # said no

    table.selectionModel().clearSelection()
    table.setCurrentIndex(model.index(0, 0))
    for row, col in ((0, 0), (1, 1), (0, 1), (1, 0)):
        table.selectionModel().select(model.index(row, col), QItemSelectionModel.Select)
    QTest.keyClick(table, Qt.Key_C, Qt.ControlModifier)
    assert QGuiApplication.clipboard().text() == '写周报\t进行中\n评审\t待办'


def test_switching_days_and_unsaved_prompt(view, monkeypatch):
    view.go_to('2026-10-30')
    assert view.plan_date == '2026-10-30' and view.day_model.tasks == []
    view.go_to(DAY)
    view.day_model.setData(view.day_model.index(0, 0), '没存')
    monkeypatch.setattr(view, '_confirm_leave', lambda: False)
    view.go_to('2026-10-30')
    assert view.plan_date == DAY and view.day_model.tasks[0].text == '没存'


def test_manual_note(view):
    view.note_input.setText('手动加的')
    view._add_note()
    assert [n.text for n in note_store.list_day(DAY)] == ['那天的记录', '手动加的']
    assert view.notes_title.text() == '记录（2）'
    assert view.note_input.text() == ''


def test_a_day_with_many_notes_is_paged(view, db):
    db.execute('BEGIN')
    for i in range(450):
        db.execute(
            "INSERT INTO note (plan_date, kind, text, source, created_at) VALUES ('2026-10-20', 'text', ?, "
            "'excel_import', '2026-10-19T16:00:00.000Z')", (f'格子 {i}',),
        )
    db.execute('COMMIT')
    view.go_to('2026-10-20')
    shown = view.notes_view.note_model.notes()
    assert len(shown) == 200 and shown[-1].text == '格子 449'
    assert view.notes_title.text() == '记录（450，已加载 200 条，往上滚动加载更多）'
    view._load_older_notes()
    view._load_older_notes()
    view._load_older_notes()  # nothing left: no error
    assert len(view.notes_view.note_model.notes()) == 450
    assert view.notes_title.text() == '记录（450）'


def test_saving_today_updates_the_pomodoro_label(view, monkeypatch):
    calls = []

    class _Fake:
        def set_current_task(self, text):
            calls.append(text)

    from shouyu.service import pomodoro

    monkeypatch.setattr(pomodoro.PomodoroService, 'instance', classmethod(lambda cls: _Fake()))
    view.go_to(time.strftime('%Y-%m-%d'))
    view.day_table.add_row()
    view.day_model.setData(view.day_model.index(0, 0), '今天的重点')
    view.day_model.setData(view.day_model.index(0, 1), TaskStatus.IN_PROGRESS)
    assert view.save()
    assert calls == ['今天的重点']
    view.go_to(DAY)
    view.day_model.setData(view.day_model.index(0, 0), '旧日子')
    assert view.save()
    assert calls == ['今天的重点']  # other days never touch the timer


def test_shelved_days():
    assert task_table.shelved_days('2026-10-01', '2026-10-31') == 30
    assert task_table.shelved_days('', '2026-10-31') is None
