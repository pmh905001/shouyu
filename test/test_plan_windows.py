"""The morning-ritual dialog and the today panel, driven offscreen against a
temp SQLite file (PLN-01, PLN-02, PLN-04, PLN-11)."""
import os
import time

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

import pytest
from PySide6.QtWidgets import QApplication

from shouyu.service import dispatch, task_store
from shouyu.service.plan import PlanTask, TaskCategory, TaskStatus
from shouyu.service.task_store import PlanSaveRequest
from shouyu.view.msgbox import MessageBox


@pytest.fixture
def qt(db, v2env, monkeypatch):
    app = QApplication.instance() or QApplication([])
    toasts = []
    monkeypatch.setattr(MessageBox, 'pop_up_message', classmethod(lambda cls, *a, **kw: toasts.append(a or kw)))
    monkeypatch.setattr(dispatch, 'notify_recorded', lambda title, msg: toasts.append((title, msg)))
    monkeypatch.setattr(dispatch, 'kick', lambda: None, raising=False)
    yield toasts
    app.processEvents()


def _today():
    return time.strftime('%Y-%m-%d')


def _yesterday():
    return time.strftime('%Y-%m-%d', time.localtime(time.time() - 86400))


def _texts(tasks):
    return [t.text for t in tasks]


def test_habit_dialog_full_flow(qt, db, v2env):
    from shouyu.view.habit_dialog import HabitDialog

    forgot, carry = PlanTask('忘了勾'), PlanTask('要结转')
    task_store.save_plan(PlanSaveRequest(
        plan_date=_yesterday(), today_tasks=[forgot, carry, PlanTask('做完了', status=TaskStatus.DONE)],
        work_tasks=[PlanTask('积压工作')], life_tasks=[PlanTask('积压生活', category=TaskCategory.LIFE)],
    ))

    dialog = HabitDialog()
    dialog.refresh_plan_from_excel()
    assert dialog._tasks == []
    assert _texts(dialog._work_tasks) == ['积压工作']
    assert _texts(dialog._life_tasks) == ['积压生活']
    assert dialog._selected_history_date == _yesterday()
    assert _texts(dialog._yesterday_unfinished) == ['忘了勾', '要结转']
    assert not dialog._has_unsaved_changes()

    # "已完成" on the history card, then carry the rest over.
    dialog._mark_yesterday_done(next(t for t in dialog._yesterday_unfinished if t.text == '忘了勾'))
    assert dialog._has_unsaved_changes()
    dialog._set_carryover_all(True)
    dialog._carry_over_now()
    # Pull a backlog task into today and add a brand-new one.
    dialog._tasks.append(dialog._work_tasks.pop(0))
    dialog._tasks.append(PlanTask('新任务', status=TaskStatus.IN_PROGRESS))

    dialog._dispatch_save()
    assert ('今日任务已保存', '已保存到本地数据库') in qt
    assert not dialog._has_unsaved_changes()

    today = task_store.load_today(_today())
    assert _texts(today) == ['要结转', '积压工作', '新任务']
    assert today[0].task_id == carry.task_id  # PLN-11: same task across days
    yesterday = task_store.load_day(_yesterday())
    assert (yesterday.done, yesterday.total) == (2, 3)
    assert db.execute("SELECT COUNT(*) FROM backlog_entry").fetchone()[0] == 1  # only 积压生活

    # Saving again (e.g. "专注此项" then close) must not duplicate anything.
    task_count = db.execute("SELECT COUNT(*) FROM task WHERE deleted_at IS NULL").fetchone()[0]
    dialog._dispatch_save()
    assert db.execute("SELECT COUNT(*) FROM task WHERE deleted_at IS NULL").fetchone()[0] == task_count

    # Reopen: the dialog shows exactly what was saved.
    reopened = HabitDialog()
    reopened.refresh_plan_from_excel()
    assert _texts(reopened._tasks) == ['要结转', '积压工作', '新任务']
    assert _texts(reopened._life_tasks) == ['积压生活']
    assert reopened._original_in_progress == '新任务'
    assert _texts(reopened._yesterday_unfinished) == ['要结转']
    assert _texts(reopened._visible_yesterday) == []  # hidden: already on today
    dialog.deleteLater()
    reopened.deleteLater()


def test_habit_dialog_blocks_saving_during_import(qt, db, v2env):
    from shouyu.view.habit_dialog import HabitDialog

    dialog = HabitDialog()
    dialog.refresh_plan_from_excel()
    dialog._set_importing(True)
    dialog._tasks.append(PlanTask('导入时加的'))
    dialog._save_and_accept()
    assert any('正在导入' in str(t) for t in qt)
    assert task_store.load_today(_today()) == []
    dialog._set_importing(False)
    dialog.deleteLater()


def test_todo_panel_and_dialog_share_data(qt, db, v2env):
    from shouyu.view.habit_dialog import HabitDialog
    from shouyu.view.todo_panel import TodoPanel

    task_store.save_plan(PlanSaveRequest(
        plan_date=_today(), today_tasks=[PlanTask('甲'), PlanTask('乙')], work_tasks=[PlanTask('积压')],
    ))

    panel = TodoPanel()
    panel.refresh_from_excel()
    assert _texts(panel._tasks) == ['甲', '乙']
    panel._tasks[0].status = TaskStatus.DONE
    panel._tasks.append(PlanTask('面板里加的'))
    panel._dispatch_save()
    assert ('今日任务已保存', '已保存到本地数据库') in qt

    dialog = HabitDialog()
    dialog.refresh_plan_from_excel()
    assert [(t.text, t.status) for t in dialog._tasks] == [
        ('甲', TaskStatus.DONE), ('乙', TaskStatus.PENDING), ('面板里加的', TaskStatus.PENDING),
    ]
    assert _texts(dialog._work_tasks) == ['积压']  # the panel leaves the backlog alone

    dialog._tasks[1].status = TaskStatus.IN_PROGRESS
    dialog._dispatch_save()
    panel.refresh_from_excel()
    assert panel._original_in_progress == '乙'
    panel.deleteLater()
    dialog.deleteLater()


def test_windows_work_without_excel_file(qt, db, v2env):
    """PLN-01 acceptance: no Excel file, open / edit / save / reopen."""
    from shouyu.view.habit_dialog import HabitDialog

    assert not os.path.exists(v2env['excel_path'])
    dialog = HabitDialog()
    dialog.refresh_plan_from_excel()
    dialog._tasks.append(PlanTask('没有 Excel 也能存'))
    dialog._dispatch_save()
    reopened = HabitDialog()
    reopened.refresh_plan_from_excel()
    assert _texts(reopened._tasks) == ['没有 Excel 也能存']
    dialog.deleteLater()
    reopened.deleteLater()
