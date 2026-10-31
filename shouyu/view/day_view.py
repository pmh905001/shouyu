"""Day view (ctrl+alt+\\): one day's tasks, pomodoros and notes read from
SQLite, editable like a table; plus the backlog (design §4.5, DAY-01..13).

Saves go through `task_store.save_plan()` with the version read at load
time, so an edit made meanwhile in the habit dialog / todo panel is not
silently overwritten (DAY-11).
"""
from __future__ import annotations

import copy
import logging
import time
from datetime import datetime, timedelta
from typing import List, Optional

from PySide6.QtCore import QDate, QModelIndex, Qt
from PySide6.QtGui import QColor, QCursor, QGuiApplication, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QAbstractItemDelegate,
    QAbstractItemView,
    QApplication,
    QCalendarWidget,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QToolButton,
    QVBoxLayout,
    QWidget,
    QWidgetAction,
)

from shouyu.service import note_store, task_store
from shouyu.service.plan import PlanTask, TaskCategory, TaskStatus
from shouyu.service.task_store import ConflictError, PlanSaveRequest, PomodoroRecord
from shouyu.view.note_list import NoteListView
from shouyu.view.styles import SUBTEXT_COLOR_HEX, SUSPENDED_COLOR_HEX
from shouyu.view.task_table import (
    BACKLOG_COLUMNS,
    DAY_COLUMNS,
    TaskTableModel,
    TaskTableView,
    snapshot_tasks,
)

_WEEKDAYS = '一二三四五六日'
UNDO_LIMIT = 100


def _shift(date_str: str, days: int) -> str:
    return (datetime.strptime(date_str, '%Y-%m-%d') + timedelta(days=days)).strftime('%Y-%m-%d')


def _today() -> str:
    return time.strftime('%Y-%m-%d')


def _date_title(plan_date: str) -> str:
    d = datetime.strptime(plan_date, '%Y-%m-%d')
    suffix = '（今天）' if plan_date == _today() else ''
    return f'{plan_date} 星期{_WEEKDAYS[d.weekday()]}{suffix}'


def _hhmm(ts: Optional[float]) -> str:
    return time.strftime('%H:%M', time.localtime(ts)) if ts else '--:--'


class _Undo:
    """Whole-page snapshots, like the habit dialog (Ctrl+Z / Ctrl+Y)."""

    def __init__(self, get_state, set_state) -> None:
        self._get, self._set = get_state, set_state
        self._undo: list = []
        self._redo: list = []
        self._restoring = False

    def push(self) -> None:
        if self._restoring:
            return
        self._undo.append(copy.deepcopy(self._get()))
        del self._undo[:-UNDO_LIMIT]
        self._redo.clear()

    def clear(self) -> None:
        self._undo.clear()
        self._redo.clear()

    def undo(self) -> bool:
        return self._step(self._undo, self._redo)

    def redo(self) -> bool:
        return self._step(self._redo, self._undo)

    def _step(self, source: list, target: list) -> bool:
        if not source:
            return False
        self._restoring = True
        try:
            target.append(copy.deepcopy(self._get()))
            self._set(source.pop())
        finally:
            self._restoring = False
        return True


class DayView(QWidget):
    _instance: Optional['DayView'] = None

    @classmethod
    def get_or_create(cls) -> 'DayView':
        if cls._instance is None:
            cls._instance = DayView()
        return cls._instance

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle('授渔 · 日视图')
        self.setAttribute(Qt.WA_DeleteOnClose, False)
        self.resize(1400, 860)
        self.plan_date = _today()
        self._shown_once = False

        self._day_version = ''
        self._day_baseline: List[tuple] = []
        self._original_in_progress: Optional[str] = None
        self._backlog_version = ''
        self._backlog_baseline = ([], [])
        self._notes_total = 0

        self.day_model = TaskTableModel(DAY_COLUMNS, self, before_change=lambda: self._day_undo.push())
        self.work_model = TaskTableModel(BACKLOG_COLUMNS, self, category=TaskCategory.WORK,
                                         before_change=lambda: self._backlog_undo.push())
        self.life_model = TaskTableModel(BACKLOG_COLUMNS, self, category=TaskCategory.LIFE,
                                         before_change=lambda: self._backlog_undo.push())
        self._day_undo = _Undo(lambda: self.day_model.tasks, self._restore_day)
        self._backlog_undo = _Undo(lambda: (self.work_model.tasks, self.life_model.tasks), self._restore_backlog)
        for model in (self.day_model, self.work_model, self.life_model):
            model.edited.connect(self._update_dirty)
            model.modelReset.connect(self._update_dirty)

        self._build_ui()
        self._install_shortcuts()

    # ---------- ui ----------

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(8)

        bar = QHBoxLayout()
        self.prev_btn = QPushButton('◀')
        self.prev_btn.setToolTip('前一天（Ctrl+PageUp）')
        self.prev_btn.clicked.connect(lambda: self.go_to(_shift(self.plan_date, -1)))
        self.next_btn = QPushButton('▶')
        self.next_btn.setToolTip('后一天（Ctrl+PageDown）')
        self.next_btn.clicked.connect(lambda: self.go_to(_shift(self.plan_date, 1)))
        for btn in (self.prev_btn, self.next_btn):
            btn.setFixedWidth(44)
            btn.setFocusPolicy(Qt.NoFocus)
        self.date_label = QLabel('')
        self.date_label.setObjectName('TitleLabel')
        self.today_btn = QPushButton('今天')
        self.today_btn.setFocusPolicy(Qt.NoFocus)
        self.today_btn.clicked.connect(lambda: self.go_to(_today()))

        self.pick_btn = QToolButton()
        self.pick_btn.setText('选日期…')
        self.pick_btn.setPopupMode(QToolButton.InstantPopup)
        self.pick_btn.setFocusPolicy(Qt.NoFocus)
        menu = QMenu(self.pick_btn)
        self.calendar = QCalendarWidget()
        self.calendar.setGridVisible(False)
        self.calendar.clicked.connect(self._on_calendar_picked)
        action = QWidgetAction(menu)
        action.setDefaultWidget(self.calendar)
        menu.addAction(action)
        self.pick_btn.setMenu(menu)
        self._calendar_menu = menu

        self.dirty_label = QLabel('● 未保存')
        self.dirty_label.setStyleSheet(f'color: {SUSPENDED_COLOR_HEX};')
        self.dirty_label.hide()
        self.save_btn = QPushButton('保存 Ctrl+S')
        self.save_btn.setObjectName('PrimaryButton')
        self.save_btn.setFocusPolicy(Qt.NoFocus)
        self.save_btn.clicked.connect(self.save)

        bar.addWidget(self.prev_btn)
        bar.addWidget(self.date_label)
        bar.addWidget(self.next_btn)
        bar.addSpacing(12)
        bar.addWidget(self.today_btn)
        bar.addWidget(self.pick_btn)
        bar.addStretch(1)
        bar.addWidget(self.dirty_label)
        bar.addWidget(self.save_btn)
        root.addLayout(bar)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_day_page(), '当天')
        self.tabs.addTab(self._build_backlog_page(), 'Backlog')
        root.addWidget(self.tabs, 1)

        hint = QLabel('Ctrl+Enter 新增 · Delete 删除 · Alt+↑↓ 移动 · Ctrl+Z / Ctrl+Y 撤销重做 · '
                      'Ctrl+C 复制选中格子 · Ctrl+PageUp / PageDown 换一天 · Ctrl+S 保存')
        hint.setObjectName('HintLabel')
        root.addWidget(hint)

    def _build_day_page(self) -> QWidget:
        page = QWidget()
        layout = QHBoxLayout(page)
        layout.setContentsMargins(0, 8, 0, 0)
        split = QSplitter(Qt.Horizontal)

        left = QSplitter(Qt.Vertical)
        tasks_box = QWidget()
        tasks_layout = QVBoxLayout(tasks_box)
        tasks_layout.setContentsMargins(0, 0, 0, 0)
        self.tasks_title = QLabel('任务')
        self.tasks_title.setObjectName('SubtitleLabel')
        self.day_table = TaskTableView(self.day_model)
        add_btn = QPushButton('＋ 新增任务')
        add_btn.setFocusPolicy(Qt.NoFocus)
        add_btn.clicked.connect(self.day_table.add_row)
        tasks_layout.addWidget(self.tasks_title)
        tasks_layout.addWidget(self.day_table, 1)
        tasks_layout.addWidget(add_btn, 0, Qt.AlignLeft)
        left.addWidget(tasks_box)

        pomo_box = QWidget()
        pomo_layout = QVBoxLayout(pomo_box)
        pomo_layout.setContentsMargins(0, 0, 0, 0)
        self.pomodoro_title = QLabel('番茄（只读）')
        self.pomodoro_title.setObjectName('SubtitleLabel')
        self.pomodoro_table = QTableWidget(0, 5)
        self.pomodoro_table.setHorizontalHeaderLabels(['时间', '分钟', '任务', '模式', '备注'])
        self.pomodoro_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.pomodoro_table.verticalHeader().hide()
        self.pomodoro_table.horizontalHeader().setStretchLastSection(True)
        self.pomodoro_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.pomodoro_table.setStyleSheet(self.day_table.styleSheet())
        pomo_layout.addWidget(self.pomodoro_title)
        pomo_layout.addWidget(self.pomodoro_table, 1)
        left.addWidget(pomo_box)
        left.setSizes([560, 220])
        split.addWidget(left)

        notes_box = QWidget()
        notes_layout = QVBoxLayout(notes_box)
        notes_layout.setContentsMargins(8, 0, 0, 0)
        self.notes_title = QLabel('记录')
        self.notes_title.setObjectName('SubtitleLabel')
        self.notes_view = NoteListView()
        self.notes_view.allow_open_day = False
        self.notes_view.copied.connect(self._on_note_copied)
        self.notes_view.notes_changed.connect(self._on_notes_changed)
        self.notes_view.older_requested.connect(self._load_older_notes)
        self.note_input = QLineEdit()
        self.note_input.setPlaceholderText('＋ 加一条记录，回车保存')
        self.note_input.returnPressed.connect(self._add_note)
        notes_layout.addWidget(self.notes_title)
        notes_layout.addWidget(self.notes_view, 1)
        notes_layout.addWidget(self.note_input)
        split.addWidget(notes_box)
        split.setSizes([900, 460])
        layout.addWidget(split)
        return page

    def _build_backlog_page(self) -> QWidget:
        page = QWidget()
        layout = QHBoxLayout(page)
        layout.setContentsMargins(0, 8, 0, 0)
        self.work_table = TaskTableView(self.work_model)
        self.life_table = TaskTableView(self.life_model)
        for title, table, target, arrow in (
            ('工作', self.work_table, TaskCategory.LIFE, '移到生活 →'),
            ('生活', self.life_table, TaskCategory.WORK, '← 移到工作'),
        ):
            box = QWidget()
            col = QVBoxLayout(box)
            col.setContentsMargins(0, 0, 0, 0)
            label = QLabel(title)
            label.setObjectName('SubtitleLabel')
            buttons = QHBoxLayout()
            add_btn = QPushButton('＋ 新增')
            add_btn.setFocusPolicy(Qt.NoFocus)
            add_btn.clicked.connect(table.add_row)
            move_btn = QPushButton(arrow)
            move_btn.setFocusPolicy(Qt.NoFocus)
            move_btn.clicked.connect(lambda _c=False, t=table, c=target: self.move_backlog_task(t, c))
            buttons.addWidget(add_btn)
            buttons.addWidget(move_btn)
            buttons.addStretch(1)
            col.addWidget(label)
            col.addWidget(table, 1)
            col.addLayout(buttons)
            layout.addWidget(box, 1)
        return page

    def _install_shortcuts(self) -> None:
        for key, slot in (
            ('Ctrl+S', self.save),
            ('Ctrl+PgUp', lambda: self.go_to(_shift(self.plan_date, -1))),
            ('Ctrl+PgDown', lambda: self.go_to(_shift(self.plan_date, 1))),
            ('Ctrl+Z', self.undo),
            ('Ctrl+Y', self.redo),
            ('Ctrl+Shift+Z', self.redo),
        ):
            QShortcut(QKeySequence(key), self).activated.connect(slot)

    # ---------- open / load ----------

    def open_date(self, plan_date: Optional[str] = None) -> None:
        target = plan_date or _today()
        if not self.isVisible() or target != self.plan_date:
            if self.isVisible() and not self._confirm_leave():
                self._bring_to_front()
                return
            self.load(target)
        if not self._shown_once:
            self._shown_once = True
            screen = QGuiApplication.screenAt(QCursor.pos()) or QGuiApplication.primaryScreen()
            if screen is not None:
                self.setGeometry(screen.availableGeometry())
            self.showMaximized()
        else:
            self.showNormal() if self.isMinimized() else self.show()
        self._bring_to_front()

    def _bring_to_front(self) -> None:
        self.raise_()
        self.activateWindow()

    def load(self, plan_date: str) -> None:
        """Read one day (+ the backlog); unsaved edits are dropped."""
        self.plan_date = plan_date
        try:
            tasks, self._day_version = task_store.load_day_for_edit(plan_date)
            work, life, self._backlog_version = task_store.load_backlog_for_edit()
            pomodoros = task_store.load_pomodoros(plan_date)
            notes = note_store.day_page(plan_date)
            self._notes_total = note_store.count_day(plan_date)
        except Exception:
            logging.exception(f'failed to load day view for {plan_date}')
            QMessageBox.warning(self, '日视图', f'读取 {plan_date} 失败，详情见日志。')
            return
        self._set_day_tasks(tasks)
        self.work_model.set_tasks(work)
        self.life_model.set_tasks(life)
        self._backlog_baseline = (snapshot_tasks(work), snapshot_tasks(life))
        self._day_undo.clear()
        self._backlog_undo.clear()
        self._show_pomodoros(pomodoros)
        self.notes_view.set_notes(notes, [])
        self._update_notes_title()
        self.date_label.setText(_date_title(plan_date))
        self.calendar.setSelectedDate(QDate.fromString(plan_date, 'yyyy-MM-dd'))
        self.setWindowTitle(f'授渔 · 日视图 · {plan_date}')
        self._update_dirty()

    def _set_day_tasks(self, tasks: List[PlanTask]) -> None:
        self.day_model.set_tasks(tasks)
        self._day_baseline = snapshot_tasks(tasks)
        in_progress = next((t for t in tasks if t.status == TaskStatus.IN_PROGRESS), None)
        self._original_in_progress = in_progress.text if in_progress else None

    def _show_pomodoros(self, records: List[PomodoroRecord]) -> None:
        table = self.pomodoro_table
        table.setRowCount(len(records))
        grey = QColor(SUBTEXT_COLOR_HEX)
        work = 0
        for row, r in enumerate(records):
            planning = r.kind == task_store.KIND_PLANNING
            work += 0 if planning else 1
            notes = []
            if planning:
                notes.append('规划')
            if r.ended_early:
                notes.append('提前结束')
            mode = {'deep': '深度', 'classic': '经典'}.get(r.mode or '', r.mode or '')
            values = [
                f'{_hhmm(r.started_at)}–{_hhmm(r.ended_at)}',
                str(r.actual_minutes or ''),
                r.task_text or ('（规划）' if planning else ''),
                mode,
                '、'.join(notes),
            ]
            for col, value in enumerate(values):
                item = QTableWidgetItem(value)
                if planning:
                    item.setForeground(grey)
                table.setItem(row, col, item)
        table.resizeColumnToContents(0)
        self.pomodoro_title.setText(f'番茄（只读）· {work} 个')

    def _update_notes_title(self) -> None:
        shown = len(self.notes_view.note_model.notes())
        total = max(shown, self._notes_total)
        more = f'，已加载 {shown} 条，往上滚动加载更多' if shown < total else ''
        self.notes_title.setText(f'记录（{total}{more}）')

    def _load_older_notes(self) -> None:
        notes = self.notes_view.note_model.notes()
        if not notes or len(notes) >= self._notes_total:
            return
        try:
            older = note_store.day_page(self.plan_date, notes[0].id)
        except Exception:
            logging.exception('failed to load older notes of the day')
            return
        self.notes_view.prepend_notes(older)
        self._update_notes_title()

    def _on_notes_changed(self) -> None:
        self._notes_total = note_store.count_day(self.plan_date)
        self._update_notes_title()

    # ---------- dirty state ----------

    def day_dirty(self) -> bool:
        return snapshot_tasks(self.day_model.tasks) != self._day_baseline

    def backlog_dirty(self) -> bool:
        return (snapshot_tasks(self.work_model.tasks), snapshot_tasks(self.life_model.tasks)) != self._backlog_baseline

    def is_dirty(self) -> bool:
        return self.day_dirty() or self.backlog_dirty()

    def _update_dirty(self) -> None:
        self.dirty_label.setVisible(self.is_dirty())
        tasks = [t for t in self.day_model.tasks if (t.text or '').strip()]
        done = sum(1 for t in tasks if t.status == TaskStatus.DONE)
        self.tasks_title.setText(f'任务 · {done}/{len(tasks)} 完成')

    def _confirm_leave(self) -> bool:
        """True when it's fine to drop / move away from the current edits."""
        if not self.is_dirty():
            return True
        box = QMessageBox(self)
        box.setWindowTitle('日视图')
        box.setText(f'{self.plan_date} 有未保存的改动。')
        save = box.addButton('保存', QMessageBox.AcceptRole)
        discard = box.addButton('不保存', QMessageBox.DestructiveRole)
        box.addButton('取消', QMessageBox.RejectRole)
        box.exec()
        clicked = box.clickedButton()
        if clicked is save:
            return self.save()
        return clicked is discard

    def go_to(self, plan_date: str) -> None:
        if plan_date == self.plan_date:
            return
        if not self._confirm_leave():
            return
        self.load(plan_date)

    def _on_calendar_picked(self, qdate: QDate) -> None:
        self._calendar_menu.hide()
        self.go_to(qdate.toString('yyyy-MM-dd'))

    # ---------- undo ----------

    def _restore_day(self, tasks) -> None:
        self.day_model.set_tasks(tasks)

    def _restore_backlog(self, state) -> None:
        work, life = state
        self.work_model.set_tasks(work)
        self.life_model.set_tasks(life)

    def _active_undo(self) -> _Undo:
        return self._backlog_undo if self.tabs.currentIndex() == 1 else self._day_undo

    def undo(self) -> None:
        if self._active_undo().undo():
            self._update_dirty()

    def redo(self) -> None:
        if self._active_undo().redo():
            self._update_dirty()

    # ---------- backlog ----------

    def move_backlog_task(self, table: TaskTableView, target: TaskCategory) -> None:
        row = table.current_row()
        if row < 0:
            return
        self._backlog_undo.push()
        task = table.task_model.take_task(row, record_undo=False)
        task.category = target
        dest_table = self.life_table if target == TaskCategory.LIFE else self.work_table
        dest = dest_table.task_model
        dest.beginInsertRows(QModelIndex(), len(dest.tasks), len(dest.tasks))
        dest.tasks.append(task)
        dest.endInsertRows()
        dest_table.setCurrentIndex(dest.index(len(dest.tasks) - 1, 0))
        self._update_dirty()

    # ---------- save ----------

    def save(self) -> bool:
        """Save whatever changed. True when nothing is left unsaved."""
        self._commit_open_editor()
        from shouyu.service import task_import

        if task_import.is_running():
            QMessageBox.information(self, '日视图', '正在从 Excel 导入历史任务，导入完成后再保存。')
            return False
        try:
            if self.day_dirty() and not self._save_day():
                return False
            if self.backlog_dirty() and not self._save_backlog():
                return False
        except Exception:
            logging.exception('day view save failed')
            QMessageBox.warning(self, '日视图', '保存失败，详情见日志。改动还在窗口里，可以再试一次。')
            return False
        self._update_dirty()
        return True

    def _commit_open_editor(self) -> None:
        """Ctrl+S while a cell is still being edited: keep that edit."""
        focus = QApplication.focusWidget()
        if focus is None:
            return
        for table in (self.day_table, self.work_table, self.life_table):
            if table.isAncestorOf(focus) and table.state() == QAbstractItemView.EditingState:
                table.commitData(focus)
                table.closeEditor(focus, QAbstractItemDelegate.NoHint)
                return

    def _save_day(self, check_version: bool = True) -> bool:
        tasks = self.day_model.tasks
        try:
            task_store.save_plan(PlanSaveRequest(
                plan_date=self.plan_date,
                today_tasks=tasks,
                original_in_progress=self._original_in_progress,
                expected_day_version=self._day_version if check_version else None,
            ))
        except ConflictError:
            choice = self._ask_conflict(f'{self.plan_date} 的任务')
            if choice == 'reload':
                fresh, self._day_version = task_store.load_day_for_edit(self.plan_date)
                self._set_day_tasks(fresh)
                self._day_undo.clear()
                return True
            if choice == 'force':
                return self._save_day(check_version=False)
            return False
        self._day_version = task_store.day_version(self.plan_date)
        self._day_baseline = snapshot_tasks(tasks)
        in_progress = next((t for t in tasks if t.status == TaskStatus.IN_PROGRESS), None)
        self._original_in_progress = in_progress.text if in_progress else None
        if self.plan_date == _today():
            self._sync_pomodoro_task(in_progress)
        return True

    def _save_backlog(self, check_version: bool = True) -> bool:
        work, life = self.work_model.tasks, self.life_model.tasks
        try:
            task_store.save_plan(PlanSaveRequest(
                plan_date=self.plan_date,
                today_tasks=None,
                work_tasks=work,
                life_tasks=life,
                expected_backlog_version=self._backlog_version if check_version else None,
            ))
        except ConflictError:
            choice = self._ask_conflict('Backlog')
            if choice == 'reload':
                work, life, self._backlog_version = task_store.load_backlog_for_edit()
                self.work_model.set_tasks(work)
                self.life_model.set_tasks(life)
                self._backlog_baseline = (snapshot_tasks(work), snapshot_tasks(life))
                self._backlog_undo.clear()
                return True
            if choice == 'force':
                return self._save_backlog(check_version=False)
            return False
        self._backlog_version = task_store.backlog_version()
        self._backlog_baseline = (snapshot_tasks(work), snapshot_tasks(life))
        return True

    def _ask_conflict(self, what: str) -> str:
        box = QMessageBox(self)
        box.setWindowTitle('日视图')
        box.setIcon(QMessageBox.Warning)
        box.setText(f'{what}已经在别的窗口改过。')
        box.setInformativeText('重新加载：丢掉日视图里的改动，显示最新的内容。\n仍然保存：用日视图里的内容覆盖。')
        reload_btn = box.addButton('重新加载', QMessageBox.AcceptRole)
        force_btn = box.addButton('仍然保存', QMessageBox.DestructiveRole)
        box.addButton('取消', QMessageBox.RejectRole)
        box.exec()
        clicked = box.clickedButton()
        if clicked is reload_btn:
            return 'reload'
        if clicked is force_btn:
            return 'force'
        return 'cancel'

    @staticmethod
    def _sync_pomodoro_task(in_progress: Optional[PlanTask]) -> None:
        try:
            from shouyu.service.pomodoro import PomodoroService

            PomodoroService.instance().set_current_task(in_progress.text if in_progress is not None else '')
        except Exception:
            logging.exception('failed to sync pomodoro task label')

    # ---------- notes ----------

    def _add_note(self) -> None:
        text = self.note_input.text().strip()
        if not text:
            return
        try:
            note = note_store.add_manual(self.plan_date, text)
            self.notes_view.note_model.append_note(note)
            self.notes_view.select_last()
            self._notes_total += 1
        except Exception:
            logging.exception('failed to add a manual note')
            QMessageBox.warning(self, '日视图', '记录没有保存，详情见日志。')
            return
        self.note_input.clear()
        self._update_notes_title()

    def _on_note_copied(self, note) -> None:
        try:
            from shouyu.view.msgbox import MessageBox

            preview = '截图' if note.kind == note_store.KIND_IMAGE else (note.text or '').strip()[:60]
            MessageBox.pop_up_message(title='已复制', msg=preview)
        except Exception:
            logging.exception('failed to show copied toast')

    # ---------- close ----------

    def closeEvent(self, event) -> None:
        if not self._confirm_leave():
            event.ignore()
            return
        # Next open re-reads the database instead of showing stale edits.
        self.load(self.plan_date)
        event.accept()
