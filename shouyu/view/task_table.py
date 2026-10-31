"""Editable task table shared by the day view's day page and Backlog page
(design §4.5, DAY-04 / DAY-05 / DAY-07 / DAY-10).

The model edits a plain `List[PlanTask]` in place; the owner keeps the
undo snapshots (`before_change` is called before every mutation) and saves
the list through `task_store.save_plan()`.
"""
from __future__ import annotations

import time
from datetime import datetime
from typing import Callable, List, Optional

from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt, Signal
from PySide6.QtGui import QColor, QFont, QGuiApplication, QKeySequence
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QHeaderView,
    QMessageBox,
    QSpinBox,
    QStyledItemDelegate,
    QTableView,
    QWidget,
)

from shouyu.service.plan import PlanTask, TaskCategory, TaskPriority, TaskStatus
from shouyu.view.styles import (
    DONE_COLOR_HEX,
    IN_PROGRESS_COLOR_HEX,
    PENDING_COLOR_HEX,
    SUBTEXT_COLOR_HEX,
    SUSPENDED_COLOR_HEX,
    TEXT_COLOR_HEX,
)

COL_TEXT, COL_STATUS, COL_PRIORITY, COL_DURATION, COL_CATEGORY, COL_REFLECTION, COL_SHELVED = range(7)

DAY_COLUMNS = [COL_TEXT, COL_STATUS, COL_PRIORITY, COL_DURATION, COL_CATEGORY, COL_REFLECTION]
BACKLOG_COLUMNS = [COL_TEXT, COL_PRIORITY, COL_DURATION, COL_SHELVED]

HEADERS = {
    COL_TEXT: '任务',
    COL_STATUS: '状态',
    COL_PRIORITY: '优先级',
    COL_DURATION: '预估（分钟）',
    COL_CATEGORY: '类别',
    COL_REFLECTION: '反思',
    COL_SHELVED: '已搁置',
}

STATUS_LABELS = [
    (TaskStatus.PENDING, '待办'),
    (TaskStatus.IN_PROGRESS, '进行中'),
    (TaskStatus.SUSPENDED, '搁置'),
    (TaskStatus.DONE, '完成'),
]
PRIORITY_LABELS = [
    (TaskPriority.NONE, '—'),
    (TaskPriority.P1, 'P1'),
    (TaskPriority.P2, 'P2'),
    (TaskPriority.P3, 'P3'),
]
CATEGORY_LABELS = [
    (TaskCategory.WORK, '工作'),
    (TaskCategory.LIFE, '生活'),
]
_STATUS_COLOR = {
    TaskStatus.PENDING: TEXT_COLOR_HEX,
    TaskStatus.IN_PROGRESS: IN_PROGRESS_COLOR_HEX,
    TaskStatus.SUSPENDED: SUSPENDED_COLOR_HEX,
    TaskStatus.DONE: DONE_COLOR_HEX,
}
_CHOICES = {COL_STATUS: STATUS_LABELS, COL_PRIORITY: PRIORITY_LABELS, COL_CATEGORY: CATEGORY_LABELS}


def _label(choices, value) -> str:
    return next((label for v, label in choices if v == value), '')


def shelved_days(created_date: str, today: Optional[str] = None) -> Optional[int]:
    try:
        entered = datetime.strptime((created_date or '').strip(), '%Y-%m-%d')
        now = datetime.strptime(today or time.strftime('%Y-%m-%d'), '%Y-%m-%d')
    except ValueError:
        return None
    return max(0, (now - entered).days)


def snapshot_tasks(tasks: List[PlanTask]) -> List[tuple]:
    """Structural comparison for unsaved-change checks; empty rows are
    dropped (they are not saved either), same as the habit dialog."""
    return [
        (
            (t.text or '').strip(), t.status, int(t.duration_minutes or 0), t.priority,
            (t.reflection or '').strip(), t.category,
        )
        for t in tasks if (t.text or '').strip()
    ]


class TaskTableModel(QAbstractTableModel):
    """Columns are picked per page (`DAY_COLUMNS` / `BACKLOG_COLUMNS`)."""

    edited = Signal()

    def __init__(self, columns: List[int], parent=None, category: Optional[TaskCategory] = None,
                 before_change: Optional[Callable[[], None]] = None,
                 today_fn: Callable[[], str] = lambda: time.strftime('%Y-%m-%d')) -> None:
        super().__init__(parent)
        self.columns = list(columns)
        self.tasks: List[PlanTask] = []
        # Backlog tables hold one category; new rows get it.
        self.category = category
        self.before_change = before_change or (lambda: None)
        self._today_fn = today_fn

    # ---- structure ----
    def rowCount(self, parent=QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.tasks)

    def columnCount(self, parent=QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.columns)

    def column_kind(self, column: int) -> int:
        return self.columns[column]

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return HEADERS[self.columns[section]]
        return str(section + 1)

    def flags(self, index):
        if not index.isValid():
            return Qt.NoItemFlags
        base = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if self.columns[index.column()] == COL_SHELVED:
            return base
        return base | Qt.ItemIsEditable

    # ---- data ----
    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid() or not (0 <= index.row() < len(self.tasks)):
            return None
        task = self.tasks[index.row()]
        kind = self.columns[index.column()]
        if role in (Qt.DisplayRole, Qt.ToolTipRole):
            if kind == COL_TEXT:
                return task.text
            if kind == COL_STATUS:
                return _label(STATUS_LABELS, task.status)
            if kind == COL_PRIORITY:
                return _label(PRIORITY_LABELS, task.priority)
            if kind == COL_DURATION:
                return str(task.duration_minutes) if task.duration_minutes else ''
            if kind == COL_CATEGORY:
                return _label(CATEGORY_LABELS, task.category)
            if kind == COL_REFLECTION:
                return task.reflection or ''
            if kind == COL_SHELVED:
                days = shelved_days(task.created_date, self._today_fn())
                return '' if days is None else f'{days} 天'
        if role == Qt.EditRole:
            if kind == COL_TEXT:
                return task.text
            if kind == COL_STATUS:
                return task.status
            if kind == COL_PRIORITY:
                return task.priority
            if kind == COL_DURATION:
                return int(task.duration_minutes or 0)
            if kind == COL_CATEGORY:
                return task.category
            if kind == COL_REFLECTION:
                return task.reflection or ''
        if role == Qt.ForegroundRole:
            if kind in (COL_TEXT, COL_STATUS) and COL_STATUS in self.columns:
                return QColor(_STATUS_COLOR.get(task.status, TEXT_COLOR_HEX))
            if kind == COL_SHELVED:
                return QColor(SUBTEXT_COLOR_HEX)
        if role == Qt.FontRole and kind == COL_TEXT and task.status == TaskStatus.DONE:
            font = QFont()
            font.setStrikeOut(True)
            return font
        if role == Qt.TextAlignmentRole and kind in (COL_DURATION, COL_SHELVED):
            return int(Qt.AlignRight | Qt.AlignVCenter)
        return None

    def setData(self, index, value, role=Qt.EditRole) -> bool:
        if role != Qt.EditRole or not index.isValid():
            return False
        task = self.tasks[index.row()]
        kind = self.columns[index.column()]
        if kind == COL_TEXT:
            value = str(value or '').strip()
            if value == task.text:
                return False
            self.before_change()
            task.text = value
        elif kind == COL_STATUS:
            value = TaskStatus(value)
            if value == task.status:
                return False
            self.before_change()
            if value == TaskStatus.IN_PROGRESS:
                # Only one task is in progress (DAY-05, same as the habit dialog).
                for row, other in enumerate(self.tasks):
                    if other is not task and other.status == TaskStatus.IN_PROGRESS:
                        other.status = TaskStatus.PENDING
                        self._row_changed(row)
            task.status = value
        elif kind == COL_PRIORITY:
            value = TaskPriority(value)
            if value == task.priority:
                return False
            self.before_change()
            task.priority = value
        elif kind == COL_DURATION:
            value = max(0, int(value or 0))
            if value == int(task.duration_minutes or 0):
                return False
            self.before_change()
            task.duration_minutes = value
        elif kind == COL_CATEGORY:
            value = TaskCategory(value)
            if value == task.category:
                return False
            self.before_change()
            task.category = value
        elif kind == COL_REFLECTION:
            value = str(value or '')
            if value == (task.reflection or ''):
                return False
            self.before_change()
            task.reflection = value
        else:
            return False
        self._row_changed(index.row())
        self.edited.emit()
        return True

    def _row_changed(self, row: int) -> None:
        self.dataChanged.emit(self.index(row, 0), self.index(row, len(self.columns) - 1))

    # ---- whole-list operations ----
    def set_tasks(self, tasks: List[PlanTask]) -> None:
        self.beginResetModel()
        self.tasks = tasks
        self.endResetModel()

    def insert_task(self, row: int, task: Optional[PlanTask] = None) -> int:
        row = max(0, min(row, len(self.tasks)))
        task = task or PlanTask('', category=self.category or TaskCategory.WORK)
        if self.category is not None:
            task.category = self.category
        self.before_change()
        self.beginInsertRows(QModelIndex(), row, row)
        self.tasks.insert(row, task)
        self.endInsertRows()
        self.edited.emit()
        return row

    def take_task(self, row: int, record_undo: bool = True) -> Optional[PlanTask]:
        if not (0 <= row < len(self.tasks)):
            return None
        if record_undo:
            self.before_change()
        self.beginRemoveRows(QModelIndex(), row, row)
        task = self.tasks.pop(row)
        self.endRemoveRows()
        self.edited.emit()
        return task

    def move_task(self, row: int, delta: int) -> int:
        target = row + delta
        if not (0 <= row < len(self.tasks)) or not (0 <= target < len(self.tasks)):
            return row
        self.before_change()
        self.beginMoveRows(QModelIndex(), row, row, QModelIndex(), target + (1 if delta > 0 else 0))
        self.tasks.insert(target, self.tasks.pop(row))
        self.endMoveRows()
        self.edited.emit()
        return target


class TaskItemDelegate(QStyledItemDelegate):
    """Combo boxes for status / priority / category, a spin box for minutes."""

    def createEditor(self, parent, option, index):
        model: TaskTableModel = index.model()
        kind = model.column_kind(index.column())
        if kind in _CHOICES:
            combo = QComboBox(parent)
            for value, label in _CHOICES[kind]:
                combo.addItem(label, value)
            # Commit as soon as a choice is made, no extra Enter needed.
            combo.activated.connect(lambda _i, c=combo: (self.commitData.emit(c), self.closeEditor.emit(c)))
            return combo
        if kind == COL_DURATION:
            spin = QSpinBox(parent)
            spin.setRange(0, 24 * 60)
            spin.setSingleStep(5)
            spin.setSuffix(' 分钟')
            return spin
        return super().createEditor(parent, option, index)

    def setEditorData(self, editor, index):
        if isinstance(editor, QComboBox):
            value = index.data(Qt.EditRole)
            pos = editor.findData(value)
            editor.setCurrentIndex(max(0, pos))
            return
        if isinstance(editor, QSpinBox):
            editor.setValue(int(index.data(Qt.EditRole) or 0))
            return
        super().setEditorData(editor, index)

    def setModelData(self, editor, model, index):
        if isinstance(editor, QComboBox):
            model.setData(index, editor.currentData(), Qt.EditRole)
            return
        if isinstance(editor, QSpinBox):
            editor.interpretText()
            model.setData(index, editor.value(), Qt.EditRole)
            return
        super().setModelData(editor, model, index)


def selection_tsv(view: QTableView) -> str:
    """Selected cells as tab / newline separated text (DAY-10)."""
    indexes = view.selectionModel().selectedIndexes() if view.selectionModel() else []
    if not indexes:
        current = view.currentIndex()
        indexes = [current] if current.isValid() else []
    if not indexes:
        return ''
    rows = sorted({i.row() for i in indexes})
    cols = sorted({i.column() for i in indexes})
    chosen = {(i.row(), i.column()) for i in indexes}
    model = view.model()
    lines = []
    for r in rows:
        cells = []
        for c in cols:
            text = str(model.index(r, c).data(Qt.DisplayRole) or '') if (r, c) in chosen else ''
            cells.append(text.replace('\t', ' ').replace('\r', ' ').replace('\n', ' '))
        lines.append('\t'.join(cells))
    return '\n'.join(lines)


class TaskTableView(QTableView):
    """Keyboard: Ctrl+Enter adds a row below, Delete removes (asks when the
    task has a name), Alt+↑/↓ moves, Ctrl+C copies the selected cells."""

    rows_changed = Signal()

    def __init__(self, model: TaskTableModel, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setModel(model)
        self.task_model = model
        self.setItemDelegate(TaskItemDelegate(self))
        self.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setEditTriggers(
            QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed | QAbstractItemView.AnyKeyPressed
        )
        self.setAlternatingRowColors(False)
        self.setWordWrap(False)
        self.verticalHeader().setDefaultSectionSize(30)
        self.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        header = self.horizontalHeader()
        header.setStretchLastSection(True)
        header.setHighlightSections(False)
        self.setStyleSheet(
            'QTableView { background-color: #232323; gridline-color: #333333; border: 1px solid #3A3A3A;'
            ' border-radius: 6px; selection-background-color: rgba(15, 98, 254, 0.35); }'
            'QHeaderView::section { background-color: #2B2B2B; color: #9A9A9A; border: none;'
            ' border-right: 1px solid #333333; padding: 4px 6px; }'
            'QTableCornerButton::section { background-color: #2B2B2B; border: none; }'
        )
        self.apply_column_widths()

    def keyPressEvent(self, event) -> None:
        key, mods = event.key(), event.modifiers()
        if key in (Qt.Key_Return, Qt.Key_Enter) and mods & Qt.ControlModifier:
            self.add_row()
        elif key == Qt.Key_Delete and not mods:
            self.delete_current()
        elif key in (Qt.Key_Up, Qt.Key_Down) and mods & Qt.AltModifier:
            self.move_current(-1 if key == Qt.Key_Up else 1)
        elif event.matches(QKeySequence.Copy):
            self.copy_selection()
        else:
            super().keyPressEvent(event)
            return
        event.accept()

    def apply_column_widths(self) -> None:
        widths = {COL_TEXT: 380, COL_STATUS: 80, COL_PRIORITY: 70, COL_DURATION: 110, COL_CATEGORY: 70,
                  COL_SHELVED: 80}
        for col, kind in enumerate(self.task_model.columns):
            if kind in widths:
                self.setColumnWidth(col, widths[kind])

    def current_row(self) -> int:
        index = self.currentIndex()
        return index.row() if index.isValid() else -1

    def add_row(self) -> None:
        row = self.current_row()
        row = self.task_model.insert_task(row + 1 if row >= 0 else len(self.task_model.tasks))
        index = self.task_model.index(row, 0)
        self.setCurrentIndex(index)
        self.edit(index)
        self.rows_changed.emit()

    def delete_current(self) -> None:
        row = self.current_row()
        if row < 0:
            return
        task = self.task_model.tasks[row]
        if (task.text or '').strip():
            answer = QMessageBox.question(self, '删除任务', f'删除"{task.text}"？')
            if answer != QMessageBox.Yes:
                return
        self.task_model.take_task(row)
        if self.task_model.tasks:
            self.setCurrentIndex(self.task_model.index(min(row, len(self.task_model.tasks) - 1), 0))
        self.rows_changed.emit()

    def move_current(self, delta: int) -> None:
        index = self.currentIndex()
        if not index.isValid():
            return
        row = self.task_model.move_task(index.row(), delta)
        self.setCurrentIndex(self.task_model.index(row, index.column()))

    def copy_selection(self) -> None:
        text = selection_tsv(self)
        if text:
            QGuiApplication.clipboard().setText(text)
