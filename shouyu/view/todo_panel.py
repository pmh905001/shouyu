"""PySide6 task panel.

Triggered by `ctrl+alt+\``. Shares interaction patterns with HabitDialog so
users only have to learn the keys once. Saves run on a background daemon
thread so closing feels instant.

Keys:
    Up/Down                navigate tasks
    Alt+Up / Alt+Down      reorder (or just drag)
    F2 / Enter             start editing the current task
    Enter (in editor)      commit and start editing next task (Excel-like)
    Ctrl+Plus              add a new task below current
    Ctrl+Minus             delete current task (no confirmation; use Ctrl+Z to restore)
    Ctrl+Z / Ctrl+Y        undo / redo any task-level edit
    Ctrl+P                 start a pomodoro on the current task (focus mode)
    Right-click            change status, priority, duration, reflection…
    Esc                    save & close
"""
from __future__ import annotations

import logging
import threading
import time
from typing import List, Optional

from PySide6.QtCore import QPoint, Qt, QTimer, Signal
from PySide6.QtGui import QColor, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QAbstractItemDelegate,
    QAbstractItemView,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QProgressBar,
    QPushButton,
    QStackedLayout,
    QVBoxLayout,
    QWidget,
)

from shouyu.config import Config
from shouyu.service.plan import (
    DEFAULT_PLAN_TASKS,
    PlanTask,
    TaskPriority,
    TaskStatus,
)
from shouyu.view.duration_dialog import DurationPickerDialog
from shouyu.view.styles import (
    ACCENT_COLOR_HEX,
    DONE_COLOR_HEX,
    IN_PROGRESS_COLOR_HEX,
    PENDING_COLOR_HEX,
    SUBTEXT_COLOR_HEX,
)


_STATUS_GLYPH = {
    TaskStatus.PENDING: "○",
    TaskStatus.IN_PROGRESS: "▶",
    TaskStatus.DONE: "✓",
}

_STATUS_COLORS = {
    TaskStatus.PENDING: PENDING_COLOR_HEX,
    TaskStatus.IN_PROGRESS: IN_PROGRESS_COLOR_HEX,
    TaskStatus.DONE: DONE_COLOR_HEX,
}

_PRIORITY_BADGE = {
    TaskPriority.P1: "🔴 P1",
    TaskPriority.P2: "🟡 P2",
    TaskPriority.P3: "⚪ P3",
}

_TASK_ROLE = Qt.UserRole + 1


def _format_item(task: PlanTask) -> str:
    glyph = _STATUS_GLYPH.get(task.status, "○")
    parts = [glyph, "  ", task.text or ""]
    badge = _PRIORITY_BADGE.get(task.priority)
    if badge:
        parts.append(f"   {badge}")
    if task.duration_minutes:
        parts.append(f"   ⏱ {task.duration_minutes}m")
    if task.status == TaskStatus.IN_PROGRESS:
        parts.append("   🎯 重点")
    if (task.reflection or "").strip():
        parts.append("   📝")
    return "".join(parts)


def _apply_item_style(item: QListWidgetItem, status: TaskStatus) -> None:
    color = QColor(_STATUS_COLORS.get(status, PENDING_COLOR_HEX))
    item.setForeground(color)
    font = item.font()
    font.setBold(status == TaskStatus.IN_PROGRESS)
    font.setStrikeOut(status == TaskStatus.DONE)
    base_size = max(font.pointSize(), 10)
    font.setPointSize(base_size + 2 if status == TaskStatus.IN_PROGRESS else base_size)
    item.setFont(font)


_RETRY_DELAYS_S = [1, 2, 5, 10, 15]


def _persist_plan_in_background(
    tasks_snapshot: List[PlanTask], original_in_progress: Optional[str]
) -> None:
    """Background save with auto-retry on PermissionError (Excel-locked).
    See habit_dialog._persist_plan_in_background for the rationale."""

    def _stage(excel) -> None:
        non_empty = [t for t in tasks_snapshot if (t.text or "").strip()]
        plan = excel.plan_service()
        plan.write_plan_tasks(non_empty)
        new_in_progress = next(
            (t for t in non_empty if t.status == TaskStatus.IN_PROGRESS), None
        )
        if new_in_progress is not None and new_in_progress.text != original_in_progress:
            plan.switch_in_progress(new_in_progress)
        excel.mark_changed()

    def _toast(level, title, msg):
        try:
            from shouyu.view.msgbox import MessageBox, MessageType

            MessageBox.pop_up_message(
                title=title,
                msg=msg,
                level=MessageType.ERROR if level == 'error' else MessageType.SUCCESS,
            )
        except Exception:
            logging.exception("toast failed")

    def _worker():
        from shouyu.service.excel import KbExcel

        excel: Optional[KbExcel] = None
        last_err: Optional[Exception] = None
        notified_retry = False

        for attempt in range(len(_RETRY_DELAYS_S) + 1):
            try:
                if excel is None:
                    excel = KbExcel()
                    _stage(excel)
                excel.force_save()
                if attempt > 0:
                    _toast('ok', "已保存", f"第 {attempt + 1} 次重试成功")
                return
            except PermissionError as e:
                last_err = e
                logging.warning(f"save attempt {attempt + 1} locked: {e}")
            except Exception as e:
                last_err = e
                logging.exception(f"save attempt {attempt + 1} failed")
                break

            if attempt < len(_RETRY_DELAYS_S):
                if not notified_retry:
                    notified_retry = True
                    _toast(
                        'error',
                        "保存失败 · 后台重试中",
                        "Excel 文件被占用（最多 30 秒，请关闭 Excel 后等待）",
                    )
                time.sleep(_RETRY_DELAYS_S[attempt])

        preserved = ''
        if excel is not None:
            try:
                preserved = excel.preserve_unsaved() or ''
            except Exception:
                logging.exception("preserve failed")
        try:
            from shouyu.view.qt_app import QtApp

            extra = (
                f"\n\n改动已保存到备用文件：\n{preserved}"
                if preserved
                else "\n\n（备用文件也写入失败；请手动核对最近的备份。）"
            )
            QtApp.show_save_status(
                'error',
                "任务保存失败",
                f"保存失败：{last_err}{extra}",
            )
        except Exception:
            logging.exception("final notify failed")

    threading.Thread(target=_worker, name="shouyu-save-plan", daemon=True).start()


class TodoPanel(QWidget):
    """A floating window that lets the user manage today's plan tasks."""

    closed = Signal()

    _instance: Optional["TodoPanel"] = None

    @classmethod
    def get_or_create(cls) -> "TodoPanel":
        if cls._instance is None:
            cls._instance = TodoPanel()
        return cls._instance

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("授渔 · 今日任务")
        self.setMinimumSize(680, 520)
        self.setWindowFlag(Qt.WindowStaysOnTopHint, True)

        self._tasks: List[PlanTask] = []
        self._original_in_progress: Optional[str] = None
        self._pending_duration_prompt: set = set()

        # See HabitDialog._undo_stack for the full rationale.
        self._undo_stack: List[List[PlanTask]] = []
        self._redo_stack: List[List[PlanTask]] = []
        self._in_undo_redo = False

        self._build_ui()
        self._install_shortcuts()

    # ---------- public API ----------

    def refresh_from_excel(self) -> None:
        from shouyu.service.excel import KbExcel

        try:
            excel = KbExcel()
            tasks = excel.plan_service().read_plan_tasks()
        except Exception:
            logging.exception("failed to read plan from Excel")
            tasks = []

        if not tasks:
            tasks = [PlanTask(text=t, status=TaskStatus.PENDING) for t in DEFAULT_PLAN_TASKS]

        self._tasks = tasks
        in_progress = next((t for t in tasks if t.status == TaskStatus.IN_PROGRESS), None)
        self._original_in_progress = in_progress.text if in_progress else None
        # Re-loading from Excel invalidates undo history (PlanTask identity
        # changes on disk).
        self._undo_stack.clear()
        self._redo_stack.clear()
        self._reload_list_widget()
        self._update_stats()

    def show_centered(self) -> None:
        self.show()
        self.raise_()
        self.activateWindow()
        screen = self.screen() or self.window().screen()
        if screen is not None:
            geometry = screen.availableGeometry()
            self.move(
                geometry.center().x() - self.width() // 2,
                geometry.center().y() - self.height() // 2,
            )
        self.list_widget.setFocus()

    # ---------- ui ----------

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)

        header = QHBoxLayout()
        title = QLabel("今日任务")
        title.setObjectName("TitleLabel")
        header.addWidget(title)
        header.addStretch(1)
        self.stats_label = QLabel("")
        self.stats_label.setObjectName("SubtitleLabel")
        header.addWidget(self.stats_label)
        layout.addLayout(header)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(5)
        self.progress_bar.setStyleSheet(
            "QProgressBar { background-color: #2B2B2B; border: none; border-radius: 3px; } "
            f"QProgressBar::chunk {{ background-color: {ACCENT_COLOR_HEX}; border-radius: 3px; }}"
        )
        layout.addWidget(self.progress_bar)

        hint = QLabel(
            "↑↓ 选择 ·  F2 / 回车 编辑 ·  Alt+↑↓ 重排 ·  "
            "Ctrl+ + 添加 ·  Ctrl+ − 删除 ·  Ctrl+Z 撤销 / Ctrl+Y 重做 ·  "
            "Ctrl+P 专注 ·  右键 → 状态 / 优先级 / 时长"
        )
        hint.setObjectName("HintLabel")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        self.task_stack_host = QWidget()
        self.task_stack = QStackedLayout(self.task_stack_host)

        self.list_widget = QListWidget()
        self.list_widget.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked
            | QAbstractItemView.EditTrigger.EditKeyPressed
        )
        self.list_widget.setSelectionMode(QAbstractItemView.SingleSelection)
        self.list_widget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.list_widget.setDragDropMode(QAbstractItemView.InternalMove)
        self.list_widget.setMovement(QListWidget.Snap)
        self.list_widget.itemChanged.connect(self._on_item_text_edited)
        self.list_widget.customContextMenuRequested.connect(self._show_context_menu)
        self.list_widget.model().rowsMoved.connect(self._on_rows_moved)
        # Excel-like Enter behavior — see HabitDialog._on_editor_closed for details.
        self.list_widget.itemDelegate().closeEditor.connect(self._on_editor_closed)
        self.task_stack.addWidget(self.list_widget)

        self.empty_label = QLabel("🎉  没有任务了\n按 Ctrl+ + 添加或在上方输入框快速创建")
        self.empty_label.setAlignment(Qt.AlignCenter)
        self.empty_label.setStyleSheet(
            f"color: {SUBTEXT_COLOR_HEX}; font-size: 14px; line-height: 22px;"
        )
        self.task_stack.addWidget(self.empty_label)

        layout.addWidget(self.task_stack_host, stretch=1)

        button_row = QHBoxLayout()
        focus_btn = QPushButton("🍅 专注此项 (Ctrl+P)")
        focus_btn.setAutoDefault(False)
        focus_btn.setDefault(False)
        focus_btn.clicked.connect(self._focus_pomodoro_on_selected)
        button_row.addWidget(focus_btn)

        button_row.addStretch(1)

        close_btn = QPushButton("保存并关闭 (Esc)")
        close_btn.setObjectName("PrimaryButton")
        close_btn.setAutoDefault(False)
        close_btn.setDefault(False)
        close_btn.clicked.connect(self._save_and_close)
        button_row.addWidget(close_btn)

        layout.addLayout(button_row)

    def _install_shortcuts(self) -> None:
        # Window-wide shortcuts (NOT plain Return/Enter — those need to reach the
        # inline editor when one is open).
        #
        # Space-to-cycle-status was removed on purpose: users found it too
        # easy to flip a task's state accidentally. Status changes now go
        # exclusively through the right-click context menu.
        window_bindings = [
            ("F2", self._edit_selected),
            ("Alt+Up", lambda: self._move_selected(-1)),
            ("Alt+Down", lambda: self._move_selected(1)),
            ("Ctrl++", self._add_new_task),
            ("Ctrl+=", self._add_new_task),
            ("Ctrl+Shift+=", self._add_new_task),
            ("Ctrl+-", self._delete_selected),
            ("Ctrl+Minus", self._delete_selected),
            ("Ctrl+P", self._focus_pomodoro_on_selected),
            ("Ctrl+Z", self._undo),
            ("Ctrl+Y", self._redo),
            ("Ctrl+Shift+Z", self._redo),
            ("Esc", self._save_and_close),
            ("Escape", self._save_and_close),
        ]
        for sequence, callback in window_bindings:
            shortcut = QShortcut(QKeySequence(sequence), self)
            shortcut.setContext(Qt.WindowShortcut)
            shortcut.activated.connect(callback)

        # Plain Enter only fires when the list itself has focus (not its editor),
        # so the inline editor can naturally consume Enter to commit.
        for sequence in ("Return", "Enter"):
            sc = QShortcut(QKeySequence(sequence), self.list_widget)
            sc.setContext(Qt.WidgetShortcut)
            sc.activated.connect(self._edit_selected)

    # ---------- list helpers ----------

    def _reload_list_widget(self) -> None:
        self.list_widget.blockSignals(True)
        self.list_widget.clear()
        for task in self._tasks:
            item = QListWidgetItem(_format_item(task))
            flags = item.flags()
            flags |= Qt.ItemIsEditable | Qt.ItemIsDragEnabled
            flags &= ~Qt.ItemIsDropEnabled
            item.setFlags(flags)
            item.setData(_TASK_ROLE, task)
            _apply_item_style(item, task.status)
            reflection = (task.reflection or "").strip()
            if reflection:
                item.setToolTip(f"📝 反思\n\n{reflection}")
            self.list_widget.addItem(item)
        self.list_widget.blockSignals(False)
        if self.list_widget.count() > 0:
            self.list_widget.setCurrentRow(0)
        self._update_empty_state()

    def _update_empty_state(self) -> None:
        if self.list_widget.count() == 0:
            self.task_stack.setCurrentWidget(self.empty_label)
        else:
            self.task_stack.setCurrentWidget(self.list_widget)

    def _selected_index(self) -> int:
        return self.list_widget.currentRow()

    def _refresh_item(self, index: int) -> None:
        if not (0 <= index < len(self._tasks)):
            return
        task = self._tasks[index]
        item = self.list_widget.item(index)
        self.list_widget.blockSignals(True)
        item.setText(_format_item(task))
        _apply_item_style(item, task.status)
        item.setData(_TASK_ROLE, task)
        reflection = (task.reflection or "").strip()
        item.setToolTip(f"📝 反思\n\n{reflection}" if reflection else "")
        self.list_widget.blockSignals(False)

    def _on_item_text_edited(self, item: QListWidgetItem) -> None:
        """Sync data only; cursor advance is driven by `_on_editor_closed`."""
        index = self.list_widget.row(item)
        if 0 <= index < len(self._tasks):
            new_text = self._strip_glyph(item.text())
            if self._tasks[index].text == new_text:
                self._refresh_item(index)
                return
            self._push_undo()
            self._tasks[index].text = new_text
            self._refresh_item(index)
            self._update_stats()

    def _on_editor_closed(self, _editor, hint) -> None:
        if hint == QAbstractItemDelegate.EndEditHint.RevertModelCache:
            return
        QTimer.singleShot(0, self._advance_editing_after_commit)

    def _advance_editing_after_commit(self) -> None:
        index = self._selected_index()

        # See HabitDialog._advance_editing_after_commit for the rationale.
        if 0 <= index < len(self._tasks):
            task = self._tasks[index]
            should_prompt = (
                Config.auto_prompt_duration_for_new_tasks()
                and id(task) in self._pending_duration_prompt
                and (task.text or "").strip()
                and task.duration_minutes <= 0
            )
            self._pending_duration_prompt.discard(id(task))
            if should_prompt:
                value = DurationPickerDialog.get_duration(
                    current=30, task_text=task.text, parent=self
                )
                if value > 0:
                    self._push_undo()
                    task.duration_minutes = value
                    self._refresh_item(index)
                    self._update_stats()

        next_row = index + 1
        if next_row < self.list_widget.count():
            self.list_widget.setCurrentRow(next_row)
            self._edit_row(next_row)
        elif (
            0 <= index < len(self._tasks)
            and (self._tasks[index].text or "").strip()
        ):
            self._push_undo()
            new_task = PlanTask(text="", status=TaskStatus.PENDING)
            self._pending_duration_prompt.add(id(new_task))
            self._tasks.append(new_task)
            new_row = len(self._tasks) - 1
            self._reload_list_widget()
            self.list_widget.setCurrentRow(new_row)
            self._update_stats()
            self._edit_row(new_row)

    @staticmethod
    def _strip_glyph(text: str) -> str:
        text = text or ""
        for glyph in _STATUS_GLYPH.values():
            if text.startswith(glyph):
                text = text[len(glyph):]
                break
        text = text.lstrip()
        first_marker = -1
        for marker in ("   🎯", "   ⏱", "   🔴", "   🟡", "   ⚪", "   📝"):
            idx = text.find(marker)
            if idx >= 0 and (first_marker < 0 or idx < first_marker):
                first_marker = idx
        if first_marker >= 0:
            text = text[:first_marker]
        return text.strip()

    def _edit_selected(self) -> None:
        self._edit_row(self._selected_index())

    def _edit_row(self, index: int) -> None:
        if 0 <= index < self.list_widget.count():
            item = self.list_widget.item(index)
            if item is not None:
                self.list_widget.editItem(item)

    def _on_rows_moved(self, *args) -> None:
        # See HabitDialog._on_rows_moved for why undo capture happens here.
        new_order: List[PlanTask] = []
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            task = item.data(_TASK_ROLE)
            if isinstance(task, PlanTask):
                new_order.append(task)
        if len(new_order) == len(self._tasks) and new_order != self._tasks:
            self._undo_stack.append(self._clone_tasks(self._tasks))
            if len(self._undo_stack) > 100:
                del self._undo_stack[0:len(self._undo_stack) - 100]
            self._redo_stack.clear()
            self._tasks = new_order

    # ---------- undo / redo ----------

    @staticmethod
    def _clone_tasks(tasks: List[PlanTask]) -> List[PlanTask]:
        return [
            PlanTask(
                text=t.text,
                status=t.status,
                row=t.row,
                duration_minutes=t.duration_minutes,
                priority=t.priority,
                reflection=t.reflection,
            )
            for t in tasks
        ]

    def _push_undo(self) -> None:
        if self._in_undo_redo:
            return
        self._undo_stack.append(self._clone_tasks(self._tasks))
        if len(self._undo_stack) > 100:
            del self._undo_stack[0:len(self._undo_stack) - 100]
        self._redo_stack.clear()

    def _undo(self) -> None:
        if not self._undo_stack:
            return
        self._in_undo_redo = True
        try:
            self._redo_stack.append(self._clone_tasks(self._tasks))
            self._tasks = self._undo_stack.pop()
            self._original_in_progress = next(
                (t.text for t in self._tasks if t.status == TaskStatus.IN_PROGRESS),
                None,
            )
            self._reload_list_widget()
            self._update_stats()
        finally:
            self._in_undo_redo = False

    def _redo(self) -> None:
        if not self._redo_stack:
            return
        self._in_undo_redo = True
        try:
            self._undo_stack.append(self._clone_tasks(self._tasks))
            self._tasks = self._redo_stack.pop()
            self._original_in_progress = next(
                (t.text for t in self._tasks if t.status == TaskStatus.IN_PROGRESS),
                None,
            )
            self._reload_list_widget()
            self._update_stats()
        finally:
            self._in_undo_redo = False

    # ---------- actions ----------

    def _set_status(self, new_status: TaskStatus) -> None:
        """Replace the old Space-cycle behavior with direct status setters.

        Marking DONE additionally pops up the (skippable) reflection prompt.
        """
        index = self._selected_index()
        if not (0 <= index < len(self._tasks)):
            return
        task = self._tasks[index]
        if task.status == new_status:
            return
        self._push_undo()
        if new_status == TaskStatus.IN_PROGRESS:
            for other in self._tasks:
                if other is not task and other.status == TaskStatus.IN_PROGRESS:
                    other.status = TaskStatus.PENDING
        task.status = new_status
        self._reload_list_widget()
        self.list_widget.setCurrentRow(index)
        self._update_stats()
        if new_status == TaskStatus.DONE:
            self._prompt_reflection_for(index)

    def _prompt_reflection_for(self, index: int) -> None:
        if not (0 <= index < len(self._tasks)):
            return
        task = self._tasks[index]
        text, ok = QInputDialog.getMultiLineText(
            self,
            "记录反思（可跳过）",
            f"完成了「{task.text or '（空任务）'}」想留下点什么？\n"
            "卡在哪？哪步可以下次做得更好？空着按确定也行。",
            task.reflection or "",
        )
        if not ok:
            return
        new_reflection = text.strip()
        if new_reflection == (task.reflection or "").strip():
            return
        self._push_undo()
        task.reflection = new_reflection
        self._refresh_item(index)

    def _edit_reflection_for_selected(self) -> None:
        self._prompt_reflection_for(self._selected_index())

    def _add_new_task(self, text: str = "") -> None:
        self._push_undo()
        new_task = PlanTask(text=text, status=TaskStatus.PENDING)
        self._pending_duration_prompt.add(id(new_task))
        index = self._selected_index()
        if 0 <= index < len(self._tasks):
            self._tasks.insert(index + 1, new_task)
            target = index + 1
        else:
            self._tasks.append(new_task)
            target = len(self._tasks) - 1
        self._reload_list_widget()
        self.list_widget.setCurrentRow(target)
        self._update_stats()
        if not text:
            QTimer.singleShot(0, lambda r=target: self._edit_row(r))

    def _delete_selected(self) -> None:
        # No confirmation by design — Ctrl+Z restores the task.
        index = self._selected_index()
        if not (0 <= index < len(self._tasks)):
            return
        self._push_undo()
        self._tasks.pop(index)
        self._reload_list_widget()
        if self._tasks:
            self.list_widget.setCurrentRow(min(index, len(self._tasks) - 1))
        self._update_stats()

    def _move_selected(self, offset: int) -> None:
        index = self._selected_index()
        new_index = index + offset
        if not (0 <= index < len(self._tasks)) or not (0 <= new_index < len(self._tasks)):
            return
        self._push_undo()
        self._tasks[index], self._tasks[new_index] = self._tasks[new_index], self._tasks[index]
        self._reload_list_widget()
        self.list_widget.setCurrentRow(new_index)

    def _show_context_menu(self, pos: QPoint) -> None:
        item = self.list_widget.itemAt(pos)
        menu = QMenu(self.list_widget)
        if item is not None:
            self.list_widget.setCurrentItem(item)
            task = self._tasks[self._selected_index()] \
                if 0 <= self._selected_index() < len(self._tasks) else None
            menu.addAction("编辑  (F2 / Enter)", self._edit_selected)
            status_menu = menu.addMenu("🚥 状态")
            pending_act = status_menu.addAction(
                "○  待办", lambda: self._set_status(TaskStatus.PENDING)
            )
            in_progress_act = status_menu.addAction(
                "▶  进行中", lambda: self._set_status(TaskStatus.IN_PROGRESS)
            )
            done_act = status_menu.addAction(
                "✓  已完成（写反思）", lambda: self._set_status(TaskStatus.DONE)
            )
            if task is not None:
                pending_act.setCheckable(True)
                in_progress_act.setCheckable(True)
                done_act.setCheckable(True)
                pending_act.setChecked(task.status == TaskStatus.PENDING)
                in_progress_act.setChecked(task.status == TaskStatus.IN_PROGRESS)
                done_act.setChecked(task.status == TaskStatus.DONE)
            if task is not None and task.status == TaskStatus.DONE:
                menu.addAction("📝 编辑反思…", self._edit_reflection_for_selected)
            menu.addAction("🍅 专注此项  (Ctrl+P)", self._focus_pomodoro_on_selected)
            menu.addAction("⏱ 设置时长…", self._set_duration_for_selected)
            priority_menu = menu.addMenu("🚦 优先级")
            priority_menu.addAction("🔴 P1  必做", lambda: self._set_priority(TaskPriority.P1))
            priority_menu.addAction("🟡 P2  应做", lambda: self._set_priority(TaskPriority.P2))
            priority_menu.addAction("⚪ P3  可做", lambda: self._set_priority(TaskPriority.P3))
            priority_menu.addSeparator()
            priority_menu.addAction("清除优先级", lambda: self._set_priority(TaskPriority.NONE))
            menu.addAction("上移  (Alt+↑)", lambda: self._move_selected(-1))
            menu.addAction("下移  (Alt+↓)", lambda: self._move_selected(1))
            menu.addSeparator()
        menu.addAction("新增任务  (Ctrl+ +)", self._add_new_task)
        if item is not None:
            menu.addAction("删除任务  (Ctrl+ −)", self._delete_selected)
        menu.addSeparator()
        undo_act = menu.addAction("撤销  (Ctrl+Z)", self._undo)
        undo_act.setEnabled(bool(self._undo_stack))
        redo_act = menu.addAction("重做  (Ctrl+Y)", self._redo)
        redo_act.setEnabled(bool(self._redo_stack))
        menu.exec(self.list_widget.mapToGlobal(pos))

    def _set_priority(self, priority: TaskPriority) -> None:
        index = self._selected_index()
        if not (0 <= index < len(self._tasks)):
            return
        if self._tasks[index].priority == priority:
            return
        self._push_undo()
        self._tasks[index].priority = priority
        self._refresh_item(index)
        self._update_stats()

    def _set_duration_for_selected(self) -> None:
        index = self._selected_index()
        if not (0 <= index < len(self._tasks)):
            return
        task = self._tasks[index]
        value = DurationPickerDialog.get_duration(
            current=task.duration_minutes,
            task_text=task.text or '（空任务）',
            parent=self,
        )
        if value < 0:
            return
        if value == task.duration_minutes:
            return
        self._push_undo()
        task.duration_minutes = value
        self._refresh_item(index)
        self._update_stats()

    def _focus_pomodoro_on_selected(self) -> None:
        index = self._selected_index()
        if not (0 <= index < len(self._tasks)):
            return
        task = self._tasks[index]
        if task.status != TaskStatus.IN_PROGRESS:
            self._push_undo()
        for other in self._tasks:
            if other is not task and other.status == TaskStatus.IN_PROGRESS:
                other.status = TaskStatus.PENDING
        task.status = TaskStatus.IN_PROGRESS
        self._reload_list_widget()
        self.list_widget.setCurrentRow(index)
        self._update_stats()
        self._dispatch_save()
        text = task.text

        def _kick():
            try:
                from shouyu.service.pomodoro import PomodoroService

                PomodoroService.instance().start_work(task_text=text)
            except Exception:
                logging.exception("failed to start pomodoro for selected task")

        QTimer.singleShot(50, _kick)

    # ---------- stats ----------

    def _update_stats(self) -> None:
        total = len(self._tasks)
        done = sum(1 for t in self._tasks if t.status == TaskStatus.DONE)
        in_progress = sum(1 for t in self._tasks if t.status == TaskStatus.IN_PROGRESS)
        total_minutes = sum(t.duration_minutes for t in self._tasks if t.duration_minutes > 0)
        try:
            overload_threshold = Config.overload_threshold_minutes()
        except Exception:
            overload_threshold = 360
        unestimated = sum(
            1
            for t in self._tasks
            if (t.text or '').strip() and t.duration_minutes <= 0
        )
        p1_count = sum(1 for t in self._tasks if t.priority == TaskPriority.P1)
        parts = [f"今日 {done}/{total} 完成"]
        if in_progress:
            parts.append(f"{in_progress} 进行中")
        if p1_count:
            parts.append(f"🔴 {p1_count} 项 P1")
        if total_minutes:
            hours = total_minutes / 60
            if total_minutes > overload_threshold:
                parts.append(f"⚠ 预计 {hours:.1f}h 超载")
            else:
                parts.append(f"≈ {hours:.1f}h 已估时")
        if unestimated > 0:
            parts.append(f"💡 {unestimated} 项未估时")
        self.stats_label.setText("  ·  ".join(parts))
        if total_minutes > overload_threshold:
            self.stats_label.setStyleSheet("color: #FFB454; font-weight: 600;")
        else:
            self.stats_label.setStyleSheet("")
        self.progress_bar.setRange(0, max(total, 1))
        self.progress_bar.setValue(done)

    # ---------- persist ----------

    def _save_and_close(self) -> None:
        self._dispatch_save()
        self.hide()
        self.closed.emit()

    def _dispatch_save(self) -> None:
        # Each save dispatches a fresh worker (writes are idempotent), so
        # later edits never get lost behind an earlier "already saving" flag.
        tasks_snapshot = self._clone_tasks(self._tasks)
        _persist_plan_in_background(tasks_snapshot, self._original_in_progress)

    # ---------- qt overrides ----------

    def closeEvent(self, event) -> None:
        self._dispatch_save()
        event.accept()
