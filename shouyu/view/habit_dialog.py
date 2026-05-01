"""Full-screen startup dialog: habit reminders + today's plan editor.

Now packs (P0 / P1 / P2 / P3 from the previous design review):

    - Greeting + weekday + date + close button (P3 dynamic background tint)
    - Streak counter (🔥 连续 N 天)
    - Today's stats + progress bar
    - Yesterday-at-a-glance (X/Y completed · N pomodoros)
    - Carry-over panel: pick yesterday's unfinished tasks into today
    - Quick-add input (Things 3 / Todoist style)
    - Task list with: MIT highlight, duration badges, drag-and-drop reorder,
      right-click context menu, Excel-like Enter-to-edit-next, autoextend
    - "Focus this task" → demotes others, marks this in_progress, starts a
      pomodoro for it (without waiting for Excel I/O)
    - Reflection text area (saves to a separate `reflections` sheet)
    - Empty-state hint when there are no tasks left

Hotkeys are configured in kb.ini (`show_habits=ctrl+alt+h`) and visible in
the footer hint.
"""
from __future__ import annotations

import logging
import threading
import time
from typing import List, Optional

from PySide6.QtCore import QPoint, Qt, QTimer
from PySide6.QtGui import QColor, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QDialog,
    QFrame,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSplitter,
    QStackedLayout,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)
from PySide6.QtWidgets import QLineEdit

from shouyu.service.plan import (
    DEFAULT_PLAN_TASKS,
    PlanTask,
    TaskStatus,
)
from shouyu.util.state import AppState
from shouyu.view.styles import (
    ACCENT_COLOR_HEX,
    DONE_COLOR_HEX,
    IN_PROGRESS_COLOR_HEX,
    PANEL_COLOR_HEX,
    PENDING_COLOR_HEX,
    SUBTEXT_COLOR_HEX,
    TEXT_COLOR_HEX,
)


_GLYPH = {
    TaskStatus.PENDING: "○",
    TaskStatus.IN_PROGRESS: "▶",
    TaskStatus.DONE: "✓",
}

_COLORS = {
    TaskStatus.PENDING: PENDING_COLOR_HEX,
    TaskStatus.IN_PROGRESS: IN_PROGRESS_COLOR_HEX,
    TaskStatus.DONE: DONE_COLOR_HEX,
}

_NEXT_STATUS = {
    TaskStatus.PENDING: TaskStatus.IN_PROGRESS,
    TaskStatus.IN_PROGRESS: TaskStatus.DONE,
    TaskStatus.DONE: TaskStatus.PENDING,
}

_TASK_ROLE = Qt.UserRole + 1


def _format(task: PlanTask) -> str:
    glyph = _GLYPH.get(task.status, "○")
    parts = [glyph, " ", task.text or ""]
    if task.duration_minutes:
        parts.append(f"   ⏱ {task.duration_minutes}m")
    if task.status == TaskStatus.IN_PROGRESS:
        parts.append("   🎯 重点")
    return "".join(parts)


def _apply_item_style(item: QListWidgetItem, status: TaskStatus) -> None:
    item.setForeground(QColor(_COLORS.get(status, PENDING_COLOR_HEX)))
    font = item.font()
    font.setBold(status == TaskStatus.IN_PROGRESS)
    font.setStrikeOut(status == TaskStatus.DONE)
    base_size = max(font.pointSize(), 10)
    font.setPointSize(base_size + 2 if status == TaskStatus.IN_PROGRESS else base_size)
    item.setFont(font)


def _greeting() -> str:
    h = time.localtime().tm_hour
    if 5 <= h < 11:
        return "🌅 早上好"
    if 11 <= h < 13:
        return "☕ 中午好"
    if 13 <= h < 18:
        return "🌤 下午好"
    if 18 <= h < 23:
        return "🌙 晚上好"
    return "🌌 深夜了"


def _weekday_zh(t: time.struct_time) -> str:
    return ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"][t.tm_wday]


def _today_header_text() -> str:
    t = time.localtime()
    return f"{_greeting()}  ·  {_weekday_zh(t)}  ·  {time.strftime('%Y-%m-%d', t)}"


def _card_color_for_time() -> str:
    h = time.localtime().tm_hour
    if 5 <= h < 11:
        return "#332B26"  # warm morning
    if 11 <= h < 18:
        return PANEL_COLOR_HEX  # neutral day
    return "#262A33"  # cool evening


def _card_qss(bg_hex: str) -> str:
    return f"QFrame#Card {{ background-color: {bg_hex}; border-radius: 10px; }}"


def _persist_plan_in_background(
    tasks_snapshot: List[PlanTask],
    original_in_progress: Optional[str],
    reflection_text: Optional[str] = None,
) -> None:
    """Run the openpyxl backup+save on a worker thread; the UI returns instantly."""

    def _worker():
        try:
            from shouyu.service.excel import KbExcel

            non_empty = [t for t in tasks_snapshot if (t.text or "").strip()]
            excel = KbExcel()
            plan = excel.plan_service()
            plan.write_plan_tasks(non_empty)
            new_in_progress = next(
                (t for t in non_empty if t.status == TaskStatus.IN_PROGRESS), None
            )
            if new_in_progress is not None and new_in_progress.text != original_in_progress:
                plan.switch_in_progress(new_in_progress)
            excel.mark_changed()
            excel.force_save()
            if reflection_text is not None:
                excel.write_reflection(reflection_text)
        except Exception:
            logging.exception("failed to persist plan from habit dialog")

    threading.Thread(target=_worker, name="shouyu-save-plan", daemon=True).start()


def _read_yesterday_snapshot() -> dict:
    """Return {'unfinished': [PlanTask...], 'done': N, 'total': N, 'pomodoros': N}."""
    out = {"unfinished": [], "done": 0, "total": 0, "pomodoros": 0}
    try:
        from shouyu.service.excel import KbExcel

        excel = KbExcel()
        plan = excel.plan_service_for(AppState.yesterday_str())
        if plan is None:
            return out
        tasks = plan.read_plan_tasks()
        out["total"] = len(tasks)
        out["done"] = sum(1 for t in tasks if t.status == TaskStatus.DONE)
        out["unfinished"] = [
            PlanTask(text=t.text, status=TaskStatus.PENDING, duration_minutes=t.duration_minutes)
            for t in tasks
            if t.status != TaskStatus.DONE
        ]
        out["pomodoros"] = plan.count_pomodoros_logged()
    except Exception:
        logging.exception("failed to read yesterday snapshot")
    return out


class HabitDialog(QDialog):
    _instance: Optional["HabitDialog"] = None

    @classmethod
    def get_or_create(cls) -> "HabitDialog":
        if cls._instance is None:
            cls._instance = HabitDialog()
        return cls._instance

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("授渔 · 今日仪式")
        self.setModal(False)
        self.setSizeGripEnabled(False)

        self._habits: List[str] = []
        self._tasks: List[PlanTask] = []
        self._original_in_progress: Optional[str] = None
        self._closing = False
        self._save_already_dispatched = False
        self._yesterday_unfinished: List[PlanTask] = []
        self._yesterday_checkboxes: List[QCheckBox] = []
        self._reflection_dirty = False

        self._build_ui()
        self._install_shortcuts()

    # ---------- public API ----------

    def set_habits(self, habits: List[str]) -> None:
        self._habits = list(habits)
        self._render_habits()

    def refresh_plan_from_excel(self) -> None:
        from shouyu.service.excel import KbExcel

        try:
            excel = KbExcel()
            tasks = excel.plan_service().read_plan_tasks()
            reflection = excel.read_reflection()
        except Exception:
            logging.exception("failed to read plan from Excel")
            tasks = []
            reflection = ''

        if not tasks:
            tasks = [PlanTask(text=t, status=TaskStatus.PENDING) for t in DEFAULT_PLAN_TASKS]

        self._tasks = tasks
        in_progress = next((t for t in tasks if t.status == TaskStatus.IN_PROGRESS), None)
        self._original_in_progress = in_progress.text if in_progress else None

        self.reflection_edit.blockSignals(True)
        self.reflection_edit.setPlainText(reflection)
        self.reflection_edit.blockSignals(False)
        self._reflection_dirty = False

        self._render_tasks()
        self._update_stats()

        # Yesterday & streak: cheap enough to fetch on every open.
        snap = _read_yesterday_snapshot()
        self._yesterday_unfinished = list(snap["unfinished"])
        self._render_yesterday(snap)
        self._render_streak()

    def show_fullscreen(self) -> None:
        self._closing = False
        self._save_already_dispatched = False
        self._reset_action_buttons()
        self._update_header()
        self._apply_time_theme()
        self.showFullScreen()
        self.raise_()
        self.activateWindow()
        self.list_widget.setFocus()

    # ---------- ui ----------

    def _build_ui(self) -> None:
        self.setObjectName("HabitDialogRoot")

        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 22, 28, 22)
        outer.setSpacing(14)

        outer.addLayout(self._build_header_row())
        outer.addLayout(self._build_secondary_header_row())

        body = QSplitter(Qt.Horizontal)
        body.setHandleWidth(2)
        body.setChildrenCollapsible(False)
        body.addWidget(self._build_habit_card())
        body.addWidget(self._build_task_card())
        body.setStretchFactor(0, 1)
        body.setStretchFactor(1, 2)
        outer.addWidget(body, stretch=1)

        outer.addLayout(self._build_footer_row())

    def _build_header_row(self) -> QHBoxLayout:
        row = QHBoxLayout()
        row.setSpacing(12)

        self.header_label = QLabel(_today_header_text())
        self.header_label.setObjectName("TitleLabel")
        row.addWidget(self.header_label)

        row.addStretch(1)

        self.streak_label = QLabel("")
        self.streak_label.setStyleSheet(
            f"color: #FF8A4C; font-size: 14px; font-weight: 600;"
        )
        row.addWidget(self.streak_label)

        close_btn = QPushButton("×")
        close_btn.setToolTip("关闭 (Esc)")
        close_btn.setFixedSize(36, 32)
        close_btn.setAutoDefault(False)
        close_btn.setDefault(False)
        close_btn.setFocusPolicy(Qt.NoFocus)
        close_btn.setStyleSheet(
            "QPushButton {"
            f"  font-size: 22px; font-weight: 700;"
            f"  background: transparent; border: none; color: {SUBTEXT_COLOR_HEX};"
            "}"
            f"QPushButton:hover {{ color: {TEXT_COLOR_HEX}; }}"
        )
        close_btn.clicked.connect(self.reject)
        row.addWidget(close_btn)
        return row

    def _build_secondary_header_row(self) -> QHBoxLayout:
        row = QHBoxLayout()
        row.setSpacing(16)

        self.stats_label = QLabel("")
        self.stats_label.setObjectName("SubtitleLabel")
        row.addWidget(self.stats_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(6)
        self.progress_bar.setMinimumWidth(180)
        self.progress_bar.setStyleSheet(
            "QProgressBar { background-color: #2B2B2B; border: none; border-radius: 3px; } "
            f"QProgressBar::chunk {{ background-color: {ACCENT_COLOR_HEX}; border-radius: 3px; }}"
        )
        row.addWidget(self.progress_bar, stretch=1)

        self.yesterday_glance_label = QLabel("")
        self.yesterday_glance_label.setObjectName("SubtitleLabel")
        row.addWidget(self.yesterday_glance_label)

        return row

    def _build_habit_card(self) -> QWidget:
        self.habit_card = QFrame()
        self.habit_card.setObjectName("Card")
        self.habit_card.setStyleSheet(_card_qss(PANEL_COLOR_HEX))

        layout = QVBoxLayout(self.habit_card)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(10)

        title = QLabel("今日习惯")
        title.setObjectName("TitleLabel")
        layout.addWidget(title)

        subtitle = QLabel("回顾这些原则可以减少每天反复内耗")
        subtitle.setObjectName("SubtitleLabel")
        layout.addWidget(subtitle)

        self.habit_container = QWidget()
        self.habit_layout = QVBoxLayout(self.habit_container)
        self.habit_layout.setContentsMargins(0, 0, 0, 0)
        self.habit_layout.setSpacing(8)

        scroll = QScrollArea()
        scroll.setWidget(self.habit_container)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")
        layout.addWidget(scroll, stretch=1)

        return self.habit_card

    def _build_task_card(self) -> QWidget:
        self.task_card = QFrame()
        self.task_card.setObjectName("Card")
        self.task_card.setStyleSheet(_card_qss(PANEL_COLOR_HEX))

        layout = QVBoxLayout(self.task_card)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(10)

        title_row = QHBoxLayout()
        title = QLabel("今日要事")
        title.setObjectName("TitleLabel")
        title_row.addWidget(title)
        title_row.addStretch(1)

        add_btn = QPushButton("+ 新增任务")
        add_btn.setToolTip("Ctrl++ 添加一个任务到当前下方")
        add_btn.setAutoDefault(False)
        add_btn.setDefault(False)
        add_btn.clicked.connect(self._add_new_task)
        title_row.addWidget(add_btn)
        layout.addLayout(title_row)

        hint = QLabel(
            "↑↓ 选择 ·  F2 / 回车 编辑 ·  Space 切换状态 ·  Alt+↑↓ 重排 ·  "
            "拖拽排序 ·  Ctrl+ + 添加 ·  Ctrl+ − 删除 ·  右键查看更多"
        )
        hint.setObjectName("HintLabel")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        # Quick add (Things-3 style): always visible, Enter posts a new task.
        self.quick_add = QLineEdit()
        self.quick_add.setPlaceholderText("➕ 快速添加任务并按回车…")
        self.quick_add.returnPressed.connect(self._on_quick_add_submitted)
        layout.addWidget(self.quick_add)

        # Yesterday carry-over (only shown when there's data to carry).
        self.carryover_card = QFrame()
        self.carryover_card.setObjectName("CarryCard")
        self.carryover_card.setStyleSheet(
            "QFrame#CarryCard { background-color: rgba(15, 98, 254, 0.10); "
            "border: 1px solid rgba(15, 98, 254, 0.35); border-radius: 8px; }"
        )
        self.carryover_layout = QVBoxLayout(self.carryover_card)
        self.carryover_layout.setContentsMargins(12, 10, 12, 10)
        self.carryover_layout.setSpacing(6)
        self.carryover_card.setVisible(False)
        layout.addWidget(self.carryover_card)

        # Stacked: list_widget OR empty-state placeholder.
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
        self.task_stack.addWidget(self.list_widget)

        self.empty_label = QLabel("🎉  今天还没有任务\n按 Ctrl+ + 添加一项重要的事")
        self.empty_label.setAlignment(Qt.AlignCenter)
        self.empty_label.setStyleSheet(
            f"color: {SUBTEXT_COLOR_HEX}; font-size: 14px; line-height: 22px;"
        )
        self.task_stack.addWidget(self.empty_label)

        layout.addWidget(self.task_stack_host, stretch=1)

        # Reflection
        reflection_title = QLabel("📝 今日反思 (晚上回来再写也行)")
        reflection_title.setObjectName("SubtitleLabel")
        layout.addWidget(reflection_title)

        self.reflection_edit = QTextEdit()
        self.reflection_edit.setPlaceholderText(
            "今天最有成就感的一件事？哪个任务卡住了，下次怎么避免？"
        )
        self.reflection_edit.setFixedHeight(96)
        self.reflection_edit.textChanged.connect(self._on_reflection_changed)
        layout.addWidget(self.reflection_edit)

        return self.task_card

    def _build_footer_row(self) -> QHBoxLayout:
        row = QHBoxLayout()
        row.setSpacing(10)

        tip = QLabel(
            "Esc 跳过 ·  Ctrl+Enter 开始今天 ·  Ctrl+Alt+H 重新打开仪式"
        )
        tip.setObjectName("HintLabel")
        row.addWidget(tip)

        row.addStretch(1)

        self.skip_btn = QPushButton("跳过 (Esc)")
        self.skip_btn.setAutoDefault(False)
        self.skip_btn.setDefault(False)
        self.skip_btn.clicked.connect(self.reject)
        row.addWidget(self.skip_btn)

        self.start_btn = QPushButton("开始今天 (Ctrl+Enter)")
        self.start_btn.setObjectName("PrimaryButton")
        self.start_btn.setAutoDefault(False)
        self.start_btn.setDefault(False)
        self.start_btn.clicked.connect(self._save_and_accept)
        row.addWidget(self.start_btn)

        return row

    def _render_habits(self) -> None:
        while self.habit_layout.count():
            item = self.habit_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        for i, text in enumerate(self._habits, start=1):
            label = QLabel(f"{i}.  {text}")
            label.setObjectName("HabitLabel")
            label.setWordWrap(True)
            self.habit_layout.addWidget(label)
        self.habit_layout.addStretch(1)

    def _render_tasks(self) -> None:
        self.list_widget.blockSignals(True)
        self.list_widget.clear()
        for task in self._tasks:
            item = QListWidgetItem(_format(task))
            flags = item.flags()
            flags |= Qt.ItemIsEditable | Qt.ItemIsDragEnabled
            flags &= ~Qt.ItemIsDropEnabled  # don't drop INTO an item, only between
            item.setFlags(flags)
            item.setData(_TASK_ROLE, task)
            _apply_item_style(item, task.status)
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

    def _render_yesterday(self, snap: dict) -> None:
        # Glance label
        if snap["total"] > 0 or snap["pomodoros"] > 0:
            parts = []
            if snap["total"] > 0:
                parts.append(f"昨日 {snap['done']}/{snap['total']} 完成")
            if snap["pomodoros"] > 0:
                parts.append(f"🍅 {snap['pomodoros']}")
            self.yesterday_glance_label.setText("  ·  ".join(parts))
        else:
            self.yesterday_glance_label.setText("")

        # Carry-over card body
        for cb in self._yesterday_checkboxes:
            cb.deleteLater()
        self._yesterday_checkboxes = []
        # Wipe layout
        while self.carryover_layout.count():
            item = self.carryover_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

        if not self._yesterday_unfinished:
            self.carryover_card.setVisible(False)
            return

        header = QHBoxLayout()
        header.setSpacing(6)
        header_label = QLabel(f"📅 昨日还有 {len(self._yesterday_unfinished)} 项未完成")
        header_label.setStyleSheet(f"font-weight: 600; color: {TEXT_COLOR_HEX};")
        header.addWidget(header_label)
        header.addStretch(1)

        # Compact button stylesheet — smaller padding so the labels don't get clipped
        # by the global QPushButton rule (which uses 8px vertical padding).
        compact_qss = (
            "QPushButton {"
            f"  background-color: rgba(255,255,255,0.06);"
            f"  color: {TEXT_COLOR_HEX};"
            "  border: 1px solid rgba(255,255,255,0.12);"
            "  border-radius: 4px;"
            "  padding: 3px 10px;"
            "  font-size: 12px;"
            "  min-height: 0;"
            "}"
            "QPushButton:hover { background-color: rgba(255,255,255,0.14); }"
        )
        primary_compact_qss = (
            "QPushButton {"
            f"  background-color: {ACCENT_COLOR_HEX};"
            "  color: white;"
            "  border: none;"
            "  border-radius: 4px;"
            "  padding: 3px 12px;"
            "  font-size: 12px;"
            "  font-weight: 600;"
            "  min-height: 0;"
            "}"
            "QPushButton:hover { background-color: #2D7BF1; }"
        )

        select_all_btn = QPushButton("全选")
        select_all_btn.setAutoDefault(False)
        select_all_btn.setStyleSheet(compact_qss)
        select_all_btn.clicked.connect(lambda: self._set_carryover_all(True))
        header.addWidget(select_all_btn)

        none_btn = QPushButton("全不选")
        none_btn.setAutoDefault(False)
        none_btn.setStyleSheet(compact_qss)
        none_btn.clicked.connect(lambda: self._set_carryover_all(False))
        header.addWidget(none_btn)

        carry_btn = QPushButton("结转选中 →")
        carry_btn.setAutoDefault(False)
        carry_btn.setStyleSheet(primary_compact_qss)
        carry_btn.clicked.connect(self._carry_over_now)
        header.addWidget(carry_btn)

        self.carryover_layout.addLayout(header)

        for task in self._yesterday_unfinished:
            cb = QCheckBox(task.text or "（空任务）")
            cb.setChecked(True)
            cb.setStyleSheet(f"QCheckBox {{ color: {TEXT_COLOR_HEX}; }}")
            self._yesterday_checkboxes.append(cb)
            self.carryover_layout.addWidget(cb)

        self.carryover_card.setVisible(True)

    def _render_streak(self) -> None:
        days = AppState.streak_days()
        if days <= 0:
            self.streak_label.setText("")
            self.streak_label.setVisible(False)
        else:
            self.streak_label.setText(f"🔥 连续 {days} 天")
            self.streak_label.setVisible(True)

    def _install_shortcuts(self) -> None:
        bindings = [
            ("Space", self._cycle_status),
            ("F2", self._edit_selected),
            ("Return", self._edit_selected),
            ("Enter", self._edit_selected),
            ("Alt+Up", lambda: self._move(-1)),
            ("Alt+Down", lambda: self._move(1)),
            ("Ctrl++", self._add_new_task),
            ("Ctrl+=", self._add_new_task),
            ("Ctrl+Shift+=", self._add_new_task),
            ("Ctrl+-", self._delete_selected),
            ("Ctrl+Minus", self._delete_selected),
            ("Ctrl+Return", self._save_and_accept),
            ("Ctrl+Enter", self._save_and_accept),
            ("Ctrl+L", self._quick_add_focus),
            ("Ctrl+P", self._focus_pomodoro_on_selected),
            ("Escape", self.reject),
        ]
        for sequence, callback in bindings:
            shortcut = QShortcut(QKeySequence(sequence), self)
            shortcut.setContext(Qt.WindowShortcut)
            shortcut.activated.connect(callback)

    # ---------- list interactions ----------

    @staticmethod
    def _strip_glyph(text: str) -> str:
        text = text or ""
        for glyph in _GLYPH.values():
            if text.startswith(glyph):
                text = text[len(glyph):]
                break
        text = text.lstrip()
        # Strip trailing badges that we add ("⏱", "🎯") so user editing stays sane.
        for marker in ("   🎯", "   ⏱"):
            idx = text.find(marker)
            if idx >= 0:
                text = text[:idx]
        return text.strip()

    def _on_item_text_edited(self, item: QListWidgetItem) -> None:
        index = self.list_widget.row(item)
        if 0 <= index < len(self._tasks):
            self._tasks[index].text = self._strip_glyph(item.text())
            self._refresh_item(index)
            self._update_stats()

        next_row = index + 1
        if next_row < self.list_widget.count():
            self.list_widget.setCurrentRow(next_row)
            QTimer.singleShot(0, lambda r=next_row: self._edit_row(r))
        elif (
            0 <= index < len(self._tasks)
            and (self._tasks[index].text or "").strip()
        ):
            self._tasks.append(PlanTask(text="", status=TaskStatus.PENDING))
            new_row = len(self._tasks) - 1
            self._render_tasks()
            self.list_widget.setCurrentRow(new_row)
            self._update_stats()
            QTimer.singleShot(0, lambda r=new_row: self._edit_row(r))

    def _edit_selected(self) -> None:
        self._edit_row(self.list_widget.currentRow())

    def _edit_row(self, index: int) -> None:
        if 0 <= index < self.list_widget.count():
            item = self.list_widget.item(index)
            if item is not None:
                self.list_widget.editItem(item)

    def _refresh_item(self, index: int) -> None:
        if not (0 <= index < len(self._tasks)):
            return
        task = self._tasks[index]
        item = self.list_widget.item(index)
        self.list_widget.blockSignals(True)
        item.setText(_format(task))
        _apply_item_style(item, task.status)
        item.setData(_TASK_ROLE, task)
        self.list_widget.blockSignals(False)

    def _cycle_status(self) -> None:
        index = self.list_widget.currentRow()
        if not (0 <= index < len(self._tasks)):
            return
        task = self._tasks[index]
        next_status = _NEXT_STATUS[task.status]
        if next_status == TaskStatus.IN_PROGRESS:
            for other in self._tasks:
                if other is not task and other.status == TaskStatus.IN_PROGRESS:
                    other.status = TaskStatus.PENDING
        task.status = next_status
        self._render_tasks()
        self.list_widget.setCurrentRow(index)
        self._update_stats()

    def _move(self, offset: int) -> None:
        index = self.list_widget.currentRow()
        new_index = index + offset
        if not (0 <= index < len(self._tasks)) or not (0 <= new_index < len(self._tasks)):
            return
        self._tasks[index], self._tasks[new_index] = self._tasks[new_index], self._tasks[index]
        self._render_tasks()
        self.list_widget.setCurrentRow(new_index)

    def _on_rows_moved(self, *args) -> None:
        # Rebuild self._tasks to match the new visual order using PlanTask refs
        # we stashed in _TASK_ROLE.
        new_order: List[PlanTask] = []
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            task = item.data(_TASK_ROLE)
            if isinstance(task, PlanTask):
                new_order.append(task)
        if len(new_order) == len(self._tasks):
            self._tasks = new_order

    def _add_new_task(self, text: str = "") -> None:
        new_task = PlanTask(text=text, status=TaskStatus.PENDING)
        index = self.list_widget.currentRow()
        if 0 <= index < len(self._tasks):
            self._tasks.insert(index + 1, new_task)
            target = index + 1
        else:
            self._tasks.append(new_task)
            target = len(self._tasks) - 1
        self._render_tasks()
        self.list_widget.setCurrentRow(target)
        self._update_stats()
        if not text:
            QTimer.singleShot(0, lambda r=target: self._edit_row(r))

    def _delete_selected(self) -> None:
        index = self.list_widget.currentRow()
        if not (0 <= index < len(self._tasks)):
            return
        text = self._tasks[index].text or "（空任务）"
        confirm = QMessageBox.question(
            self,
            "删除任务",
            f"确定要删除「{text}」吗？此操作不可撤销。",
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        self._tasks.pop(index)
        self._render_tasks()
        if self._tasks:
            self.list_widget.setCurrentRow(min(index, len(self._tasks) - 1))
        self._update_stats()

    def _show_context_menu(self, pos: QPoint) -> None:
        item = self.list_widget.itemAt(pos)
        menu = QMenu(self.list_widget)
        if item is not None:
            self.list_widget.setCurrentItem(item)
            menu.addAction("编辑  (F2 / Enter)", self._edit_selected)
            menu.addAction("切换状态  (Space)", self._cycle_status)
            menu.addAction("🍅 专注此项  (Ctrl+P)", self._focus_pomodoro_on_selected)
            menu.addAction("⏱ 设置时长…", self._set_duration_for_selected)
            menu.addAction("上移  (Alt+↑)", lambda: self._move(-1))
            menu.addAction("下移  (Alt+↓)", lambda: self._move(1))
            menu.addSeparator()
        menu.addAction("新增任务  (Ctrl+ +)", self._add_new_task)
        if item is not None:
            menu.addAction("删除任务  (Ctrl+ −)", self._delete_selected)
        menu.exec(self.list_widget.mapToGlobal(pos))

    def _set_duration_for_selected(self) -> None:
        index = self.list_widget.currentRow()
        if not (0 <= index < len(self._tasks)):
            return
        task = self._tasks[index]
        value, ok = QInputDialog.getInt(
            self,
            "设置预计时长",
            f"为「{task.text or '（空任务）'}」设置时长（分钟，0 表示清除）：",
            value=task.duration_minutes,
            minValue=0,
            maxValue=480,
            step=5,
        )
        if not ok:
            return
        task.duration_minutes = value
        self._refresh_item(index)
        self._update_stats()

    # ---------- quick add ----------

    def _quick_add_focus(self) -> None:
        self.quick_add.setFocus()
        self.quick_add.selectAll()

    def _on_quick_add_submitted(self) -> None:
        text = (self.quick_add.text() or "").strip()
        if not text:
            return
        self.quick_add.clear()
        self._tasks.append(PlanTask(text=text, status=TaskStatus.PENDING))
        self._render_tasks()
        self.list_widget.setCurrentRow(len(self._tasks) - 1)
        self._update_stats()

    # ---------- carry-over ----------

    def _set_carryover_all(self, checked: bool) -> None:
        for cb in self._yesterday_checkboxes:
            cb.setChecked(checked)

    def _carry_over_now(self) -> None:
        if not self._yesterday_unfinished:
            return
        chosen = [
            task
            for task, cb in zip(self._yesterday_unfinished, self._yesterday_checkboxes)
            if cb.isChecked()
        ]
        if not chosen:
            self._yesterday_unfinished = []
            self._render_yesterday({"unfinished": [], "done": 0, "total": 0, "pomodoros": 0})
            return
        existing_texts = {(t.text or "").strip() for t in self._tasks if (t.text or "").strip()}
        added = 0
        for task in chosen:
            text = (task.text or "").strip()
            if not text or text in existing_texts:
                continue
            self._tasks.append(
                PlanTask(
                    text=task.text,
                    status=TaskStatus.PENDING,
                    duration_minutes=task.duration_minutes,
                )
            )
            existing_texts.add(text)
            added += 1
        self._yesterday_unfinished = []
        self._render_yesterday({"unfinished": [], "done": 0, "total": 0, "pomodoros": 0})
        self._render_tasks()
        if added > 0 and self._tasks:
            self.list_widget.setCurrentRow(len(self._tasks) - 1)
        self._update_stats()

    # ---------- focus pomodoro ----------

    def _focus_pomodoro_on_selected(self) -> None:
        index = self.list_widget.currentRow()
        if not (0 <= index < len(self._tasks)):
            return
        task = self._tasks[index]
        # promote to in_progress
        for other in self._tasks:
            if other is not task and other.status == TaskStatus.IN_PROGRESS:
                other.status = TaskStatus.PENDING
        task.status = TaskStatus.IN_PROGRESS
        self._render_tasks()
        self.list_widget.setCurrentRow(index)
        self._update_stats()
        # Save first (async) so the daily sheet has the task in active area; then start pomodoro.
        self._dispatch_save()
        text = task.text

        def _kick():
            try:
                from shouyu.service.pomodoro import PomodoroService

                PomodoroService.instance().start_work(task_text=text)
            except Exception:
                logging.exception("failed to start pomodoro for selected task")

        QTimer.singleShot(50, _kick)

    # ---------- reflection ----------

    def _on_reflection_changed(self) -> None:
        self._reflection_dirty = True

    # ---------- stats / header ----------

    def _update_header(self) -> None:
        self.header_label.setText(_today_header_text())

    def _apply_time_theme(self) -> None:
        bg = _card_color_for_time()
        qss = _card_qss(bg)
        self.habit_card.setStyleSheet(qss)
        self.task_card.setStyleSheet(qss)

    def _update_stats(self) -> None:
        total = len(self._tasks)
        done = sum(1 for t in self._tasks if t.status == TaskStatus.DONE)
        in_progress = sum(1 for t in self._tasks if t.status == TaskStatus.IN_PROGRESS)
        total_minutes = sum(t.duration_minutes for t in self._tasks if t.duration_minutes > 0)
        parts = [f"今日 {done}/{total} 完成"]
        if in_progress:
            parts.append(f"{in_progress} 进行中")
        if total_minutes:
            parts.append(f"≈ {total_minutes}m 估时")
        self.stats_label.setText("  ·  ".join(parts))

        self.progress_bar.setRange(0, max(total, 1))
        self.progress_bar.setValue(done)

    def _reset_action_buttons(self) -> None:
        self.start_btn.setEnabled(True)
        self.start_btn.setText("开始今天 (Ctrl+Enter)")
        self.skip_btn.setEnabled(True)

    # ---------- persist ----------

    def _save_and_accept(self) -> None:
        if self._closing:
            return
        self._closing = True

        self.start_btn.setText("保存中…")
        self.start_btn.setEnabled(False)
        self.skip_btn.setEnabled(False)

        # Mark today's ritual as completed and update the streak.
        try:
            new_streak = AppState.update_ritual_streak()
            logging.info(f"ritual streak now: {new_streak}")
        except Exception:
            logging.exception("failed to update ritual streak")

        self._dispatch_save()
        QTimer.singleShot(0, self.accept)

    def _dispatch_save(self) -> None:
        if self._save_already_dispatched:
            return
        self._save_already_dispatched = True
        tasks_snapshot = [
            PlanTask(
                text=t.text,
                status=t.status,
                row=t.row,
                duration_minutes=t.duration_minutes,
            )
            for t in self._tasks
        ]
        reflection = (
            self.reflection_edit.toPlainText() if self._reflection_dirty else None
        )
        _persist_plan_in_background(tasks_snapshot, self._original_in_progress, reflection)

    # ---------- qt overrides ----------

    def closeEvent(self, event) -> None:
        self._closing = True
        self._dispatch_save()
        event.accept()
