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

from PySide6.QtCore import QMimeData, QPoint, Qt, QTimer
from PySide6.QtGui import QColor, QDrag, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QAbstractItemDelegate,
    QAbstractItemView,
    QApplication,
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
    QVBoxLayout,
    QWidget,
)
from shouyu.view.duration_dialog import DurationPickerDialog

from shouyu.config import Config
from shouyu.service.plan import (
    DEFAULT_PLAN_TASKS,
    PlanTask,
    TaskCategory,
    TaskPriority,
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

_PRIORITY_BADGE = {
    TaskPriority.P1: "🔴 P1",
    TaskPriority.P2: "🟡 P2",
    TaskPriority.P3: "⚪ P3",
}

_CATEGORY_BADGE = {
    TaskCategory.WORK: "🏢",
    TaskCategory.LIFE: "🏠",
}

_TASK_ROLE = Qt.UserRole + 1


def _days_since(created_date: str) -> int:
    """Whole days elapsed since `created_date` (YYYY-MM-DD). 0 on parse error."""
    if not created_date:
        return 0
    try:
        created = time.strptime(created_date.strip(), "%Y-%m-%d")
        created_days = time.mktime(created) // 86400
        today_days = time.mktime(time.localtime()) // 86400
        return max(0, int(today_days - created_days))
    except Exception:
        return 0


def _format(task: PlanTask) -> str:
    glyph = _GLYPH.get(task.status, "○")
    parts = [glyph, " ", task.text or ""]
    cat_badge = _CATEGORY_BADGE.get(task.category)
    if cat_badge:
        parts.append(f"   {cat_badge}")
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


def _format_backlog(task: PlanTask) -> str:
    """Backlog rows always render as pending (the pool has no status), and add
    a "已搁置 N 天" badge instead of the 🎯/category decorations. Category is
    implicit in which section the row lives in, so no 🏢/🏠 badge here."""
    parts = ["○ ", task.text or ""]
    badge = _PRIORITY_BADGE.get(task.priority)
    if badge:
        parts.append(f"   {badge}")
    if task.duration_minutes:
        parts.append(f"   ⏱ {task.duration_minutes}m")
    days = _days_since(task.created_date)
    if days >= 2:
        parts.append(f"   ⏳ {days}d")
    if (task.reflection or "").strip():
        parts.append("   📝")
    return "".join(parts)


class _DragDropList(QListWidget):
    """QListWidget that supports dragging items *between* three sibling lists
    (today / work-backlog / life-backlog) as well as internal reordering.

    QListWidget's default InternalMove can't move the PlanTask payload across
    widgets (item roles aren't serialized into the drag MIME). So instead we
    run our own QDrag and let the dialog reconcile the in-memory task lists on
    drop — the whole thing is driven by `dialog._begin_drag` / `_handle_drop`
    which mutate the backing lists and re-render. We deliberately do NOT call
    super().dropEvent / rely on the model, so Qt never removes/inserts rows
    behind our back.
    """

    def __init__(self, dialog: "HabitDialog") -> None:
        super().__init__()
        self._dialog = dialog

    def startDrag(self, supported_actions) -> None:  # noqa: N802 (Qt override)
        item = self.currentItem()
        if item is None:
            return
        self._dialog._begin_drag(self, item)
        drag = QDrag(self)
        mime = QMimeData()
        mime.setText(item.text())
        drag.setMimeData(mime)
        drag.exec(Qt.MoveAction)

    def dragEnterEvent(self, event) -> None:  # noqa: N802
        event.acceptProposedAction()

    def dragMoveEvent(self, event) -> None:  # noqa: N802
        event.acceptProposedAction()

    def dropEvent(self, event) -> None:  # noqa: N802
        target_row = self._drop_row(event)
        event.acceptProposedAction()
        # Defer so the source list's startDrag/exec loop fully unwinds before
        # we clear & rebuild the widgets (mutating during the drag can crash).
        dialog = self._dialog
        QTimer.singleShot(0, lambda: dialog._handle_drop(self, target_row))

    def _drop_row(self, event) -> int:
        try:
            pos = event.position().toPoint()
        except Exception:
            pos = event.pos()
        item = self.itemAt(pos)
        if item is None:
            return self.count()
        row = self.row(item)
        rect = self.visualItemRect(item)
        if pos.y() > rect.center().y():
            row += 1
        return row


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


_RETRY_DELAYS_S = [1, 2, 5, 10, 15]  # ~33s of background retry budget


def _persist_plan_in_background(
    tasks_snapshot: List[PlanTask],
    original_in_progress: Optional[str],
    yesterday_done_rows: Optional[List[int]] = None,
    work_snapshot: Optional[List[PlanTask]] = None,
    life_snapshot: Optional[List[PlanTask]] = None,
) -> None:
    """Run the openpyxl backup+save on a worker thread; the UI returns instantly.

    The most common failure is `PermissionError` because the canonical Excel
    file is currently open in MS Excel / WPS. We retry transparently with
    exponential-ish backoff (so closing Excel within ~30 seconds rescues the
    save automatically), and only bother the user with a popup if all retries
    are exhausted. In that case we also write the in-memory workbook to a
    sibling `<name>.unsaved_<ts>.xlsx` so the data is never silently lost.

    `yesterday_done_rows`: plan-area row numbers in *yesterday's* worksheet
    that the user just marked as DONE via the carry-over card. Useful for the
    "I forgot to check off this task yesterday" case.

    NOTE: reflections live alongside each task (PlanTask.reflection -> plan
    sheet column E) now, so there is no separate `reflection_text` channel.
    """

    def _stage_changes(excel) -> None:
        non_empty = [t for t in tasks_snapshot if (t.text or "").strip()]
        plan = excel.plan_service()
        plan.write_plan_tasks(non_empty)
        new_in_progress = next(
            (t for t in non_empty if t.status == TaskStatus.IN_PROGRESS), None
        )
        if new_in_progress is not None and new_in_progress.text != original_in_progress:
            plan.switch_in_progress(new_in_progress)
        if yesterday_done_rows:
            yesterday_plan = excel.plan_service_for(AppState.yesterday_str())
            if yesterday_plan is not None:
                for row in yesterday_done_rows:
                    if row:
                        yesterday_plan.mark_plan_done(row)
        # Backlog pools: full rewrite of each sheet from the in-memory copy.
        if work_snapshot is not None:
            excel.backlog_service(TaskCategory.WORK).write(
                [t for t in work_snapshot if (t.text or "").strip()]
            )
        if life_snapshot is not None:
            excel.backlog_service(TaskCategory.LIFE).write(
                [t for t in life_snapshot if (t.text or "").strip()]
            )
        excel.mark_changed()

    def _notify_retry_started(reason: str) -> None:
        try:
            from shouyu.view.msgbox import MessageBox, MessageType

            MessageBox.pop_up_message(
                title="保存失败 · 后台重试中",
                msg=f"{reason}（最多 30 秒，请关闭 Excel 后等待）",
                level=MessageType.ERROR,
            )
        except Exception:
            logging.exception("failed to show retry toast")

    def _notify_success_after_retry(attempt: int) -> None:
        try:
            from shouyu.view.msgbox import MessageBox, MessageType

            MessageBox.pop_up_message(
                title="已保存",
                msg=f"第 {attempt} 次重试成功",
                level=MessageType.SUCCESS,
            )
        except Exception:
            logging.exception("failed to show success toast")

    def _notify_final_failure(err: Exception, preserved: str) -> None:
        try:
            from shouyu.view.qt_app import QtApp

            is_lock = isinstance(err, PermissionError) or 'Permission' in str(err)
            if is_lock:
                msg = (
                    "保存失败：无法写入 Excel 文件，文件可能正被其他程序（如 MS Excel / WPS）占用。\n\n"
                    f"已重试 {len(_RETRY_DELAYS_S) + 1} 次仍然失败。"
                )
            else:
                msg = f"保存失败：{err}"
            if preserved:
                msg += (
                    f"\n\n你的改动已经写入备用文件，不会丢失：\n{preserved}\n\n"
                    "请关闭 Excel，然后下次打开今日任务时按提示恢复，或手动用此备用文件覆盖主文件。"
                )
            else:
                msg += "\n\n（备用文件也写入失败；请手动核对最近的备份。）"
            QtApp.show_save_status('error', "今日任务保存失败", msg)
        except Exception:
            logging.exception("failed to dispatch final-failure popup")

    def _worker():
        from shouyu.service.excel import KbExcel

        excel: Optional[KbExcel] = None
        last_err: Optional[Exception] = None
        notified_retry = False

        for attempt in range(len(_RETRY_DELAYS_S) + 1):
            try:
                if excel is None:
                    excel = KbExcel()
                    _stage_changes(excel)
                excel.force_save()
                if attempt > 0:
                    _notify_success_after_retry(attempt + 1)
                return
            except PermissionError as e:
                last_err = e
                logging.warning(
                    f"save attempt {attempt + 1} failed (locked): {e}"
                )
            except Exception as e:
                last_err = e
                logging.exception(f"save attempt {attempt + 1} failed")
                # Non-recoverable error class — don't bother retrying.
                break

            if attempt < len(_RETRY_DELAYS_S):
                if not notified_retry:
                    notified_retry = True
                    _notify_retry_started("Excel 文件被占用")
                time.sleep(_RETRY_DELAYS_S[attempt])

        # All retries exhausted — preserve in-memory state and notify.
        preserved = ''
        if excel is not None:
            try:
                preserved = excel.preserve_unsaved() or ''
            except Exception:
                logging.exception("failed to preserve unsaved changes")
        _notify_final_failure(last_err or RuntimeError("unknown"), preserved)

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
        # NOTE: we deliberately preserve `row` here so the carry-over UI can
        # later mark a yesterday task as DONE in the original worksheet via
        # PlanService.mark_plan_done(row).
        out["unfinished"] = [
            PlanTask(
                text=t.text,
                status=TaskStatus.PENDING,
                row=t.row,
                duration_minutes=t.duration_minutes,
                priority=t.priority,
                reflection=t.reflection,
                category=t.category,
            )
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
        self.setWindowTitle("授渔 · 今日任务")
        self.setModal(False)
        self.setSizeGripEnabled(False)
        # Frameless + always-on-top: the morning ritual MUST be visible above
        # whatever the user was doing (browser, IDE, etc.). Otherwise the dialog
        # gets buried and the ritual silently dies.
        self.setWindowFlag(Qt.FramelessWindowHint, True)
        self.setWindowFlag(Qt.WindowStaysOnTopHint, True)

        self._habits: List[str] = []
        self._tasks: List[PlanTask] = []
        # Cross-day Backlog pools (work / life), each backed by its own sheet.
        self._work_tasks: List[PlanTask] = []
        self._life_tasks: List[PlanTask] = []
        # Snapshot of the backlog right after the last load, for unsaved-change
        # detection (parallel to _initial_tasks_snapshot).
        self._initial_backlog_snapshot: tuple = ((), ())
        # Transient drag state shared across the three lists (see _DragDropList).
        self._drag_source: Optional[QListWidget] = None
        self._drag_row: int = -1
        self._drag_payload: Optional[PlanTask] = None
        self._original_in_progress: Optional[str] = None
        self._closing = False
        # Snapshot of the plan state right after the last load from Excel.
        # Used by `_has_unsaved_changes` so Esc / 关闭 can warn the user before
        # silently dropping their edits.
        self._initial_tasks_snapshot: List[tuple] = []
        # When the user picks "不保存" in the confirmation dialog we still
        # need closeEvent to skip the auto-save it would normally do.
        self._skip_save_on_close = False
        self._yesterday_unfinished: List[PlanTask] = []
        # Subset of _yesterday_unfinished that we actually render in the
        # carry-over card right now (after filtering out items the user
        # already has on today's list). Kept aligned 1-to-1 with
        # _yesterday_checkboxes so `_carry_over_now` can pair them via zip().
        self._visible_yesterday: List[PlanTask] = []
        # Yesterday rows the user just marked as DONE (because they forgot to
        # tick them yesterday). The row numbers are persisted via
        # PlanService.mark_plan_done() during the next save. Tracking them
        # also lets `_has_unsaved_changes` flag this kind of edit.
        self._yesterday_marked_done_rows: set = set()
        # Cache of the latest snapshot dict passed into _render_yesterday so
        # we can re-render after today-list mutations without re-reading Excel.
        self._yesterday_snap: dict = {
            "unfinished": [], "done": 0, "total": 0, "pomodoros": 0,
        }
        self._yesterday_checkboxes: List[QCheckBox] = []
        # Tracks Python id() of PlanTask objects we just created via the UI.
        # Used so we can pop the duration picker exactly once after the user
        # finishes typing the new task's name (rather than nagging on every edit).
        self._pending_duration_prompt: set = set()
        self._suppress_advance_once = False

        # Undo / redo. Each entry is a deep snapshot of self._tasks taken
        # right BEFORE a user-initiated mutation. Ctrl+Z pops from undo and
        # pushes the current state onto redo; Ctrl+Y / Ctrl+Shift+Z reverses.
        # Any new mutation clears redo (standard editor semantics).
        self._undo_stack: List[List[PlanTask]] = []
        self._redo_stack: List[List[PlanTask]] = []
        # Guard so undo/redo themselves don't recursively push onto their own
        # stacks while mutating _tasks.
        self._in_undo_redo = False

        self._build_ui()
        self._install_shortcuts()

    # ---------- public API ----------

    def set_habits(self, habits: List[str]) -> None:
        self._habits = list(habits)
        self._render_habits()

    def refresh_plan_from_excel(self) -> None:
        from shouyu.service.excel import KbExcel

        excel = None
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

        # Backlog pools (cross-day). Read on every open so the two sections
        # reflect the on-disk pool; the in-memory copy is the working set until
        # the next save (same open-time-snapshot model as the plan area).
        try:
            if excel is None:
                excel = KbExcel()
            self._work_tasks = excel.backlog_service(TaskCategory.WORK).read()
            self._life_tasks = excel.backlog_service(TaskCategory.LIFE).read()
        except Exception:
            logging.exception("failed to read backlog from Excel")
            self._work_tasks = []
            self._life_tasks = []

        # Capture the just-loaded state so we can detect unsaved edits later.
        self._initial_tasks_snapshot = self._snapshot_tasks(self._tasks)
        self._initial_backlog_snapshot = (
            self._snapshot_tasks(self._work_tasks),
            self._snapshot_tasks(self._life_tasks),
        )
        # Re-opening the dialog wipes undo history — otherwise Ctrl+Z would
        # reach back across reload boundaries and try to restore PlanTask
        # objects whose row/id may have shifted on disk.
        self._undo_stack.clear()
        self._redo_stack.clear()

        self._render_tasks()
        self._render_backlogs()
        self._update_stats()

        # Yesterday & streak: cheap enough to fetch on every open.
        snap = _read_yesterday_snapshot()
        self._yesterday_unfinished = list(snap["unfinished"])
        self._yesterday_marked_done_rows = set()
        self._render_yesterday(snap)
        self._render_streak()

    def show_fullscreen(self) -> None:
        """Open the dialog full-screen but respect the Windows taskbar.

        We deliberately use availableGeometry() (which excludes the taskbar)
        instead of showFullScreen() so the bottom action buttons (Esc / 开始今天)
        are never covered by the OS task bar.
        """
        self._closing = False
        self._skip_save_on_close = False
        self._reset_action_buttons()
        self._update_header()
        self._apply_time_theme()

        screen = self.screen() or QApplication.primaryScreen()
        if screen is not None:
            rect = screen.availableGeometry()
            self.setGeometry(rect)
        self.show()
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
        body.addWidget(self._build_backlog_card())
        body.setStretchFactor(0, 1)
        body.setStretchFactor(1, 2)
        body.setStretchFactor(2, 2)
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
        # Explicit font-family avoids the emoji+CJK fallback issue that caused
        # the streak text to render as gibberish on some Windows setups.
        self.streak_label.setStyleSheet(
            'color: #FF8A4C; font-size: 14px; font-weight: 600; '
            'font-family: "Microsoft YaHei UI", "Segoe UI", "Segoe UI Emoji";'
        )
        row.addWidget(self.streak_label)

        close_btn = QPushButton("✕  关闭")
        close_btn.setToolTip("关闭今日任务窗口 (Esc)")
        close_btn.setMinimumSize(82, 34)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setAutoDefault(False)
        close_btn.setDefault(False)
        close_btn.setFocusPolicy(Qt.NoFocus)
        close_btn.setStyleSheet(
            "QPushButton {"
            "  font-size: 13px;"
            "  font-weight: 600;"
            "  background-color: rgba(255, 255, 255, 0.06);"
            "  border: 1px solid rgba(255, 255, 255, 0.18);"
            "  border-radius: 6px;"
            f"  color: {TEXT_COLOR_HEX};"
            "  padding: 4px 14px;"
            "  min-height: 0;"
            "}"
            "QPushButton:hover {"
            "  background-color: rgba(232, 17, 35, 0.85);"  # Windows close-red
            "  color: white;"
            "  border: 1px solid rgba(232, 17, 35, 0.95);"
            "}"
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

        sweep_btn = QPushButton("清理未完成 → Backlog")
        sweep_btn.setToolTip(
            "把今天所有未完成的任务按「工作/生活」分类移入 Backlog（可 Ctrl+Z 撤销）"
        )
        sweep_btn.setAutoDefault(False)
        sweep_btn.setCursor(Qt.PointingHandCursor)
        sweep_btn.setFocusPolicy(Qt.NoFocus)
        sweep_btn.setStyleSheet(
            "QPushButton {"
            "  background-color: rgba(255,255,255,0.06);"
            f"  color: {TEXT_COLOR_HEX};"
            "  border: 1px solid rgba(255,255,255,0.12);"
            "  border-radius: 4px;"
            "  padding: 3px 10px;"
            "  font-size: 12px;"
            "  min-height: 0;"
            "}"
            "QPushButton:hover { background-color: rgba(255,255,255,0.14); }"
        )
        sweep_btn.clicked.connect(self._sweep_unfinished_to_backlog)
        title_row.addWidget(sweep_btn)
        layout.addLayout(title_row)

        hint = QLabel(
            "↑↓ 选择 ·  F2 / 回车 编辑（编辑中再按回车跳到下一项） ·  "
            "Alt+↑↓ 重排 ·  Ctrl+ + 添加 ·  Ctrl+ − 删除 ·  "
            "Ctrl+Z 撤销 / Ctrl+Y 重做 ·  右键 → 状态 / 优先级 / 时长"
        )
        hint.setObjectName("HintLabel")
        hint.setWordWrap(True)
        layout.addWidget(hint)

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

        self.list_widget = _DragDropList(self)
        self._setup_task_list(self.list_widget, is_backlog=False)
        # Excel-like Enter: when the inline editor closes (Enter / Tab / focus
        # loss) jump to next row. Esc aborts and we leave the cursor put.
        self.list_widget.itemDelegate().closeEditor.connect(self._on_editor_closed)
        self.task_stack.addWidget(self.list_widget)

        self.empty_label = QLabel("🎉  今天还没有任务\n按 Ctrl+ + 添加一项重要的事")
        self.empty_label.setAlignment(Qt.AlignCenter)
        self.empty_label.setStyleSheet(
            f"color: {SUBTEXT_COLOR_HEX}; font-size: 14px; line-height: 22px;"
        )
        self.task_stack.addWidget(self.empty_label)

        layout.addWidget(self.task_stack_host, stretch=1)

        # Day-end review: estimated vs actual (pomodoros). Hidden until there's
        # data worth showing, so it doesn't clutter the morning view.
        self.review_label = QLabel("")
        self.review_label.setWordWrap(True)
        self.review_label.setStyleSheet(
            "padding: 8px 10px; "
            "background-color: rgba(15, 98, 254, 0.08); "
            "border: 1px solid rgba(15, 98, 254, 0.30); "
            "border-radius: 6px; "
            "font-size: 12px;"
        )
        self.review_label.setVisible(False)
        layout.addWidget(self.review_label)

        return self.task_card

    def _build_footer_row(self) -> QHBoxLayout:
        row = QHBoxLayout()
        row.setSpacing(10)

        tip = QLabel(
            "Esc 跳过 ·  Ctrl+Enter 开始今天 ·  Ctrl+Alt+H 重新打开今日任务"
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

    def _setup_task_list(self, lw: QListWidget, is_backlog: bool) -> None:
        """Shared config for the three drag-and-drop task lists."""
        lw.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked
            | QAbstractItemView.EditTrigger.EditKeyPressed
        )
        lw.setSelectionMode(QAbstractItemView.SingleSelection)
        lw.setContextMenuPolicy(Qt.CustomContextMenu)
        # Custom cross-widget drag (see _DragDropList): enable dragging out and
        # dropping in, but we own the reconciliation in _handle_drop.
        lw.setDragEnabled(True)
        lw.setAcceptDrops(True)
        lw.setDropIndicatorShown(True)
        lw.setDragDropMode(QAbstractItemView.DragDrop)
        lw.setDefaultDropAction(Qt.MoveAction)
        lw.setMovement(QListWidget.Snap)
        if is_backlog:
            lw.itemChanged.connect(self._on_backlog_text_edited)
            lw.customContextMenuRequested.connect(self._show_backlog_context_menu)
        else:
            lw.itemChanged.connect(self._on_item_text_edited)
            lw.customContextMenuRequested.connect(self._show_context_menu)

    def _build_backlog_card(self) -> QWidget:
        self.backlog_card = QFrame()
        self.backlog_card.setObjectName("Card")
        self.backlog_card.setStyleSheet(_card_qss(PANEL_COLOR_HEX))

        layout = QVBoxLayout(self.backlog_card)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(10)

        title = QLabel("Backlog")
        title.setObjectName("TitleLabel")
        layout.addWidget(title)

        hint = QLabel("拖到「今日要事」开始做 · 从今日拖回这里暂存 · 双击重命名 · 右键更多")
        hint.setObjectName("HintLabel")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        sections = QSplitter(Qt.Vertical)
        sections.setHandleWidth(2)
        sections.setChildrenCollapsible(False)
        sections.addWidget(self._build_backlog_section(TaskCategory.WORK))
        sections.addWidget(self._build_backlog_section(TaskCategory.LIFE))
        sections.setStretchFactor(0, 1)
        sections.setStretchFactor(1, 1)
        layout.addWidget(sections, stretch=1)

        return self.backlog_card

    def _build_backlog_section(self, category: TaskCategory) -> QWidget:
        container = QWidget()
        v = QVBoxLayout(container)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(6)

        header_row = QHBoxLayout()
        header_row.setSpacing(6)
        header = QLabel("")
        header.setStyleSheet(f"font-weight: 600; color: {TEXT_COLOR_HEX};")
        header_row.addWidget(header)
        header_row.addStretch(1)

        add_btn = QPushButton("+ 新增")
        add_btn.setAutoDefault(False)
        add_btn.setCursor(Qt.PointingHandCursor)
        add_btn.setStyleSheet(
            "QPushButton {"
            "  background-color: rgba(255,255,255,0.06);"
            f"  color: {TEXT_COLOR_HEX};"
            "  border: 1px solid rgba(255,255,255,0.12);"
            "  border-radius: 4px;"
            "  padding: 2px 10px;"
            "  font-size: 12px;"
            "  min-height: 0;"
            "}"
            "QPushButton:hover { background-color: rgba(255,255,255,0.14); }"
        )
        header_row.addWidget(add_btn)
        v.addLayout(header_row)

        lw = _DragDropList(self)
        self._setup_task_list(lw, is_backlog=True)
        lw.setMinimumHeight(120)
        v.addWidget(lw, stretch=1)

        if category == TaskCategory.WORK:
            self.work_list = lw
            self.work_header = header
            add_btn.clicked.connect(lambda: self._backlog_add(self.work_list))
        else:
            self.life_list = lw
            self.life_header = header
            add_btn.clicked.connect(lambda: self._backlog_add(self.life_list))

        return container

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

    # ---------- backlog rendering / mapping ----------

    def _tasks_for(self, lw: QListWidget) -> List[PlanTask]:
        """Return the backing list for one of the three drag-and-drop lists."""
        if lw is self.list_widget:
            return self._tasks
        if lw is self.work_list:
            return self._work_tasks
        if lw is self.life_list:
            return self._life_tasks
        return []

    def _category_for(self, lw: QListWidget) -> Optional[TaskCategory]:
        """The category a list represents, or None for the today list."""
        if lw is self.work_list:
            return TaskCategory.WORK
        if lw is self.life_list:
            return TaskCategory.LIFE
        return None

    def _list_for_category(self, category: TaskCategory) -> QListWidget:
        return self.work_list if category == TaskCategory.WORK else self.life_list

    def _render_backlogs(self) -> None:
        self._render_backlog(self.work_list, self._work_tasks)
        self._render_backlog(self.life_list, self._life_tasks)
        self.work_header.setText(
            f"🏢 工作 ({sum(1 for t in self._work_tasks if (t.text or '').strip())})"
        )
        self.life_header.setText(
            f"🏠 生活 ({sum(1 for t in self._life_tasks if (t.text or '').strip())})"
        )

    def _render_backlog(self, lw: QListWidget, tasks: List[PlanTask]) -> None:
        lw.blockSignals(True)
        lw.clear()
        for task in tasks:
            item = QListWidgetItem(_format_backlog(task))
            flags = item.flags()
            flags |= Qt.ItemIsEditable | Qt.ItemIsDragEnabled
            flags &= ~Qt.ItemIsDropEnabled  # drop between rows, not into one
            item.setFlags(flags)
            item.setData(_TASK_ROLE, task)
            item.setForeground(QColor(PENDING_COLOR_HEX))
            reflection = (task.reflection or "").strip()
            if reflection:
                item.setToolTip(f"📝 反思\n\n{reflection}")
            lw.addItem(item)
        lw.blockSignals(False)

    def _render_yesterday(self, snap: dict) -> None:
        # Cache the snapshot so `_refresh_carryover_visibility` can re-render
        # later without re-reading Excel.
        self._yesterday_snap = snap

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

        # Filter out yesterday items that already exist in today's task list
        # (matched by stripped text). This avoids the carry-over card nagging
        # the user about tasks they've already brought forward — whether by
        # the "结转选中 →" button, by typing them in manually, or by leaving
        # them in today's plan from a previous session.
        today_texts = {
            (t.text or "").strip()
            for t in self._tasks
            if (t.text or "").strip()
        }
        self._visible_yesterday = [
            t for t in self._yesterday_unfinished
            if (t.text or "").strip()
            and (t.text or "").strip() not in today_texts
        ]

        if not self._visible_yesterday:
            self.carryover_card.setVisible(False)
            return

        header = QHBoxLayout()
        header.setSpacing(6)
        header_label = QLabel(f"📅 昨日还有 {len(self._visible_yesterday)} 项未完成")
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

        # Per-row mark-done button styling — green tint to read as a
        # confirmation/completion affordance, distinct from the blue carry-over
        # primary button.
        done_btn_qss = (
            "QPushButton {"
            "  background-color: rgba(16, 124, 16, 0.18);"
            f"  color: {TEXT_COLOR_HEX};"
            "  border: 1px solid rgba(16, 124, 16, 0.45);"
            "  border-radius: 4px;"
            "  padding: 2px 8px;"
            "  font-size: 11px;"
            "  min-height: 0;"
            "}"
            "QPushButton:hover { background-color: rgba(16, 124, 16, 0.32); }"
        )

        for task in self._visible_yesterday:
            row = QHBoxLayout()
            row.setSpacing(6)

            cb = QCheckBox(task.text or "（空任务）")
            cb.setChecked(True)
            cb.setStyleSheet(f"QCheckBox {{ color: {TEXT_COLOR_HEX}; }}")
            self._yesterday_checkboxes.append(cb)
            row.addWidget(cb, stretch=1)

            done_btn = QPushButton("✓ 已完成")
            done_btn.setToolTip(
                "把昨天这条任务标记为「已完成」（保存后写回昨天的工作表）"
            )
            done_btn.setAutoDefault(False)
            done_btn.setCursor(Qt.PointingHandCursor)
            done_btn.setStyleSheet(done_btn_qss)
            # Snapshot the PlanTask reference so the lambda doesn't capture
            # the loop variable by name.
            done_btn.clicked.connect(
                lambda _checked=False, t=task: self._mark_yesterday_done(t)
            )
            # If the task didn't carry a row (shouldn't normally happen — read
            # via PlanService always populates row), hide the button to avoid
            # silently doing nothing.
            if not task.row:
                done_btn.setEnabled(False)
                done_btn.setToolTip("（无法定位到昨日工作表的对应行）")
            row.addWidget(done_btn)

            self.carryover_layout.addLayout(row)

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
        # Window-wide shortcuts. NOTE: we intentionally do NOT bind plain Return
        # / Enter here — those must reach the inline editor when one is open.
        # The "Enter to start editing the current row" affordance is bound to
        # the list_widget itself with WidgetShortcut context below.
        #
        # Space-to-cycle-status was removed on purpose: users found it too
        # easy to flip a task's state accidentally while navigating. Status
        # changes now go exclusively through the right-click context menu.
        window_bindings = [
            ("F2", self._edit_selected),
            ("Alt+Up", lambda: self._move(-1)),
            ("Alt+Down", lambda: self._move(1)),
            ("Ctrl++", self._add_new_task),
            ("Ctrl+=", self._add_new_task),
            ("Ctrl+Shift+=", self._add_new_task),
            ("Ctrl+-", self._delete_selected),
            ("Ctrl+Minus", self._delete_selected),
            ("Ctrl+Return", self._save_and_accept),
            ("Ctrl+Enter", self._save_and_accept),
            ("Ctrl+P", self._focus_pomodoro_on_selected),
            ("Ctrl+Z", self._undo),
            ("Ctrl+Y", self._redo),
            ("Ctrl+Shift+Z", self._redo),
            ("Escape", self.reject),
        ]
        for sequence, callback in window_bindings:
            shortcut = QShortcut(QKeySequence(sequence), self)
            shortcut.setContext(Qt.WindowShortcut)
            shortcut.activated.connect(callback)

        # Plain Enter on the list widget itself starts editing the current row.
        # WidgetShortcut means it only fires when list_widget has focus, NOT
        # when its inline editor (a child QLineEdit) has focus — so the editor
        # is free to consume Enter normally and commit.
        for sequence in ("Return", "Enter"):
            sc = QShortcut(QKeySequence(sequence), self.list_widget)
            sc.setContext(Qt.WidgetShortcut)
            sc.activated.connect(self._edit_selected)

    # ---------- list interactions ----------

    @staticmethod
    def _strip_glyph(text: str) -> str:
        text = text or ""
        for glyph in _GLYPH.values():
            if text.startswith(glyph):
                text = text[len(glyph):]
                break
        text = text.lstrip()
        # Strip trailing badges we appended ("🔴", "⏱", "🎯", "📝") so the
        # text the editor commits doesn't carry the decorations back into
        # PlanTask.text. Find earliest match — everything from there
        # onwards is decoration.
        first_marker = -1
        for marker in (
            "   🎯", "   ⏱", "   🔴", "   🟡", "   ⚪", "   📝",
            "   🏢", "   🏠", "   ⏳",
        ):
            idx = text.find(marker)
            if idx >= 0 and (first_marker < 0 or idx < first_marker):
                first_marker = idx
        if first_marker >= 0:
            text = text[:first_marker]
        return text.strip()

    def _on_item_text_edited(self, item: QListWidgetItem) -> None:
        """Called whenever the underlying item text changes. Keeps self._tasks
        in sync but does NOT advance the cursor — that happens from the
        delegate's `closeEditor` signal (so we know the editor really closed)."""
        index = self.list_widget.row(item)
        if 0 <= index < len(self._tasks):
            new_text = self._strip_glyph(item.text())
            if self._tasks[index].text == new_text:
                # _refresh_item runs unconditionally below so we still call it
                # to clean up any stray decorations the editor may have left.
                self._refresh_item(index)
                return
            self._push_undo()
            self._tasks[index].text = new_text
            self._refresh_item(index)
            self._update_stats()

    def _on_editor_closed(self, _editor, hint) -> None:
        """Excel-like: after the user commits an inline edit, jump editing to
        the next row. On the last row we stop (no auto-added blank row) — use
        Ctrl+ + to add another task."""
        if hint == QAbstractItemDelegate.EndEditHint.RevertModelCache:
            # User pressed Esc. Don't advance.
            return
        if self._suppress_advance_once:
            # Avoid runaway advance when we've just programmatically committed
            # an edit (e.g. from the duration picker re-render).
            self._suppress_advance_once = False
            return
        # Defer so Qt finishes tearing down the previous editor first.
        QTimer.singleShot(0, self._advance_editing_after_commit)

    def _advance_editing_after_commit(self) -> None:
        index = self.list_widget.currentRow()

        # If the task we just edited is brand new, has content, and still has
        # no duration, give the user a chance to commit a time budget first.
        # This is the one place where it's worth interrupting the Excel-like
        # flow (it's literally what they asked for).
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
                    current=30,  # Sensible default — most tasks are about 30min.
                    task_text=task.text,
                    parent=self,
                )
                if value > 0:
                    self._push_undo()
                    task.duration_minutes = value
                    self._refresh_item(index)
                    self._update_stats()
                # Whether user picked a duration or skipped, fall through to
                # normal advance behavior below.

        next_row = index + 1
        if next_row < self.list_widget.count():
            self.list_widget.setCurrentRow(next_row)
            self._edit_row(next_row)
        # On the last row we deliberately do NOT auto-append a new blank row:
        # committing the final task just ends editing. Use Ctrl+ + (or the
        # right-click menu) to add another task.

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
        reflection = (task.reflection or "").strip()
        item.setToolTip(f"📝 反思\n\n{reflection}" if reflection else "")
        self.list_widget.blockSignals(False)

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
                category=t.category,
                created_date=t.created_date,
            )
            for t in tasks
        ]

    def _snapshot_all(self) -> tuple:
        """Deep snapshot of all three lists (today / work / life) so a single
        Ctrl+Z can undo a cross-list drag as one atomic step."""
        return (
            self._clone_tasks(self._tasks),
            self._clone_tasks(self._work_tasks),
            self._clone_tasks(self._life_tasks),
        )

    def _restore_all(self, snap: tuple) -> None:
        today, work, life = snap
        self._tasks = today
        self._work_tasks = work
        self._life_tasks = life
        self._original_in_progress = next(
            (t.text for t in self._tasks if t.status == TaskStatus.IN_PROGRESS),
            None,
        )
        self._render_tasks()
        self._render_backlogs()
        self._update_stats()

    def _push_undo(self) -> None:
        """Snapshot all three lists before a user-initiated mutation.

        Called from every public mutation entry point (add / delete / move /
        edit / status / priority / duration / reflection / carry-over /
        focus-pomodoro / drag between lists / backlog edits). Mutations
        triggered while replaying an undo or redo are skipped so the two
        stacks don't fight.
        """
        if self._in_undo_redo:
            return
        self._undo_stack.append(self._snapshot_all())
        # Cap the history so a long editing session can't bloat memory. 100
        # steps is plenty for this kind of UI.
        if len(self._undo_stack) > 100:
            del self._undo_stack[0:len(self._undo_stack) - 100]
        self._redo_stack.clear()

    def _undo(self) -> None:
        if not self._undo_stack:
            return
        self._in_undo_redo = True
        try:
            self._redo_stack.append(self._snapshot_all())
            self._restore_all(self._undo_stack.pop())
        finally:
            self._in_undo_redo = False

    def _redo(self) -> None:
        if not self._redo_stack:
            return
        self._in_undo_redo = True
        try:
            self._undo_stack.append(self._snapshot_all())
            self._restore_all(self._redo_stack.pop())
        finally:
            self._in_undo_redo = False

    # ---------- status / mutation ----------

    def _set_status(self, new_status: TaskStatus) -> None:
        """Set the selected task's status. Replaces the old Space-cycle flow.

        When transitioning into DONE we offer a (skippable) reflection prompt
        so users can capture lessons learned right at the moment the task is
        marked complete.
        """
        index = self.list_widget.currentRow()
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
        self._render_tasks()
        self.list_widget.setCurrentRow(index)
        self._update_stats()
        if new_status == TaskStatus.DONE:
            # Prompt is intentionally non-blocking-feeling: cancelling just
            # skips the reflection, the DONE state is already committed.
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
        # Treat reflection edits as a separate undoable step so users can
        # Ctrl+Z back to "DONE but no reflection" without losing the status flip.
        self._push_undo()
        task.reflection = new_reflection
        self._refresh_item(index)

    def _move(self, offset: int) -> None:
        index = self.list_widget.currentRow()
        new_index = index + offset
        if not (0 <= index < len(self._tasks)) or not (0 <= new_index < len(self._tasks)):
            return
        self._push_undo()
        self._tasks[index], self._tasks[new_index] = self._tasks[new_index], self._tasks[index]
        self._render_tasks()
        self.list_widget.setCurrentRow(new_index)

    def _add_new_task(self, text: str = "") -> None:
        self._push_undo()
        new_task = PlanTask(text=text, status=TaskStatus.PENDING)
        self._pending_duration_prompt.add(id(new_task))
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
        # No confirmation dialog by design — Ctrl+Z restores the task,
        # which is faster and less interruptive than a yes/no popup.
        index = self.list_widget.currentRow()
        if not (0 <= index < len(self._tasks)):
            return
        self._push_undo()
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
            task = self._tasks[self.list_widget.currentRow()] \
                if 0 <= self.list_widget.currentRow() < len(self._tasks) else None
            menu.addAction("编辑  (F2 / Enter)", self._edit_selected)
            # Status submenu replaces the old Space cycle. Each item is a
            # direct "go to this state" — clearer than a one-way cycle.
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
                menu.addAction("📝 编辑反思…", lambda: self._edit_reflection_for_selected())
            menu.addAction("🍅 专注此项  (Ctrl+P)", self._focus_pomodoro_on_selected)
            menu.addAction("⏱ 设置时长…", self._set_duration_for_selected)
            backlog_menu = menu.addMenu("📥 移入 Backlog")
            backlog_menu.addAction("🏢 工作", lambda: self._move_today_to_backlog(TaskCategory.WORK))
            backlog_menu.addAction("🏠 生活", lambda: self._move_today_to_backlog(TaskCategory.LIFE))
            priority_menu = menu.addMenu("🚦 优先级")
            priority_menu.addAction("🔴 P1  必做", lambda: self._set_priority(TaskPriority.P1))
            priority_menu.addAction("🟡 P2  应做", lambda: self._set_priority(TaskPriority.P2))
            priority_menu.addAction("⚪ P3  可做", lambda: self._set_priority(TaskPriority.P3))
            priority_menu.addSeparator()
            priority_menu.addAction("清除优先级", lambda: self._set_priority(TaskPriority.NONE))
            menu.addAction("上移  (Alt+↑)", lambda: self._move(-1))
            menu.addAction("下移  (Alt+↓)", lambda: self._move(1))
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

    def _edit_reflection_for_selected(self) -> None:
        self._prompt_reflection_for(self.list_widget.currentRow())

    def _set_priority(self, priority: TaskPriority) -> None:
        index = self.list_widget.currentRow()
        if not (0 <= index < len(self._tasks)):
            return
        if self._tasks[index].priority == priority:
            return
        self._push_undo()
        self._tasks[index].priority = priority
        self._refresh_item(index)
        self._update_stats()

    def _set_duration_for_selected(self) -> None:
        index = self.list_widget.currentRow()
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

    # ---------- carry-over ----------

    def _set_carryover_all(self, checked: bool) -> None:
        for cb in self._yesterday_checkboxes:
            cb.setChecked(checked)

    def _mark_yesterday_done(self, task: PlanTask) -> None:
        """Tag a yesterday task as DONE — for the common case where the user
        actually finished it but forgot to tick the box yesterday.

        We only stage the change in memory here. The actual write to
        yesterday's worksheet happens in `_persist_plan_in_background` so the
        edit is properly retried/recovered like the rest of the save path.
        """
        if not task.row:
            return  # defensive — button should already be disabled in this case
        self._yesterday_marked_done_rows.add(task.row)
        # Drop the task from today's carry-over view so it stops nagging the
        # user. We match by row (unique within the day) rather than text so
        # duplicates with the same name don't both disappear.
        self._yesterday_unfinished = [
            t for t in self._yesterday_unfinished if t.row != task.row
        ]
        self._render_yesterday(self._yesterday_snap)
        self._update_stats()

    def _carry_over_now(self) -> None:
        if not self._visible_yesterday:
            return
        chosen = [
            task
            for task, cb in zip(self._visible_yesterday, self._yesterday_checkboxes)
            if cb.isChecked()
        ]
        if not chosen:
            self._yesterday_unfinished = []
            self._render_yesterday({"unfinished": [], "done": 0, "total": 0, "pomodoros": 0})
            return
        existing_texts = {(t.text or "").strip() for t in self._tasks if (t.text or "").strip()}
        will_add = [
            t for t in chosen
            if (t.text or "").strip() and (t.text or "").strip() not in existing_texts
        ]
        if will_add:
            self._push_undo()
        added = 0
        for task in will_add:
            text = (task.text or "").strip()
            self._tasks.append(
                PlanTask(
                    text=task.text,
                    status=TaskStatus.PENDING,
                    duration_minutes=task.duration_minutes,
                    priority=task.priority,
                    reflection=task.reflection,
                    category=task.category,
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
        if task.status != TaskStatus.IN_PROGRESS:
            self._push_undo()
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

    # ---------- drag & drop between the three lists ----------

    def _begin_drag(self, source_list: QListWidget, item: QListWidgetItem) -> None:
        row = source_list.row(item)
        tasks = self._tasks_for(source_list)
        if not (0 <= row < len(tasks)):
            self._clear_drag()
            return
        self._drag_source = source_list
        self._drag_row = row
        self._drag_payload = tasks[row]

    def _clear_drag(self) -> None:
        self._drag_source = None
        self._drag_row = -1
        self._drag_payload = None

    def _handle_drop(self, target_list: QListWidget, target_row: int) -> None:
        source_list = self._drag_source
        payload = self._drag_payload
        src_index = self._drag_row
        # Always clear transient drag state first, even on early return.
        self._clear_drag()
        if source_list is None or payload is None:
            return
        src_tasks = self._tasks_for(source_list)
        tgt_tasks = self._tasks_for(target_list)
        # Guard against the model shifting under us between drag start and drop.
        if not (0 <= src_index < len(src_tasks)) or src_tasks[src_index] is not payload:
            self._render_tasks()
            self._render_backlogs()
            return

        self._push_undo()
        task = src_tasks.pop(src_index)
        # Removing from the same list shifts indices above the removal point.
        if source_list is target_list and src_index < target_row:
            target_row -= 1
        target_row = max(0, min(target_row, len(tgt_tasks)))

        if source_list is not target_list:
            target_category = self._category_for(target_list)
            if target_category is not None:
                # Into a backlog pool: stamp category + entry date (keep an
                # existing date on work<->life moves), pool is all pending.
                task.category = target_category
                task.status = TaskStatus.PENDING
                if not (task.created_date or "").strip():
                    task.created_date = time.strftime("%Y-%m-%d")
            else:
                # Into today: keep its category, come in pending, drop the
                # pool-only entry date.
                task.status = TaskStatus.PENDING
                task.created_date = ""

        tgt_tasks.insert(target_row, task)
        self._render_tasks()
        self._render_backlogs()
        self._update_stats()
        if target_list is self.list_widget:
            self.list_widget.setCurrentRow(target_row)
        else:
            target_list.setCurrentRow(target_row)

    # ---------- backlog editing / context menu ----------

    def _on_backlog_text_edited(self, item: QListWidgetItem) -> None:
        lw = self.sender()
        if not isinstance(lw, QListWidget):
            return
        tasks = self._tasks_for(lw)
        index = lw.row(item)
        if not (0 <= index < len(tasks)):
            return
        new_text = self._strip_glyph(item.text())
        if tasks[index].text != new_text:
            self._push_undo()
            tasks[index].text = new_text
        # Re-render to normalize decorations / refresh the section count.
        self._render_backlogs()

    def _edit_backlog_row(self, lw: QListWidget, index: int) -> None:
        if 0 <= index < lw.count():
            item = lw.item(index)
            if item is not None:
                lw.editItem(item)

    def _backlog_add(self, lw: QListWidget) -> None:
        category = self._category_for(lw)
        if category is None:
            return
        tasks = self._tasks_for(lw)
        self._push_undo()
        new_task = PlanTask(
            text="",
            status=TaskStatus.PENDING,
            category=category,
            created_date=time.strftime("%Y-%m-%d"),
        )
        index = lw.currentRow()
        if 0 <= index < len(tasks):
            tasks.insert(index + 1, new_task)
            target = index + 1
        else:
            tasks.append(new_task)
            target = len(tasks) - 1
        self._render_backlogs()
        lw.setCurrentRow(target)
        QTimer.singleShot(0, lambda r=target: self._edit_backlog_row(lw, r))

    def _backlog_delete(self, lw: QListWidget) -> None:
        tasks = self._tasks_for(lw)
        index = lw.currentRow()
        if not (0 <= index < len(tasks)):
            return
        self._push_undo()
        tasks.pop(index)
        self._render_backlogs()
        if tasks:
            lw.setCurrentRow(min(index, len(tasks) - 1))

    def _backlog_set_priority(self, lw: QListWidget, priority: TaskPriority) -> None:
        tasks = self._tasks_for(lw)
        index = lw.currentRow()
        if not (0 <= index < len(tasks)) or tasks[index].priority == priority:
            return
        self._push_undo()
        tasks[index].priority = priority
        self._render_backlogs()

    def _backlog_set_duration(self, lw: QListWidget) -> None:
        tasks = self._tasks_for(lw)
        index = lw.currentRow()
        if not (0 <= index < len(tasks)):
            return
        task = tasks[index]
        value = DurationPickerDialog.get_duration(
            current=task.duration_minutes,
            task_text=task.text or '（空任务）',
            parent=self,
        )
        if value < 0 or value == task.duration_minutes:
            return
        self._push_undo()
        task.duration_minutes = value
        self._render_backlogs()

    def _backlog_change_category(self, lw: QListWidget) -> None:
        src_cat = self._category_for(lw)
        if src_cat is None:
            return
        tasks = self._tasks_for(lw)
        index = lw.currentRow()
        if not (0 <= index < len(tasks)):
            return
        target_cat = TaskCategory.LIFE if src_cat == TaskCategory.WORK else TaskCategory.WORK
        self._push_undo()
        task = tasks.pop(index)
        task.category = target_cat
        # Decision I: keep created_date so the "已搁置 N 天" clock is preserved.
        self._tasks_for(self._list_for_category(target_cat)).append(task)
        self._render_backlogs()

    def _backlog_pull_to_today(self, lw: QListWidget) -> None:
        tasks = self._tasks_for(lw)
        index = lw.currentRow()
        if not (0 <= index < len(tasks)):
            return
        self._push_undo()
        task = tasks.pop(index)
        task.status = TaskStatus.PENDING
        task.created_date = ""
        self._tasks.append(task)
        self._render_tasks()
        self._render_backlogs()
        self._update_stats()
        self.list_widget.setCurrentRow(len(self._tasks) - 1)

    def _move_today_to_backlog(self, category: TaskCategory) -> None:
        index = self.list_widget.currentRow()
        if not (0 <= index < len(self._tasks)):
            return
        self._push_undo()
        task = self._tasks.pop(index)
        task.status = TaskStatus.PENDING
        task.category = category
        if not (task.created_date or "").strip():
            task.created_date = time.strftime("%Y-%m-%d")
        self._tasks_for(self._list_for_category(category)).append(task)
        self._render_tasks()
        self._render_backlogs()
        self._update_stats()
        if self._tasks:
            self.list_widget.setCurrentRow(min(index, len(self._tasks) - 1))

    def _show_backlog_context_menu(self, pos: QPoint) -> None:
        lw = self.sender()
        if not isinstance(lw, QListWidget):
            return
        tasks = self._tasks_for(lw)
        item = lw.itemAt(pos)
        menu = QMenu(lw)
        if item is not None:
            lw.setCurrentItem(item)
            menu.addAction("编辑  (双击)", lambda: self._edit_backlog_row(lw, lw.currentRow()))
            menu.addAction("⬆ 拉到今日", lambda: self._backlog_pull_to_today(lw))
            other = "生活" if self._category_for(lw) == TaskCategory.WORK else "工作"
            menu.addAction(f"🔀 改为 {other}", lambda: self._backlog_change_category(lw))
            menu.addAction("⏱ 设置时长…", lambda: self._backlog_set_duration(lw))
            priority_menu = menu.addMenu("🚦 优先级")
            priority_menu.addAction("🔴 P1  必做", lambda: self._backlog_set_priority(lw, TaskPriority.P1))
            priority_menu.addAction("🟡 P2  应做", lambda: self._backlog_set_priority(lw, TaskPriority.P2))
            priority_menu.addAction("⚪ P3  可做", lambda: self._backlog_set_priority(lw, TaskPriority.P3))
            priority_menu.addSeparator()
            priority_menu.addAction("清除优先级", lambda: self._backlog_set_priority(lw, TaskPriority.NONE))
            menu.addSeparator()
        menu.addAction("新增任务", lambda: self._backlog_add(lw))
        if item is not None:
            menu.addAction("删除任务", lambda: self._backlog_delete(lw))
        menu.addSeparator()
        undo_act = menu.addAction("撤销  (Ctrl+Z)", self._undo)
        undo_act.setEnabled(bool(self._undo_stack))
        redo_act = menu.addAction("重做  (Ctrl+Y)", self._redo)
        redo_act.setEnabled(bool(self._redo_stack))
        menu.exec(lw.mapToGlobal(pos))

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

        life_count = sum(
            1 for t in self._tasks
            if (t.text or '').strip() and t.category == TaskCategory.LIFE
        )
        work_count = sum(
            1 for t in self._tasks
            if (t.text or '').strip() and t.category == TaskCategory.WORK
        )

        parts = [f"今日 {done}/{total} 完成"]
        if in_progress:
            parts.append(f"{in_progress} 进行中")
        if p1_count:
            parts.append(f"🔴 {p1_count} 项 P1")
        # Only surface the work/life split once there's at least one life task;
        # otherwise everything defaults to work and the badge is just noise.
        if life_count:
            parts.append(f"🏢 {work_count} · 🏠 {life_count}")
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
            self.stats_label.setStyleSheet(
                "color: #FFB454; font-weight: 600;"
            )
        elif unestimated > 0:
            self.stats_label.setStyleSheet(
                f"color: {SUBTEXT_COLOR_HEX};"
            )
        else:
            self.stats_label.setStyleSheet("")

        self.progress_bar.setRange(0, max(total, 1))
        self.progress_bar.setValue(done)
        self._render_review()
        self._refresh_carryover_visibility()

    def _refresh_carryover_visibility(self) -> None:
        """Re-render the yesterday-carryover card honoring the current today
        list. Items already in today (by stripped text) are filtered out;
        if everything is already there, the card hides itself.

        Cheap to call after every today-list mutation — that's why
        `_update_stats` (which already runs after every change) drives it.

        Short-circuits when the visible-set hasn't changed, to avoid
        clobbering the user's checkbox selections on every keystroke.
        """
        if not self._yesterday_unfinished:
            return
        today_texts = {
            (t.text or "").strip()
            for t in self._tasks
            if (t.text or "").strip()
        }
        new_visible_texts = {
            (t.text or "").strip()
            for t in self._yesterday_unfinished
            if (t.text or "").strip()
            and (t.text or "").strip() not in today_texts
        }
        cur_visible_texts = {
            (t.text or "").strip()
            for t in self._visible_yesterday
            if (t.text or "").strip()
        }
        if new_visible_texts == cur_visible_texts:
            return
        self._render_yesterday(self._yesterday_snap)

    def _render_review(self) -> None:
        """Day-end estimated-vs-actual recap.

        Compares total estimated minutes for tasks marked DONE against the
        actual logged-pomodoro time for today. Only renders if there are
        completed tasks AND at least one logged pomodoro (otherwise it's not
        meaningful — too early in the day).
        """
        try:
            from shouyu.service.excel import KbExcel
            from shouyu.service.pomodoro import PomodoroService

            pomodoros = KbExcel().plan_service().count_pomodoros_logged()
        except Exception:
            logging.exception("failed to compute day review")
            self.review_label.setVisible(False)
            return

        done_tasks = [t for t in self._tasks if t.status == TaskStatus.DONE]
        if not done_tasks or pomodoros <= 0:
            self.review_label.setVisible(False)
            return

        try:
            from shouyu.config import Config as _Config

            mode = PomodoroService.instance().mode()
            per_pomodoro_min = (
                _Config.pomodoro_deep_work_minutes()
                if mode == PomodoroService.MODE_DEEP
                else _Config.pomodoro_work_minutes()
            )
        except Exception:
            per_pomodoro_min = 25

        estimated = sum(
            t.duration_minutes for t in done_tasks if t.duration_minutes > 0
        )
        actual = pomodoros * per_pomodoro_min
        skipped = AppState.get_today_counter('breaks_skipped')

        lines = [
            f"📊 <b>今日复盘</b> · 完成 {len(done_tasks)} 项 · "
            f"番茄 {pomodoros} 个 ≈ {actual} 分钟"
        ]
        if estimated > 0:
            diff = actual - estimated
            pct = (diff / estimated) * 100 if estimated else 0
            if abs(pct) < 15:
                verdict = "👍 估时基本准确，继续保持"
            elif diff > 0:
                verdict = f"⏰ 实际比估时多 {diff} 分钟（{pct:+.0f}%），下次估保守一点"
            else:
                verdict = f"⚡ 实际比估时少 {-diff} 分钟（{pct:+.0f}%），可以多承接些任务"
            lines.append(
                f"已估时 {estimated} 分钟 vs 实际 {actual} 分钟 → {verdict}"
            )
        else:
            lines.append("💡 完成的任务没有估时，明天试试在新建任务时给个时间预算")

        if skipped >= 2:
            lines.append(
                f"☕ 今天跳了 {skipped} 次休息——长期看这会让下午效率掉。"
                "明天试试至少老实休息一次。"
            )

        self.review_label.setText("<br>".join(lines))
        self.review_label.setTextFormat(Qt.RichText)
        self.review_label.setVisible(True)

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
        # Previously this short-circuited via a `_save_already_dispatched`
        # flag — that meant any second call (e.g. close-after-focus-pomodoro)
        # silently dropped subsequent edits. Each dispatch now spawns its own
        # worker; the writes are idempotent so overlapping retries are fine.
        tasks_snapshot = self._clone_tasks(self._tasks)
        yesterday_done_rows = sorted(self._yesterday_marked_done_rows)
        _persist_plan_in_background(
            tasks_snapshot,
            self._original_in_progress,
            yesterday_done_rows,
            work_snapshot=self._clone_tasks(self._work_tasks),
            life_snapshot=self._clone_tasks(self._life_tasks),
        )

    # ---------- unsaved-change detection ----------

    @staticmethod
    def _snapshot_tasks(tasks: List[PlanTask]) -> List[tuple]:
        """Tuple-ize tasks so equality / ordering comparison is structural.

        Empty-text tasks are dropped so a stray empty row created by Excel-like
        auto-extend doesn't register as an "edit". Note that this matches the
        save path (`_persist_plan_in_background` also filters empties).
        """
        out: List[tuple] = []
        for t in tasks:
            if not (t.text or "").strip():
                continue
            out.append((
                (t.text or "").strip(),
                t.status,
                int(t.duration_minutes or 0),
                t.priority,
                (t.reflection or "").strip(),
                t.category,
            ))
        return out

    def _has_unsaved_changes(self) -> bool:
        if self._snapshot_tasks(self._tasks) != self._initial_tasks_snapshot:
            return True
        current_backlog = (
            self._snapshot_tasks(self._work_tasks),
            self._snapshot_tasks(self._life_tasks),
        )
        if current_backlog != self._initial_backlog_snapshot:
            return True
        if self._yesterday_marked_done_rows:
            return True
        return False

    def _confirm_unsaved_then_close(self) -> None:
        """Esc / 关闭 / 跳过 path: ask before discarding edits."""
        box = QMessageBox(self)
        box.setWindowTitle("保存今日任务变更？")
        box.setIcon(QMessageBox.Question)
        box.setText("今日任务列表有未保存的变更。")
        box.setInformativeText("是否保存这些变更后再关闭？")
        save_btn = box.addButton("保存并关闭", QMessageBox.AcceptRole)
        discard_btn = box.addButton("不保存", QMessageBox.DestructiveRole)
        cancel_btn = box.addButton("取消", QMessageBox.RejectRole)
        box.setDefaultButton(save_btn)
        box.setEscapeButton(cancel_btn)
        box.exec()

        clicked = box.clickedButton()
        if clicked is save_btn:
            # Same path as Ctrl+Enter — but we still want reject() semantics
            # afterwards (the morning ritual is "skipped" for streak purposes
            # only via _save_and_accept, so we mimic just the save half here).
            self._dispatch_save()
            super().reject()
        elif clicked is discard_btn:
            self._skip_save_on_close = True
            super().reject()
        # Cancel: do nothing, leave the window open.

    # ---------- qt overrides ----------

    def reject(self) -> None:
        # Esc, the "✕ 关闭" button, and the "跳过" button all funnel through
        # here. Only show the confirmation when there's actually something to
        # lose; otherwise close instantly so the dialog stays out of the way.
        # NOTE: sweeping unfinished tasks into the Backlog is a *manual* action
        # (the footer button), deliberately NOT auto-prompted on close — this
        # dialog is opened often just to glance at today's tasks, and nagging
        # on every close was annoying.
        if self._closing:
            super().reject()
            return
        if not self._has_unsaved_changes():
            super().reject()
            return
        self._confirm_unsaved_then_close()

    def _sweep_unfinished_to_backlog(self) -> None:
        """Manual sweep (footer button): move today's unfinished tasks into the
        Backlog pools, routed by each task's category and deduped by text.

        Deliberately NOT prompted on close — it's an explicit, user-initiated
        housekeeping action (and undoable via Ctrl+Z). Completed tasks stay in
        today as the day's record.
        """
        unfinished = [
            t for t in self._tasks
            if (t.text or "").strip()
            and t.status in (TaskStatus.PENDING, TaskStatus.IN_PROGRESS)
        ]
        if not unfinished:
            from shouyu.view.msgbox import MessageBox, MessageType

            MessageBox.pop_up_message(
                title="没有需要清理的任务",
                msg="今天没有未完成的任务。",
                level=MessageType.SUCCESS,
            )
            return

        self._push_undo()
        for task in unfinished:
            target = self._tasks_for(self._list_for_category(task.category))
            text = (task.text or "").strip()
            if any((x.text or "").strip() == text for x in target):
                continue  # already pooled — just drop it from today below
            task.status = TaskStatus.PENDING
            if not (task.created_date or "").strip():
                task.created_date = time.strftime("%Y-%m-%d")
            target.append(task)

        unfinished_ids = {id(t) for t in unfinished}
        self._tasks = [t for t in self._tasks if id(t) not in unfinished_ids]
        self._render_tasks()
        self._render_backlogs()
        self._update_stats()

    def closeEvent(self, event) -> None:
        self._closing = True
        if not self._skip_save_on_close:
            self._dispatch_save()
        event.accept()
