"""Floating mini-window that shows the current pomodoro phase / countdown.

The window stays on top, is frameless, and can be dragged. It receives
events from PomodoroService through QtApp.emit_pomodoro_event(...).
"""
from __future__ import annotations

import random
from typing import Optional

from PySide6.QtCore import Qt, QPoint
from PySide6.QtGui import QMouseEvent
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from shouyu.view.styles import (
    ACCENT_COLOR_HEX,
    DONE_COLOR_HEX,
    IN_PROGRESS_COLOR_HEX,
    PANEL_COLOR_HEX,
    SUBTEXT_COLOR_HEX,
    TEXT_COLOR_HEX,
)


_PHASE_LABEL = {
    "idle": ("空闲", SUBTEXT_COLOR_HEX),
    "working": ("专注中", IN_PROGRESS_COLOR_HEX),
    "short_break": ("短休息", DONE_COLOR_HEX),
    "long_break": ("长休息", DONE_COLOR_HEX),
    "paused": ("已暂停", SUBTEXT_COLOR_HEX),
}


_BREAK_TIPS = [
    "🧘 起身活动脖子和肩膀",
    "👀 看远处 20 秒，让眼睛休息",
    "💧 喝口水，深呼吸 5 次",
    "🌳 看看窗外，放空一下",
    "🤸 伸展身体，活络血液",
    "☕ 给自己泡杯茶或咖啡",
    "🚶 走两步，离开座位 1 分钟",
]


_WORK_TIPS = [
    "🎯 一次只做一件事，关闭其他干扰",
    "✍️ 记下分心的念头，专注当下",
    "🤫 这是你不被打断的时间",
]


def _format_remaining(seconds: int) -> str:
    seconds = max(0, int(seconds))
    return f"{seconds // 60:02d}:{seconds % 60:02d}"


class PomodoroWindow(QWidget):
    _instance: Optional["PomodoroWindow"] = None

    @classmethod
    def get_or_create(cls) -> "PomodoroWindow":
        if cls._instance is None:
            cls._instance = PomodoroWindow()
        return cls._instance

    def __init__(self) -> None:
        super().__init__()
        self.setWindowFlag(Qt.FramelessWindowHint, True)
        self.setWindowFlag(Qt.WindowStaysOnTopHint, True)
        self.setWindowFlag(Qt.Tool, True)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setFixedSize(296, 168)

        self._drag_offset: Optional[QPoint] = None

        self._build_ui()
        self._refresh_mode_button()
        self._move_to_default_corner()

    def _build_ui(self) -> None:
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)

        self.card = QFrame()
        self.card.setStyleSheet(
            f"""
            QFrame {{
                background-color: {PANEL_COLOR_HEX};
                border-radius: 14px;
                border: 1px solid #3A3A3A;
            }}
            """
        )
        outer.addWidget(self.card)

        layout = QVBoxLayout(self.card)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(4)

        header_row = QHBoxLayout()
        self.phase_label = QLabel("空闲")
        self.phase_label.setStyleSheet(
            f"color: {SUBTEXT_COLOR_HEX}; font-size: 12px; font-weight: 600;"
        )
        header_row.addWidget(self.phase_label)

        self.mode_btn = QPushButton("经典")
        self.mode_btn.setToolTip("切换 经典 25/5 ↔ 深度 90/15")
        self.mode_btn.setCursor(Qt.PointingHandCursor)
        self.mode_btn.setFixedHeight(20)
        self.mode_btn.setStyleSheet(
            "QPushButton {"
            "  background-color: rgba(255,255,255,0.06);"
            "  color: #C9C9C9;"
            "  border: 1px solid rgba(255,255,255,0.16);"
            "  border-radius: 4px;"
            "  padding: 1px 8px;"
            "  font-size: 11px;"
            "  min-height: 0;"
            "}"
            "QPushButton:hover { background-color: rgba(255,255,255,0.14); color: white; }"
        )
        self.mode_btn.clicked.connect(self._on_toggle_mode_clicked)
        header_row.addWidget(self.mode_btn)
        header_row.addStretch(1)

        self.cycles_label = QLabel("🍅 0")
        self.cycles_label.setStyleSheet(f"color: {SUBTEXT_COLOR_HEX}; font-size: 12px;")
        header_row.addWidget(self.cycles_label)
        layout.addLayout(header_row)

        self.time_label = QLabel("00:00")
        self.time_label.setStyleSheet(
            f"color: {TEXT_COLOR_HEX}; font-size: 30px; font-weight: 700;"
        )
        layout.addWidget(self.time_label)

        self.task_label = QLabel("")
        self.task_label.setStyleSheet(f"color: {ACCENT_COLOR_HEX}; font-size: 12px;")
        self.task_label.setWordWrap(False)
        layout.addWidget(self.task_label)

        self.tip_label = QLabel("")
        self.tip_label.setStyleSheet(
            f"color: {SUBTEXT_COLOR_HEX}; font-size: 11px;"
        )
        self.tip_label.setWordWrap(True)
        self.tip_label.setVisible(False)
        layout.addWidget(self.tip_label)

        button_row = QHBoxLayout()
        button_row.setSpacing(6)
        self.toggle_btn = QPushButton("暂停")
        self.toggle_btn.setStyleSheet(self._button_style())
        self.toggle_btn.clicked.connect(self._on_toggle_clicked)
        button_row.addWidget(self.toggle_btn)

        self.extend_btn = QPushButton("+5m")
        self.extend_btn.setToolTip("延长当前阶段 5 分钟（在状态进入流时用）")
        self.extend_btn.setStyleSheet(self._button_style())
        self.extend_btn.clicked.connect(self._on_extend_clicked)
        self.extend_btn.setVisible(False)
        button_row.addWidget(self.extend_btn)

        self.skip_break_btn = QPushButton("跳过休息")
        self.skip_break_btn.setToolTip("跳过休息，回到专注（不推荐，会被记录）")
        self.skip_break_btn.setStyleSheet(self._button_style())
        self.skip_break_btn.clicked.connect(self._on_skip_break_clicked)
        self.skip_break_btn.setVisible(False)
        button_row.addWidget(self.skip_break_btn)

        stop_btn = QPushButton("停止")
        stop_btn.setStyleSheet(self._button_style())
        stop_btn.clicked.connect(self._on_stop_clicked)
        button_row.addWidget(stop_btn)

        hide_btn = QPushButton("隐藏")
        hide_btn.setStyleSheet(self._button_style())
        hide_btn.clicked.connect(self.hide)
        button_row.addWidget(hide_btn)

        layout.addLayout(button_row)

    @staticmethod
    def _button_style() -> str:
        return (
            "QPushButton {"
            "  background-color: rgba(255,255,255,0.08);"
            "  color: white;"
            "  border: none;"
            "  border-radius: 5px;"
            "  padding: 3px 8px;"
            "  font-size: 11px;"
            "}"
            "QPushButton:hover { background-color: rgba(255,255,255,0.18); }"
        )

    def _move_to_default_corner(self) -> None:
        screen = self.screen()
        if screen is None:
            return
        geometry = screen.availableGeometry()
        self.move(geometry.right() - self.width() - 24, geometry.bottom() - self.height() - 24)

    # ---------- event handling ----------

    def handle_event(self, event: str, payload: str) -> None:
        from shouyu.service.pomodoro import PomodoroService

        if event in ("started", "phase_changed"):
            phase, duration_str, *task_parts = (payload.split(":", 2) + ["", "", ""])[:3]
            try:
                duration = int(duration_str)
            except ValueError:
                duration = 0
            task_text = task_parts[0] if task_parts else ""
            self._set_phase(phase)
            self._set_remaining(duration)
            self._set_task(task_text)
            self._update_tip(phase)
            self.show()
        elif event == "tick":
            try:
                remaining = int(payload)
            except ValueError:
                remaining = 0
            self._set_remaining(remaining)
        elif event == "extended":
            # phase unchanged; remaining seconds will refresh on next tick
            pass
        elif event == "mode_changed":
            self._refresh_mode_button()
        elif event == "paused":
            self._set_phase("paused")
            self.toggle_btn.setText("继续")
        elif event == "resumed":
            snap = PomodoroService.instance().snapshot()
            self._set_phase(snap["phase"])
            self._update_tip(snap["phase"])
            self.toggle_btn.setText("暂停")
        elif event == "stopped":
            self._set_phase("idle")
            self._set_remaining(0)
            self._update_tip("idle")
            self.hide()

        snapshot = PomodoroService.instance().snapshot()
        cycles_text = f"🍅 {snapshot['completed_today']}"
        try:
            from shouyu.util.state import AppState

            skipped = AppState.get_today_counter('breaks_skipped')
            if skipped:
                cycles_text += f"  · 跳休 {skipped}"
        except Exception:
            pass
        self.cycles_label.setText(cycles_text)
        self._refresh_action_visibility(snapshot["phase"])

    def _set_phase(self, phase: str) -> None:
        label, color = _PHASE_LABEL.get(phase, ("空闲", SUBTEXT_COLOR_HEX))
        self.phase_label.setText(label)
        self.phase_label.setStyleSheet(
            f"color: {color}; font-size: 12px; font-weight: 600;"
        )

    def _set_remaining(self, seconds: int) -> None:
        self.time_label.setText(_format_remaining(seconds))

    def _set_task(self, text: str) -> None:
        if text:
            shown = text if len(text) <= 28 else text[:27] + "…"
            self.task_label.setText(f"→ {shown}")
        else:
            self.task_label.setText("")

    def _update_tip(self, phase: str) -> None:
        if phase in ("short_break", "long_break"):
            self.tip_label.setText(random.choice(_BREAK_TIPS))
            self.tip_label.setVisible(True)
        elif phase == "working":
            self.tip_label.setText(random.choice(_WORK_TIPS))
            self.tip_label.setVisible(True)
        else:
            self.tip_label.clear()
            self.tip_label.setVisible(False)

    def _refresh_action_visibility(self, phase: str) -> None:
        if phase == "working":
            self.extend_btn.setVisible(True)
            self.extend_btn.setText("+5m")
            self.skip_break_btn.setVisible(False)
        elif phase in ("short_break", "long_break"):
            self.extend_btn.setVisible(True)
            self.extend_btn.setText("休息+2m")
            self.skip_break_btn.setVisible(True)
        else:
            self.extend_btn.setVisible(False)
            self.skip_break_btn.setVisible(False)

    # ---------- button slots ----------

    def _on_toggle_clicked(self) -> None:
        from shouyu.service.pomodoro import PomodoroService

        PomodoroService.instance().toggle()

    def _on_stop_clicked(self) -> None:
        from shouyu.service.pomodoro import PomodoroService

        PomodoroService.instance().stop()

    def _on_extend_clicked(self) -> None:
        from shouyu.service.pomodoro import PomodoroService

        snap = PomodoroService.instance().snapshot()
        minutes = 2 if snap["phase"] in ("short_break", "long_break") else 5
        PomodoroService.instance().extend_current_phase(minutes)

    def _on_skip_break_clicked(self) -> None:
        from shouyu.service.pomodoro import PomodoroService

        PomodoroService.instance().skip_break()

    def _on_toggle_mode_clicked(self) -> None:
        from shouyu.service.pomodoro import PomodoroService

        svc = PomodoroService.instance()
        new_mode = (
            PomodoroService.MODE_DEEP
            if svc.mode() == PomodoroService.MODE_CLASSIC
            else PomodoroService.MODE_CLASSIC
        )
        svc.set_mode(new_mode)
        self._refresh_mode_button()

    def _refresh_mode_button(self) -> None:
        from shouyu.service.pomodoro import PomodoroService

        is_deep = PomodoroService.instance().mode() == PomodoroService.MODE_DEEP
        self.mode_btn.setText("深度" if is_deep else "经典")
        self.mode_btn.setToolTip(
            "当前: 深度 90/15 — 点击切回 经典 25/5"
            if is_deep
            else "当前: 经典 25/5 — 点击切到 深度 90/15"
        )

    # ---------- drag support ----------

    def mousePressEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.LeftButton:
            self._drag_offset = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event: QMouseEvent) -> None:
        if self._drag_offset is not None and event.buttons() & Qt.LeftButton:
            self.move(event.globalPosition().toPoint() - self._drag_offset)
            event.accept()
        else:
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event: QMouseEvent) -> None:
        self._drag_offset = None
        super().mouseReleaseEvent(event)
