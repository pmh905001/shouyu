"""Floating mini-window that shows the current pomodoro phase / countdown.

The window stays on top, is frameless, and can be dragged. It receives
events from PomodoroService through QtApp.emit_pomodoro_event(...).
"""
from __future__ import annotations

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
        self.setFixedSize(260, 110)

        self._drag_offset: Optional[QPoint] = None

        self._build_ui()
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

        button_row = QHBoxLayout()
        button_row.setSpacing(6)
        self.toggle_btn = QPushButton("暂停")
        self.toggle_btn.setStyleSheet(self._button_style())
        self.toggle_btn.clicked.connect(self._on_toggle_clicked)
        button_row.addWidget(self.toggle_btn)

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
            self.show()
        elif event == "tick":
            try:
                remaining = int(payload)
            except ValueError:
                remaining = 0
            self._set_remaining(remaining)
        elif event == "paused":
            self._set_phase("paused")
            self.toggle_btn.setText("继续")
        elif event == "resumed":
            snap = PomodoroService.instance().snapshot()
            self._set_phase(snap["phase"])
            self.toggle_btn.setText("暂停")
        elif event == "stopped":
            self._set_phase("idle")
            self._set_remaining(0)
            self.hide()

        snapshot = PomodoroService.instance().snapshot()
        self.cycles_label.setText(f"🍅 {snapshot['completed_today']}")

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
            shown = text if len(text) <= 26 else text[:25] + "…"
            self.task_label.setText(f"→ {shown}")
        else:
            self.task_label.setText("")

    # ---------- button slots ----------

    def _on_toggle_clicked(self) -> None:
        from shouyu.service.pomodoro import PomodoroService

        PomodoroService.instance().toggle()

    def _on_stop_clicked(self) -> None:
        from shouyu.service.pomodoro import PomodoroService

        PomodoroService.instance().stop()

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
