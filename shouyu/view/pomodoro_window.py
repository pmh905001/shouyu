"""Floating mini-window that shows the current pomodoro phase / countdown.

The window stays on top, is frameless, and can be dragged. It receives
events from PomodoroService through QtApp.emit_pomodoro_event(...).
"""
from __future__ import annotations

import random
from typing import Optional

from PySide6.QtCore import Qt, QPoint, QTimer
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

        # Mirror of QWidget visibility, kept in sync via show/hideEvent.
        # Reading QWidget.isVisible() from non-Qt threads (the tray thread,
        # the keyboard hotkey thread) is not strictly safe; reading a plain
        # Python bool is. We use this flag for cross-thread checks like
        # `is_visible_safe()`.
        self._visible_flag = False

        # Idle-warning blink state. The QTimer ticks at 500ms and flips
        # `_blink_on` to drive a two-frame animation on the phase label
        # (text alternates between "⚠ 已静止 Nm" and "⚠ 请回来工作"; the
        # card border alternates between the alert red and a dimmer red).
        # Phase events from PomodoroService stop the blink and restore
        # normal styles.
        self._blink_timer = QTimer(self)
        self._blink_timer.setInterval(500)
        self._blink_timer.timeout.connect(self._on_blink_tick)
        self._blink_on = False
        self._blink_idle_seconds = 0
        self._idle_warning_active = False
        # Cached so we can restore exactly what we had after blink ends —
        # the phase may switch while idle (e.g. work → break would also
        # clear the warning), and we want to reflect whatever the latest
        # phase event told us, not a stale label.
        self._normal_phase_text = "空闲"
        self._normal_phase_color = SUBTEXT_COLOR_HEX

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
            # Any phase transition cancels a stuck blink — the service
            # already emits idle_warning_off on phase change, but stopping
            # locally too avoids a 1s window where stale alert styling is
            # left over from the previous phase.
            self._stop_blink()
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
            self._stop_blink()
            self._set_phase("paused")
            self.toggle_btn.setText("继续")
        elif event == "resumed":
            snap = PomodoroService.instance().snapshot()
            self._set_phase(snap["phase"])
            self._update_tip(snap["phase"])
            self.toggle_btn.setText("暂停")
        elif event == "stopped":
            self._stop_blink()
            self._set_phase("idle")
            self._set_remaining(0)
            self._update_tip("idle")
            self.hide()
        elif event == "idle_warning_on":
            try:
                idle_seconds = int(payload)
            except ValueError:
                idle_seconds = 0
            self._start_blink(idle_seconds)
        elif event == "idle_warning_off":
            self._stop_blink()

        snapshot = PomodoroService.instance().snapshot()
        # Don't clobber the blinking 🍅 — let the blink timer keep driving
        # cycles_label until idle_warning_off lands. Otherwise every tick
        # event would briefly restore the full "🍅 N" text and cause a
        # visible jitter.
        if not self._idle_warning_active:
            self.cycles_label.setText(self._current_cycles_text())
        self._refresh_action_visibility(snapshot["phase"])

    def _set_phase(self, phase: str) -> None:
        label, color = _PHASE_LABEL.get(phase, ("空闲", SUBTEXT_COLOR_HEX))
        # Remember "what the label should look like when not blinking", so
        # _stop_blink() can restore it without needing to re-poll the service.
        self._normal_phase_text = label
        self._normal_phase_color = color
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

    # ---------- idle-warning blink ----------

    _ALERT_RED = "#FF3B30"
    _ALERT_BORDER_DIM = "#7A1F1B"

    def _start_blink(self, idle_seconds: int) -> None:
        """Enter the "user is drifting off" alert mode.

        Visual:
          * card border switches to a steady red (so the whole window
            visually screams), so user notices it from peripheral vision
          * 🍅 emoji in `cycles_label` flickers on/off every 500ms — the
            literal "tomato is blinking" the user asked for
          * phase label changes to "⚠ 已静止 Nm" in red

        We deliberately do NOT raise / re-show the window if the user has
        manually hidden it: respecting the hide gesture is more important
        than nudging here.
        """
        self._idle_warning_active = True
        self._blink_idle_seconds = idle_seconds
        minutes = max(1, idle_seconds // 60)
        self.phase_label.setText(f"⚠ 已静止 {minutes}m")
        self.phase_label.setStyleSheet(
            f"color: {self._ALERT_RED}; font-size: 12px; font-weight: 700;"
        )
        self._apply_card_alert_style(True)
        self._blink_on = True
        self._on_blink_tick()  # flip immediately, don't wait 500ms
        self._blink_timer.start()

    def _stop_blink(self) -> None:
        if not self._idle_warning_active and not self._blink_timer.isActive():
            return
        self._blink_timer.stop()
        self._idle_warning_active = False
        # Restore phase label from what _set_phase last cached.
        self.phase_label.setText(self._normal_phase_text)
        self.phase_label.setStyleSheet(
            f"color: {self._normal_phase_color}; font-size: 12px; font-weight: 600;"
        )
        # Restore the cycles_label by re-deriving it (handle_event already
        # does this after every event, but on a fresh stop the user might
        # not see another event for a while — re-deriving keeps things
        # in sync immediately).
        self._refresh_cycles_label()
        self._apply_card_alert_style(False)

    def _on_blink_tick(self) -> None:
        # Toggle the 🍅 emoji visibility — text-level blink instead of
        # opacity, because Qt's QLabel doesn't get a setOpacity for free
        # without QGraphicsOpacityEffect (overkill for a one-glyph blink).
        snapshot_text = self._current_cycles_text()
        self._blink_on = not self._blink_on
        if self._blink_on:
            self.cycles_label.setText(snapshot_text)
        else:
            # Replace the leading 🍅 with a same-width-ish space so the
            # label width doesn't jump.
            self.cycles_label.setText(snapshot_text.replace("🍅", " ", 1))

    def _apply_card_alert_style(self, alert: bool) -> None:
        border_color = self._ALERT_RED if alert else "#3A3A3A"
        self.card.setStyleSheet(
            f"""
            QFrame {{
                background-color: {PANEL_COLOR_HEX};
                border-radius: 14px;
                border: 1px solid {border_color};
            }}
            """
        )

    def _current_cycles_text(self) -> str:
        """Compute the cycles_label text (🍅 N · 跳休 M) from current state.

        Mirrors what handle_event does after every event, so we can refresh
        it standalone (e.g. after stopping the blink).
        """
        from shouyu.service.pomodoro import PomodoroService

        try:
            snap = PomodoroService.instance().snapshot()
            text = f"🍅 {snap['completed_today']}"
        except Exception:
            text = "🍅 0"
        try:
            from shouyu.util.state import AppState

            skipped = AppState.get_today_counter('breaks_skipped')
            if skipped:
                text += f"  · 跳休 {skipped}"
        except Exception:
            pass
        return text

    def _refresh_cycles_label(self) -> None:
        self.cycles_label.setText(self._current_cycles_text())

    # ---------- summon (cross-thread show) ----------

    @classmethod
    def is_visible_safe(cls) -> bool:
        """Cross-thread safe visibility check.

        Reads the Python-level `_visible_flag` rather than calling Qt's
        isVisible(), so the tray and hotkey threads can use it without
        racing with the GUI thread.
        """
        inst = cls._instance
        return inst is not None and inst._visible_flag

    def summon(self) -> None:
        """Force the floating window back on screen and to the foreground.

        Used when the user has hidden the window and wants it back. If the
        window's saved position is now off-screen (e.g. monitor unplugged),
        we snap it back to the default corner before showing.
        """
        if not self._is_geometry_on_any_screen():
            self._move_to_default_corner()
        self.show()
        self.raise_()
        self.activateWindow()

    def toggle_visibility(self) -> None:
        """Hide if visible, otherwise summon. Used by the show/hide hotkey."""
        if self._visible_flag:
            self.hide()
        else:
            self.summon()

    def _is_geometry_on_any_screen(self) -> bool:
        try:
            from PySide6.QtGui import QGuiApplication

            for screen in QGuiApplication.screens():
                if screen.availableGeometry().intersects(self.frameGeometry()):
                    return True
        except Exception:
            return True
        return False

    # ---------- visibility tracking ----------

    def showEvent(self, event) -> None:
        self._visible_flag = True
        super().showEvent(event)

    def hideEvent(self, event) -> None:
        self._visible_flag = False
        super().hideEvent(event)

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
