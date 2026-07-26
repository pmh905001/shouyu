"""Floating mini-window that shows the current pomodoro phase / countdown.

The window stays on top, is frameless, and can be dragged. It receives
events from PomodoroService through QtApp.emit_pomodoro_event(...).
"""
from __future__ import annotations

import random
from typing import Optional

from PySide6.QtCore import Qt, QPoint, QTimer
from PySide6.QtGui import QColor, QMouseEvent, QPainter
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
    "planning": ("计划中", ACCENT_COLOR_HEX),
    "working": ("专注中", IN_PROGRESS_COLOR_HEX),
    "short_break": ("短休息", DONE_COLOR_HEX),
    "long_break": ("长休息", DONE_COLOR_HEX),
    "lunch_break": ("午休", DONE_COLOR_HEX),
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
        # True only during the level-2 hard alarm (faster blink + ack button
        # + window forced to front). Level-1 soft blink leaves this False.
        self._alarm_active = False
        # Full-screen silent popup used for the office/quiet profile alarm.
        self._overlay: Optional["IdleOverlay"] = None
        # Centered "该休息了" card shown at the start of a break (see
        # _maybe_show_break_reminder). Lazily created.
        self._break_reminder: Optional["BreakReminder"] = None
        # Cached so we can restore exactly what we had after blink ends —
        # the phase may switch while idle (e.g. work → break would also
        # clear the warning), and we want to reflect whatever the latest
        # phase event told us, not a stale label.
        self._normal_phase_text = "空闲"
        self._normal_phase_color = SUBTEXT_COLOR_HEX

        self._build_ui()
        self._refresh_mode_button()
        self._refresh_env_button()
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

        self.env_btn = QPushButton("🔄 自动")
        self.env_btn.setToolTip("切换 自动 / 在家 / 在单位")
        self.env_btn.setCursor(Qt.PointingHandCursor)
        self.env_btn.setFixedHeight(20)
        self.env_btn.setStyleSheet(
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
        self.env_btn.clicked.connect(self._on_env_btn_clicked)
        header_row.addWidget(self.env_btn)
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

        # Shown only during a hard idle alarm. It is the ONLY way to silence
        # the alarm — deliberately prominent (red) so it can't be ignored.
        self.ack_btn = QPushButton("✋ 我回来了")
        self.ack_btn.setToolTip("停止警报并记一次走神")
        self.ack_btn.setCursor(Qt.PointingHandCursor)
        self.ack_btn.setStyleSheet(self._ack_button_style())
        self.ack_btn.clicked.connect(self._on_ack_clicked)
        self.ack_btn.setVisible(False)
        button_row.addWidget(self.ack_btn)

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

        # Shown only during a focus phase: finish this pomodoro ahead of time
        # (task done early) and go straight to the break.
        self.finish_btn = QPushButton("去休息")
        self.finish_btn.setToolTip("任务提前做完了？结束这个番茄，直接进入休息（计一个 🍅）")
        self.finish_btn.setStyleSheet(self._button_style())
        self.finish_btn.clicked.connect(self._on_finish_early_clicked)
        self.finish_btn.setVisible(False)
        button_row.addWidget(self.finish_btn)

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

    @staticmethod
    def _ack_button_style() -> str:
        return (
            "QPushButton {"
            "  background-color: #FF3B30;"
            "  color: white;"
            "  border: none;"
            "  border-radius: 5px;"
            "  padding: 3px 8px;"
            "  font-size: 11px;"
            "  font-weight: 700;"
            "}"
            "QPushButton:hover { background-color: #FF5A50; }"
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
            # A break just began — announce it prominently so the user doesn't
            # keep working through it. Any other phase hides a lingering card.
            if phase in ("short_break", "long_break"):
                self._maybe_show_break_reminder(phase, duration)
            else:
                self._hide_break_reminder()
        elif event == "tick":
            try:
                remaining = int(payload)
            except ValueError:
                remaining = 0
            self._set_remaining(remaining)
            if self._break_reminder is not None and self._break_reminder.isVisible():
                self._break_reminder.update_remaining(remaining)
        elif event == "extended":
            # phase unchanged; remaining seconds will refresh on next tick
            pass
        elif event == "mode_changed":
            self._refresh_mode_button()
        elif event == "env_mode_changed":
            self._refresh_env_button()
        elif event == "paused":
            self._stop_blink()
            self._hide_break_reminder()
            self._set_phase("paused")
            self.toggle_btn.setText("继续")
        elif event == "resumed":
            snap = PomodoroService.instance().snapshot()
            self._set_phase(snap["phase"])
            self._update_tip(snap["phase"])
            self.toggle_btn.setText("暂停")
        elif event == "stopped":
            self._stop_blink()
            self._hide_break_reminder()
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
        elif event == "idle_alarm_on":
            try:
                idle_seconds = int(payload)
            except ValueError:
                idle_seconds = 0
            self._start_alarm(idle_seconds)
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
        self._refresh_env_button()

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
        if phase == "planning":
            self.tip_label.setText("📝 写下今日最重要的 3 件事，其余往后排")
            self.tip_label.setVisible(True)
        elif phase == "lunch_break":
            self.tip_label.setText("🍚 午餐时间，好好吃饭休息，别工作")
            self.tip_label.setVisible(True)
        elif phase in ("short_break", "long_break"):
            self.tip_label.setText(random.choice(_BREAK_TIPS))
            self.tip_label.setVisible(True)
        elif phase == "working":
            self.tip_label.setText(random.choice(_WORK_TIPS))
            self.tip_label.setVisible(True)
        else:
            self.tip_label.clear()
            self.tip_label.setVisible(False)

    def _refresh_action_visibility(self, phase: str) -> None:
        if phase in ("working", "planning"):
            self.extend_btn.setVisible(True)
            self.extend_btn.setText("+5m")
            self.skip_break_btn.setVisible(False)
            self.finish_btn.setVisible(True)
        elif phase in ("short_break", "long_break"):
            self.extend_btn.setVisible(True)
            self.extend_btn.setText("休息+2m")
            self.skip_break_btn.setVisible(True)
            self.finish_btn.setVisible(False)
        else:
            # idle / paused / lunch_break: no extend, and lunch can't be skipped
            self.extend_btn.setVisible(False)
            self.skip_break_btn.setVisible(False)
            self.finish_btn.setVisible(False)

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

    def _on_finish_early_clicked(self) -> None:
        from shouyu.service.pomodoro import PomodoroService

        PomodoroService.instance().finish_early()

    def _on_ack_clicked(self) -> None:
        from shouyu.service.pomodoro import PomodoroService

        PomodoroService.instance().acknowledge_idle()

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

    def _on_env_btn_clicked(self) -> None:
        from shouyu.service.pomodoro import PomodoroService

        svc = PomodoroService.instance()
        order = [
            PomodoroService.ENV_MODE_AUTO,
            PomodoroService.ENV_MODE_HOME,
            PomodoroService.ENV_MODE_OFFICE,
        ]
        cur = svc.env_mode()
        nxt = order[(order.index(cur) + 1) % len(order)] if cur in order else order[0]
        svc.set_env_mode(nxt)
        self._refresh_env_button()

    def _refresh_env_button(self) -> None:
        from shouyu.service.pomodoro import PomodoroService

        svc = PomodoroService.instance()
        setting = svc.env_mode()
        resolved = svc.resolved_env_mode()
        if setting == PomodoroService.ENV_MODE_HOME:
            self.env_btn.setText("🏠 在家")
            self.env_btn.setToolTip("在家：柔和提示音 — 点击切到 在单位")
        elif setting == PomodoroService.ENV_MODE_OFFICE:
            self.env_btn.setText("🏢 单位")
            self.env_btn.setToolTip("在单位：静音 + 全屏弹窗 — 点击切回 自动")
        else:
            is_office = resolved == PomodoroService.ENV_MODE_OFFICE
            self.env_btn.setText("🔄 自动·" + ("司" if is_office else "家"))
            self.env_btn.setToolTip(
                "自动：按时间表决定，当前 = "
                + ("在单位(静音弹窗)" if is_office else "在家(有声)")
                + " — 点击切到 在家"
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
        self._alarm_active = False
        self._blink_idle_seconds = idle_seconds
        minutes = max(1, idle_seconds // 60)
        self.phase_label.setText(f"⚠ 已静止 {minutes}m")
        self.phase_label.setStyleSheet(
            f"color: {self._ALERT_RED}; font-size: 12px; font-weight: 700;"
        )
        self._apply_card_alert_style(True)
        self._blink_timer.setInterval(500)
        self._blink_on = True
        self._on_blink_tick()  # flip immediately, don't wait 500ms
        self._blink_timer.start()

    def _start_alarm(self, idle_seconds: int) -> None:
        """Enter the hard, un-ignorable alarm mode (level 2).

        Unlike the soft blink, this:
          * forces the window back on-screen and to the front even if the
            user had hidden it (respecting the hide gesture is no longer more
            important than getting you back to work);
          * blinks faster (250ms) and shows a louder "🚨 快回来工作" label;
          * surfaces the red "✋ 我回来了" button, which is the only way to
            dismiss it (moving the mouse won't).
        """
        self._idle_warning_active = True
        self._alarm_active = True
        self._blink_idle_seconds = idle_seconds
        minutes = max(1, idle_seconds // 60)
        self.phase_label.setText(f"🚨 快回来工作 {minutes}m")
        self.phase_label.setStyleSheet(
            f"color: {self._ALERT_RED}; font-size: 12px; font-weight: 800;"
        )
        self.ack_btn.setVisible(True)
        self._apply_card_alert_style(True)
        self._blink_timer.setInterval(250)
        self._blink_on = True
        self._on_blink_tick()
        self._blink_timer.start()
        # Force the window back so it can't be ignored from behind other apps.
        self.summon()

        # Office/quiet profile: no sound is played (gated in the service), so
        # take over the whole screen instead — a silent but un-ignorable
        # popup that you must dismiss with "我回来了".
        from shouyu.service.pomodoro import PomodoroService

        if PomodoroService.instance().resolved_env_mode() == PomodoroService.ENV_MODE_OFFICE:
            self._show_overlay(idle_seconds)

    def _stop_blink(self) -> None:
        if not self._idle_warning_active and not self._blink_timer.isActive():
            return
        self._blink_timer.stop()
        self._blink_timer.setInterval(500)
        self._idle_warning_active = False
        self._alarm_active = False
        self.ack_btn.setVisible(False)
        self._hide_overlay()
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

    def _show_overlay(self, idle_seconds: int) -> None:
        if self._overlay is None:
            self._overlay = IdleOverlay(on_ack=self._on_ack_clicked)
        target_screen = self.screen()
        self._overlay.show_over(idle_seconds, target_screen)

    def _hide_overlay(self) -> None:
        if self._overlay is not None:
            self._overlay.hide()

    # ---------- break reminder ----------

    def _maybe_show_break_reminder(self, phase: str, remaining: int) -> None:
        """Pop the centered break card (unless disabled) and bring the timer to
        the front, so a starting break can't be quietly worked through."""
        from shouyu.config import Config

        try:
            style = Config.pomodoro_break_reminder()
        except Exception:
            style = 'center'
        if style == 'off':
            return
        if self._break_reminder is None:
            self._break_reminder = BreakReminder(
                on_start=self._on_break_start_clicked,
                on_extend=self._on_extend_clicked,
                on_skip=self._on_skip_break_clicked,
            )
        tip = random.choice(_BREAK_TIPS)
        # Bring the floating timer forward first, then show the card on top so
        # the card ends up frontmost/active.
        self.summon()
        self._break_reminder.show_reminder(phase, remaining, tip, self.screen())

    def _hide_break_reminder(self) -> None:
        if self._break_reminder is not None:
            self._break_reminder.hide()

    def _on_break_start_clicked(self) -> None:
        # "开始休息" just dismisses the card; the break is already running.
        self._hide_break_reminder()

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
            drifts = AppState.get_today_counter('focus_drifts')
            if drifts:
                text += f"  · 走神 {drifts}"
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


class IdleOverlay(QWidget):
    """Full-screen, silent, un-ignorable popup for the office/quiet profile.

    When the hard idle alarm fires and sound is suppressed (office mode), we
    can't rely on a beep, so we cover the whole screen with a translucent red
    layer that physically blocks whatever you were staring at. The only way
    out is the big "我回来了" button (or clicking anywhere), which routes to
    the same acknowledge path as the floating window.
    """

    def __init__(self, on_ack) -> None:
        super().__init__()
        self._on_ack = on_ack
        self.setWindowFlag(Qt.FramelessWindowHint, True)
        self.setWindowFlag(Qt.WindowStaysOnTopHint, True)
        self.setWindowFlag(Qt.Tool, True)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(18)

        title = QLabel("🚨 快回来工作")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: #FFFFFF; font-size: 44px; font-weight: 800;")
        layout.addWidget(title)

        self.sub_label = QLabel("")
        self.sub_label.setAlignment(Qt.AlignCenter)
        self.sub_label.setStyleSheet("color: #FFD5D2; font-size: 18px;")
        layout.addWidget(self.sub_label)

        self.ack_btn = QPushButton("✋ 我回来了")
        self.ack_btn.setCursor(Qt.PointingHandCursor)
        self.ack_btn.setFixedSize(220, 56)
        self.ack_btn.setStyleSheet(
            "QPushButton {"
            "  background-color: #FF3B30;"
            "  color: white;"
            "  border: none;"
            "  border-radius: 12px;"
            "  font-size: 20px;"
            "  font-weight: 800;"
            "}"
            "QPushButton:hover { background-color: #FF5A50; }"
        )
        self.ack_btn.clicked.connect(self._ack)
        layout.addWidget(self.ack_btn, alignment=Qt.AlignCenter)

    def _ack(self) -> None:
        if callable(self._on_ack):
            self._on_ack()

    def show_over(self, idle_seconds: int, screen) -> None:
        minutes = max(1, int(idle_seconds) // 60)
        self.sub_label.setText(f"已静止 {minutes} 分钟 · 点“我回来了”继续专注")
        if screen is not None:
            self.setGeometry(screen.geometry())
        self.showFullScreen()
        self.raise_()
        self.activateWindow()

    def paintEvent(self, event) -> None:
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor(40, 0, 0, 205))

    def mousePressEvent(self, event: QMouseEvent) -> None:
        # Clicking anywhere on the overlay also acknowledges — make it as
        # easy as possible to dismiss once you've actually come back.
        self._ack()
        super().mousePressEvent(event)


class BreakReminder(QWidget):
    """Centered, hard-to-miss (but calm) card announcing that a break started.

    The floating timer alone is too easy to work straight through — it just
    sits in the corner. This pops a centered green card so the break actually
    registers. Unlike the red IdleOverlay it does NOT cover the whole screen
    and is dismissed in one click (开始休息 / 跳过休息), so it interrupts
    without being enraging.
    """

    def __init__(self, on_start, on_extend, on_skip) -> None:
        super().__init__()
        self._on_start = on_start
        self._on_extend = on_extend
        self._on_skip = on_skip
        self.setWindowFlag(Qt.FramelessWindowHint, True)
        self.setWindowFlag(Qt.WindowStaysOnTopHint, True)
        self.setWindowFlag(Qt.Tool, True)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setFixedSize(460, 250)
        self._drag_offset: Optional[QPoint] = None
        self._build_ui()

    def _build_ui(self) -> None:
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)

        card = QFrame()
        card.setStyleSheet(
            f"""
            QFrame {{
                background-color: {PANEL_COLOR_HEX};
                border-radius: 16px;
                border: 2px solid {DONE_COLOR_HEX};
            }}
            """
        )
        outer.addWidget(card)

        layout = QVBoxLayout(card)
        layout.setContentsMargins(28, 22, 28, 22)
        layout.setSpacing(10)
        layout.setAlignment(Qt.AlignCenter)

        self.title_label = QLabel("☕ 该休息了")
        self.title_label.setAlignment(Qt.AlignCenter)
        self.title_label.setStyleSheet(
            f"color: {DONE_COLOR_HEX}; font-size: 26px; font-weight: 800; border: none;"
        )
        layout.addWidget(self.title_label)

        self.time_label = QLabel("00:00")
        self.time_label.setAlignment(Qt.AlignCenter)
        self.time_label.setStyleSheet(
            f"color: {TEXT_COLOR_HEX}; font-size: 40px; font-weight: 700; border: none;"
        )
        layout.addWidget(self.time_label)

        self.tip_label = QLabel("")
        self.tip_label.setAlignment(Qt.AlignCenter)
        self.tip_label.setWordWrap(True)
        self.tip_label.setStyleSheet(
            f"color: {SUBTEXT_COLOR_HEX}; font-size: 13px; border: none;"
        )
        layout.addWidget(self.tip_label)

        button_row = QHBoxLayout()
        button_row.setSpacing(10)
        button_row.setAlignment(Qt.AlignCenter)

        start_btn = QPushButton("开始休息")
        start_btn.setCursor(Qt.PointingHandCursor)
        start_btn.setFixedHeight(34)
        start_btn.setStyleSheet(
            "QPushButton {"
            f"  background-color: {DONE_COLOR_HEX};"
            "  color: white;"
            "  border: none;"
            "  border-radius: 8px;"
            "  padding: 4px 18px;"
            "  font-size: 14px;"
            "  font-weight: 700;"
            "}"
            "QPushButton:hover { background-color: #2F9E44; }"
        )
        start_btn.clicked.connect(self._start)
        button_row.addWidget(start_btn)

        extend_btn = QPushButton("休息 +2m")
        extend_btn.setCursor(Qt.PointingHandCursor)
        extend_btn.setFixedHeight(34)
        extend_btn.setStyleSheet(self._neutral_button_style())
        extend_btn.clicked.connect(self._extend)
        button_row.addWidget(extend_btn)

        skip_btn = QPushButton("跳过休息")
        skip_btn.setToolTip("跳过休息，直接回到专注（会被记录，不推荐）")
        skip_btn.setCursor(Qt.PointingHandCursor)
        skip_btn.setFixedHeight(34)
        skip_btn.setStyleSheet(self._neutral_button_style())
        skip_btn.clicked.connect(self._skip)
        button_row.addWidget(skip_btn)

        layout.addLayout(button_row)

    @staticmethod
    def _neutral_button_style() -> str:
        return (
            "QPushButton {"
            "  background-color: rgba(255,255,255,0.08);"
            "  color: white;"
            "  border: none;"
            "  border-radius: 8px;"
            "  padding: 4px 16px;"
            "  font-size: 13px;"
            "}"
            "QPushButton:hover { background-color: rgba(255,255,255,0.18); }"
        )

    def _start(self) -> None:
        # "开始休息" simply dismisses the card — the break is already running;
        # the floating timer keeps the countdown visible.
        self.hide()
        if callable(self._on_start):
            self._on_start()

    def _extend(self) -> None:
        if callable(self._on_extend):
            self._on_extend()

    def _skip(self) -> None:
        self.hide()
        if callable(self._on_skip):
            self._on_skip()

    def update_remaining(self, seconds: int) -> None:
        self.time_label.setText(_format_remaining(seconds))

    def show_reminder(self, phase: str, remaining: int, tip: str, screen) -> None:
        label = _PHASE_LABEL.get(phase, ("休息", DONE_COLOR_HEX))[0]
        self.title_label.setText(f"☕ 该{label}了，起来歇会儿")
        self.update_remaining(remaining)
        self.tip_label.setText(tip)
        if screen is None:
            from PySide6.QtGui import QGuiApplication

            screen = QGuiApplication.primaryScreen()
        if screen is not None:
            geo = screen.availableGeometry()
            self.move(
                geo.center().x() - self.width() // 2,
                geo.center().y() - self.height() // 2,
            )
        self.show()
        self.raise_()
        self.activateWindow()

    # Let the user nudge the card aside without dismissing it.
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
