"""QApplication lifecycle manager.

Qt requires GUI work to happen on the same thread that owns the QApplication.
This module hosts the QApplication on a dedicated daemon thread and exposes
thread-safe entry points (signals) so the keyboard-hotkey thread, the http
server thread, and the main thread can all request UI work without blocking.
"""
from __future__ import annotations

import logging
import sys
import threading
from typing import List, Optional

from PySide6.QtCore import QObject, Qt, Signal, Slot
from PySide6.QtWidgets import QApplication


class _QtBridge(QObject):
    show_todo_signal = Signal()
    show_habit_signal = Signal(list)
    pomodoro_event_signal = Signal(str, str)
    quit_signal = Signal()

    @Slot()
    def _on_show_todo(self) -> None:
        try:
            from shouyu.view.todo_panel import TodoPanel

            panel = TodoPanel.get_or_create()
            panel.refresh_from_excel()
            panel.show_centered()
        except Exception:
            logging.exception("failed to show todo panel")

    @Slot(list)
    def _on_show_habit(self, habits: list) -> None:
        try:
            from shouyu.view.habit_dialog import HabitDialog

            dialog = HabitDialog.get_or_create()
            dialog.set_habits(list(habits))
            dialog.refresh_plan_from_excel()
            dialog.show_fullscreen()
        except Exception:
            logging.exception("failed to show habit dialog")

    @Slot(str, str)
    def _on_pomodoro_event(self, event: str, payload: str) -> None:
        try:
            from shouyu.view.pomodoro_window import PomodoroWindow

            window = PomodoroWindow.get_or_create()
            window.handle_event(event, payload)
        except Exception:
            logging.exception("failed to dispatch pomodoro event")

    @Slot()
    def _on_quit(self) -> None:
        app = QApplication.instance()
        if app is not None:
            app.quit()


class QtApp:
    _started = False
    _ready = threading.Event()
    _bridge: Optional[_QtBridge] = None
    _app: Optional[QApplication] = None

    @classmethod
    def start(cls) -> None:
        if cls._started:
            return
        cls._started = True
        threading.Thread(target=cls._main, name="qt-app", daemon=True).start()
        cls._ready.wait(timeout=10)

    @classmethod
    def _main(cls) -> None:
        try:
            cls._app = QApplication(sys.argv)
            cls._app.setQuitOnLastWindowClosed(False)
            cls._apply_global_style(cls._app)
            cls._bridge = _QtBridge()
            cls._bridge.show_todo_signal.connect(cls._bridge._on_show_todo, Qt.QueuedConnection)
            cls._bridge.show_habit_signal.connect(cls._bridge._on_show_habit, Qt.QueuedConnection)
            cls._bridge.pomodoro_event_signal.connect(cls._bridge._on_pomodoro_event, Qt.QueuedConnection)
            cls._bridge.quit_signal.connect(cls._bridge._on_quit, Qt.QueuedConnection)
        finally:
            cls._ready.set()
        cls._app.exec()

    @classmethod
    def _apply_global_style(cls, app: QApplication) -> None:
        try:
            from shouyu.view.styles import GLOBAL_QSS

            app.setStyleSheet(GLOBAL_QSS)
        except Exception:
            logging.exception("failed to apply global stylesheet")

    @classmethod
    def show_todo_panel(cls) -> None:
        if cls._bridge is None:
            logging.warning("Qt bridge not ready; cannot show todo panel")
            return
        cls._bridge.show_todo_signal.emit()

    @classmethod
    def show_habit_dialog(cls, habits: List[str]) -> None:
        if cls._bridge is None:
            logging.warning("Qt bridge not ready; cannot show habit dialog")
            return
        cls._bridge.show_habit_signal.emit(list(habits))

    @classmethod
    def emit_pomodoro_event(cls, event: str, payload: str = "") -> None:
        if cls._bridge is None:
            return
        cls._bridge.pomodoro_event_signal.emit(event, payload)

    @classmethod
    def quit(cls) -> None:
        if cls._bridge is None:
            return
        cls._bridge.quit_signal.emit()
