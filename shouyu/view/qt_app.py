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
from PySide6.QtWidgets import QApplication, QMessageBox


class _QtBridge(QObject):
    show_todo_signal = Signal()
    show_habit_signal = Signal(list)
    pomodoro_event_signal = Signal(str, str)
    summon_pomodoro_signal = Signal()
    toggle_pomodoro_window_signal = Signal()
    show_backup_signal = Signal(str)  # payload = path of auto-recovered backup, or ""
    save_status_signal = Signal(str, str, str)  # (level, title, message)
    ocr_capture_signal = Signal(str)  # payload = target column, or ""
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
    def _on_summon_pomodoro(self) -> None:
        try:
            from shouyu.view.pomodoro_window import PomodoroWindow

            window = PomodoroWindow.get_or_create()
            window.summon()
        except Exception:
            logging.exception("failed to summon pomodoro window")

    @Slot()
    def _on_toggle_pomodoro_window(self) -> None:
        try:
            from shouyu.view.pomodoro_window import PomodoroWindow

            window = PomodoroWindow.get_or_create()
            window.toggle_visibility()
        except Exception:
            logging.exception("failed to toggle pomodoro window visibility")

    @Slot(str)
    def _on_show_backup(self, recovered_from: str) -> None:
        try:
            from shouyu.view.backup_dialog import BackupRestoreDialog

            dialog = BackupRestoreDialog.get_or_create()
            dialog.show_centered(recovered_from=recovered_from or None)
        except Exception:
            logging.exception("failed to show backup restore dialog")

    @Slot(str, str, str)
    def _on_save_status(self, level: str, title: str, message: str) -> None:
        try:
            box = QMessageBox()
            if level == 'error':
                box.setIcon(QMessageBox.Critical)
            elif level == 'warning':
                box.setIcon(QMessageBox.Warning)
            else:
                box.setIcon(QMessageBox.Information)
            box.setWindowTitle(title or "授渔")
            box.setText(message)
            box.setStandardButtons(QMessageBox.Ok)
            # Force above the always-on-top habit dialog.
            box.setWindowFlag(Qt.WindowStaysOnTopHint, True)
            box.exec()
        except Exception:
            logging.exception("failed to display save status message")

    @Slot(str)
    def _on_ocr_capture(self, column: str) -> None:
        """Show the fullscreen region-selector (blocking, must run on this
        thread), then hand the cropped image off to a plain background
        thread for OCR + clipboard + enqueue - so RapidOCR inference never
        blocks the Qt event loop. See docs/screenshot-ocr-design.md."""
        try:
            from shouyu.view.region_selector import RegionSelector

            image = RegionSelector.capture()
            if image is None:
                return
            import threading

            from shouyu.action.shortcut import Shortcut

            threading.Thread(
                target=Shortcut.finish_ocr_capture,
                args=(image, column or None),
                name="shouyu-ocr",
                daemon=True,
            ).start()
        except Exception:
            logging.exception("failed to run OCR capture")

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
            cls._bridge.summon_pomodoro_signal.connect(cls._bridge._on_summon_pomodoro, Qt.QueuedConnection)
            cls._bridge.toggle_pomodoro_window_signal.connect(
                cls._bridge._on_toggle_pomodoro_window, Qt.QueuedConnection
            )
            cls._bridge.show_backup_signal.connect(cls._bridge._on_show_backup, Qt.QueuedConnection)
            cls._bridge.save_status_signal.connect(cls._bridge._on_save_status, Qt.QueuedConnection)
            cls._bridge.ocr_capture_signal.connect(cls._bridge._on_ocr_capture, Qt.QueuedConnection)
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
    def show_pomodoro_window(cls) -> None:
        """Bring the floating pomodoro window back to the foreground.

        Safe to call from any thread (tray, hotkey, http server). Creates
        the window if it doesn't exist yet. Always shows; never hides.
        """
        if cls._bridge is None:
            logging.warning("Qt bridge not ready; cannot summon pomodoro window")
            return
        cls._bridge.summon_pomodoro_signal.emit()

    @classmethod
    def toggle_pomodoro_window(cls) -> None:
        """Show the pomodoro window if hidden, hide it if visible.

        Used by the show/hide hotkey. Safe from any thread.
        """
        if cls._bridge is None:
            logging.warning("Qt bridge not ready; cannot toggle pomodoro window")
            return
        cls._bridge.toggle_pomodoro_window_signal.emit()

    @classmethod
    def show_backup_restore(cls, recovered_from: str = "") -> None:
        if cls._bridge is None:
            logging.warning("Qt bridge not ready; cannot show backup restore dialog")
            return
        cls._bridge.show_backup_signal.emit(recovered_from or "")

    @classmethod
    def show_save_status(cls, level: str, title: str, message: str) -> None:
        """Thread-safe entry to show a save-result QMessageBox in the Qt thread."""
        if cls._bridge is None:
            logging.warning("Qt bridge not ready; cannot show save status")
            return
        cls._bridge.save_status_signal.emit(level or 'info', title or "", message or "")

    @classmethod
    def request_ocr_capture(cls, column: Optional[str] = None) -> None:
        """Thread-safe entry to start an OCR region-capture. Safe from any
        thread (typically the hotkey executor thread)."""
        if cls._bridge is None:
            logging.warning("Qt bridge not ready; cannot start OCR capture")
            return
        cls._bridge.ocr_capture_signal.emit(column or "")

    @classmethod
    def quit(cls) -> None:
        if cls._bridge is None:
            return
        cls._bridge.quit_signal.emit()
