import logging
import os
import threading
from functools import partial

import keyboard
import pyperclip
import pywinauto as pywinauto
import time
from PIL import ImageGrab

from shouyu.collector.basecollector import BaseCollector
from shouyu.collector.chrome import ChromeCollector
from shouyu.config import Config
from shouyu.decorator.actionhandler import action_handler
from shouyu.queue import TaskExecutor
from shouyu.service.context import ExcelContext
from shouyu.service.excel import KbExcel
from shouyu.util.process import ProcessManager


class Shortcut:
    executor = TaskExecutor()
    _press_state = {}

    @classmethod
    def start(cls):
        threading.Thread(target=cls.executor.run, daemon=True).start()

    @classmethod
    def _create_multi_press_handler(cls, actions, time_window=1.0, wait_time=0.6):
        """Return a handler that distinguishes single / double / triple presses.

        actions: dict mapping press-count to a (callable, args) tuple.
                 e.g. {2: (fn_2x, ()), 3: (fn_3x, ())}
        Only counts present in *actions* trigger work; other counts are ignored.
        When the max count is reached the action fires immediately; otherwise a
        short timer waits to see whether another press follows.
        """
        state = {'count': 0, 'first_time': 0.0, 'timer': None}
        max_count = max(actions.keys())

        def handler():
            current_time = time.time()
            if current_time - state['first_time'] > time_window:
                state['count'] = 1
                state['first_time'] = current_time
            else:
                state['count'] += 1

            if state.get('timer'):
                state['timer'].cancel()
                state['timer'] = None

            if state['count'] >= max_count:
                count = state['count']
                state['count'] = 0
                state['first_time'] = 0
                fn, args = actions[max_count]
                cls.executor.add(fn, args)
            else:
                count_snapshot = state['count']

                def on_timeout():
                    if state['count'] == count_snapshot:
                        state['count'] = 0
                        state['first_time'] = 0
                        entry = actions.get(count_snapshot)
                        if entry:
                            fn, args = entry
                            cls.executor.add(fn, args)

                state['timer'] = threading.Timer(wait_time, on_timeout)
                state['timer'].start()

        return handler

    @staticmethod
    @action_handler
    def save_clipboard():
        img = ImageGrab.grabclipboard()
        copied_text = pyperclip.paste()
        KbExcel().append(img or copied_text)

    @classmethod
    @action_handler
    def save_clipboard_to_column(cls, column):
        ExcelContext.target_column = column
        img = ImageGrab.grabclipboard()
        copied_text = pyperclip.paste()
        KbExcel().append(img or copied_text)

    @classmethod
    def _generate_collector(cls):
        if BaseCollector.get_process_name() == 'chrome.exe':
            return ChromeCollector()
        else:
            return BaseCollector()

    @classmethod
    @action_handler
    def show_status(cls):
        logging.info('show status')
        KbExcel().move_column()

    @classmethod
    @action_handler
    def clear_pressed_events(cls):
        with keyboard._pressed_events_lock:
            keyboard._pressed_events.clear()

    @classmethod
    def health_check(cls):
        while True:
            with keyboard._pressed_events_lock:
                if cls._is_key_overtime(keyboard._pressed_events):
                    keyboard._pressed_events.clear()
            time.sleep(10)

    @staticmethod
    def _is_key_overtime(pressed_events):
        for event in pressed_events.values():
            from time import time as now
            if now() - event.time > 10:
                return True

    @classmethod
    @action_handler
    def switch_one_or_multiple_cell_mode(cls):
        ExcelContext.cross_multiple_rows = not ExcelContext.cross_multiple_rows
        KbExcel().move_column()

    @classmethod
    @action_handler
    def open_excel(cls):
        ExcelContext.show_pop_up_message = False
        KbExcel().move_column()
        ExcelContext.show_pop_up_message = True
        ProcessManager.open_file(Config.excel_path())

    @classmethod
    def _visible_excel(cls):
        excel_file_name = os.path.basename(Config.excel_path())
        desktop = pywinauto.Desktop(backend="uia")
        windows_filter = partial(desktop.windows, top_level_only=False, visible_only=False)
        windows = windows_filter(title_re=f'{excel_file_name} - WPS Office')
        if not windows:
            windows = windows_filter(title_re=r'.*WPS Office')
        if windows:
            wind = windows[0]
            wind.click_input()

    @classmethod
    @action_handler
    def close_excel(cls):
        ProcessManager.graceful_close_by_path(Config.excel_path())

    @classmethod
    def show_todo_panel(cls):
        from shouyu.view.qt_app import QtApp

        QtApp.show_todo_panel()

    @classmethod
    def show_habit_dialog(cls):
        from shouyu.config import Config
        from shouyu.view.qt_app import QtApp

        QtApp.show_habit_dialog(Config.habits())

    @classmethod
    def toggle_pomodoro(cls):
        from shouyu.service.pomodoro import PomodoroService
        from shouyu.view.pomodoro_window import PomodoroWindow
        from shouyu.view.qt_app import QtApp

        # See Tray.on_toggle_pomodoro: if a cycle is in progress but the
        # floating window is hidden, prefer summoning it back over toggling
        # the timer state. Otherwise users feel punished for pressing the
        # hotkey to "check on" the timer.
        snap = PomodoroService.instance().snapshot()
        running = snap.get("phase") in ("working", "short_break", "long_break", "paused")
        if running and not PomodoroWindow.is_visible_safe():
            QtApp.show_pomodoro_window()
            return
        PomodoroService.instance().toggle()

    @classmethod
    def toggle_pomodoro_window(cls):
        from shouyu.view.qt_app import QtApp

        QtApp.toggle_pomodoro_window()

    @classmethod
    def show_backup_restore(cls):
        from shouyu.view.qt_app import QtApp

        QtApp.show_backup_restore()

    @classmethod
    def _add_hot_key_from_config(cls, key, fun, args=(), is_in_queue=True):
        short_key = Config.get_shortcut(key)
        if short_key:
            if is_in_queue:
                keyboard.add_hotkey(short_key, cls.executor.add, args=(fun, args))
            else:
                keyboard.add_hotkey(short_key, fun, args=args)

    @classmethod
    def register_hot_keys(cls):
        # ctrl+c / print screen: 2x → column B, 3x → column A
        copy_handler = cls._create_multi_press_handler({
            2: (cls.save_clipboard_to_column, ('B',)),
            3: (cls.save_clipboard_to_column, ('A',)),
        })
        keyboard.add_hotkey('ctrl+c', copy_handler)
        keyboard.add_hotkey('print screen', copy_handler)
        keyboard.add_hotkey('windows+print screen', copy_handler)
        keyboard.add_hotkey('alt+print screen', copy_handler)

        # save_clipboard: 1x → column B, 2x → column A
        save_clipboard_handler = cls._create_multi_press_handler({
            1: (cls.save_clipboard_to_column, ('B',)),
            2: (cls.save_clipboard_to_column, ('A',)),
        })
        short_key = Config.get_shortcut('save_clipboard')
        if short_key:
            keyboard.add_hotkey(short_key, save_clipboard_handler)
        cls._add_hot_key_from_config('open_excel', cls.open_excel)
        cls._add_hot_key_from_config('close_excel', cls.close_excel, is_in_queue=False)
        cls._add_hot_key_from_config('show_status', cls.show_status)
        cls._add_hot_key_from_config('one_or_multiple_cells_mode', cls.switch_one_or_multiple_cell_mode)
        cls._add_hot_key_from_config('show_todo', cls.show_todo_panel, is_in_queue=False)
        cls._add_hot_key_from_config('show_habits', cls.show_habit_dialog, is_in_queue=False)
        cls._add_hot_key_from_config('toggle_pomodoro', cls.toggle_pomodoro, is_in_queue=False)
        cls._add_hot_key_from_config('toggle_pomodoro_window', cls.toggle_pomodoro_window, is_in_queue=False)
        cls._add_hot_key_from_config('restore_backup', cls.show_backup_restore, is_in_queue=False)
        # HACK: keyboard caught windows+l pressed event when user is locking screen,
        # but missing the released event.
        keyboard.add_hotkey('windows+l', cls.clear_pressed_events)
        threading.Thread(target=cls.health_check, daemon=True).start()
