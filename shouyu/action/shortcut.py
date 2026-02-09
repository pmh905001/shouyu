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
    last_copy_time = 0
    last_show_time = 0

    @classmethod
    def start(cls):
        threading.Thread(target=cls.executor.run, daemon=True).start()

    @staticmethod
    @action_handler
    def save_clipboard():
        img = ImageGrab.grabclipboard()
        copied_text = pyperclip.paste()
        KbExcel().append(img or copied_text)

    @classmethod
    @action_handler
    def save_clipboard_by_copy_2_times(cls):
        current_time = time.time()
        if current_time - cls.last_copy_time < 1:
            cls.save_clipboard()
        else:
            cls.last_copy_time = current_time

    @classmethod
    def _generate_collector(cls):
        if BaseCollector.get_process_name() == 'chrome.exe':
            return ChromeCollector()
        else:
            return BaseCollector()

    @classmethod
    @action_handler
    def one_key_save(cls):
        collector = cls._generate_collector()
        records = collector.collect_records()
        KbExcel().append(records)

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
    def move_to_left_or_right(cls, steps):
        KbExcel().move_column(steps)

    @classmethod
    @action_handler
    def reset_column(cls):
        ExcelContext.column_steps = 0
        ExcelContext.row_steps = 0
        KbExcel().move_column()

    @classmethod
    @action_handler
    def move_to_home(cls):
        KbExcel().move_to_home()

    @classmethod
    @action_handler
    def move_row_separator(cls, step=0):
        if ExcelContext.row_steps + step < 0:
            ExcelContext.row_steps = 0
        else:
            ExcelContext.row_steps += step
        KbExcel().move_column()

    @classmethod
    @action_handler
    def switch_one_or_multiple_cell_mode(cls):
        ExcelContext.cross_multiple_rows = not ExcelContext.cross_multiple_rows
        KbExcel().move_column()

    @classmethod
    @action_handler
    def open_excel(cls):
        KbExcel().move_column()
        ProcessManager.open_file(Config.excel_path())
        # cls._visible_excel()

    @classmethod
    def _visible_excel(cls):
        excel_file_name = os.path.basename(Config.excel_path())
        desktop = pywinauto.Desktop(backend="uia")
        # TODO: support MS excel
        windows_filter = partial(desktop.windows, top_level_only=False, visible_only=False)
        windows = windows_filter(title_re=f'{excel_file_name} - WPS Office')
        if not windows:
            windows = windows_filter(title_re=r'.*WPS Office')
        if windows:
            wind = windows[0]
            # if opened 2 tab, only click task bar, not select which tab to be visible
            # TODO: need to select which tab is visible if exists 2 or more tabs.
            wind.click_input()

    @classmethod
    @action_handler
    def close_excel(cls):
        ProcessManager.terminate_by_path(Config.excel_path())

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
        # save clipboard to kb.xlsx
        cls._add_hot_key_from_config('one_key_save', cls.one_key_save)
        keyboard.add_hotkey('ctrl+c', cls.executor.add, args=(cls.save_clipboard_by_copy_2_times, ()))
        keyboard.add_hotkey('print screen', cls.executor.add, args=(cls.save_clipboard_by_copy_2_times, ()))
        keyboard.add_hotkey('windows+print screen', cls.executor.add, args=(cls.save_clipboard_by_copy_2_times, ()))
        keyboard.add_hotkey('alt+print screen', cls.executor.add, args=(cls.save_clipboard_by_copy_2_times, ()))
        cls._add_hot_key_from_config('save_clipboard', cls.save_clipboard)
        # Open or close kb.xlsx
        cls._add_hot_key_from_config('open_excel', cls.open_excel)
        cls._add_hot_key_from_config('close_excel', cls.close_excel, is_in_queue=False)
        # show or move current column position
        cls._add_hot_key_from_config('move_to_right', cls.move_to_left_or_right, (1,))
        cls._add_hot_key_from_config('move_to_left', cls.move_to_left_or_right, (-1,))
        cls._add_hot_key_from_config('home', cls.move_to_home)
        cls._add_hot_key_from_config('reset_column', cls.reset_column)
        cls._add_hot_key_from_config('show_status', cls.show_status)
        cls._add_hot_key_from_config('insert_row_separator', cls.move_row_separator, (1,))
        cls._add_hot_key_from_config('remove_row_separator', cls.move_row_separator, (-1,))
        cls._add_hot_key_from_config('one_or_multiple_cells_mode', cls.switch_one_or_multiple_cell_mode)
        # HACK: keyboard caught windows+l pressed event when user is locking screen,
        # but missing the released event.
        keyboard.add_hotkey('windows+l', cls.clear_pressed_events)
        threading.Thread(target=cls.health_check, daemon=True).start()
