import keyboard
import logging
import pyperclip
import threading
import time
from PIL import ImageGrab

from shouyu.collector.basecollector import BaseCollector
from shouyu.collector.chrome import ChromeCollector
from shouyu.config import Config
from shouyu.service.context import ExcelContext
from shouyu.service.excel import KbExcel
from shouyu.decorator.exceptionhandler import exception_handler
from shouyu.utils.process import ProcessManager
from shouyu.queue import TaskExecutor


class Shortcut:
    executor = TaskExecutor()
    last_copy_time = 0
    last_show_time = 0

    @classmethod
    def start(cls):
        threading.Thread(target=cls.executor.run, daemon=True).start()

    @staticmethod
    @exception_handler
    def save_clipboard():
        img = ImageGrab.grabclipboard()
        copied_text = pyperclip.paste()
        KbExcel().save(img or copied_text)

    @classmethod
    @exception_handler
    def save_clipboard_by_copy_2_times(cls):
        current_time = time.time()
        if current_time - cls.last_copy_time < 3:
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
    @exception_handler
    def one_key_save(cls):
        collector = cls._generate_collector()
        records = collector.collect_records()
        for index, record in enumerate(records):
            if index == 1:
                ExcelContext.column_steps = 1
            KbExcel().save(record, open_excel_again=False)
        ExcelContext.column_steps = -1

    @classmethod
    @exception_handler
    def show_status(cls):
        logging.info('show status')
        KbExcel().move_column()

    @classmethod
    @exception_handler
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
    @exception_handler
    def move_to_left_or_right(cls, steps):
        KbExcel().move_column(steps)

    @classmethod
    @exception_handler
    def reset_column(cls):
        ExcelContext.column_steps = 0
        KbExcel().move_column()

    @classmethod
    @exception_handler
    def move_to_home(cls):
        KbExcel().move_to_home()

    @classmethod
    @exception_handler
    def insert_row_sperator(cls, step=0):
        ExcelContext.row_steps += step
        KbExcel().move_column()

    @classmethod
    @exception_handler
    def switch_one_or_multiple_cell_mode(cls):
        ExcelContext.one_cell_mode = not ExcelContext.one_cell_mode
        KbExcel().move_column()

    @classmethod
    @exception_handler
    def open_excel(cls):
        ProcessManager.open(Config.excel_path())

    @classmethod
    @exception_handler
    def close_excel(cls):
        ProcessManager.terminate_by_path(Config.excel_path())

    @classmethod
    def register_hot_keys(cls):
        # save clipboard to kb.xlsx
        one_key_save_short_key = Config.get_shortcut('one_key_save')
        if one_key_save_short_key:
            keyboard.add_hotkey(one_key_save_short_key, cls.executor.add, args=(cls.one_key_save, ()))

        keyboard.add_hotkey('ctrl+c', cls.executor.add, args=(cls.save_clipboard_by_copy_2_times, ()))
        keyboard.add_hotkey('print screen', cls.executor.add, args=(cls.save_clipboard_by_copy_2_times, ()))
        keyboard.add_hotkey('windows+print screen', cls.executor.add, args=(cls.save_clipboard_by_copy_2_times, ()))
        keyboard.add_hotkey('alt+print screen', cls.executor.add, args=(cls.save_clipboard_by_copy_2_times, ()))

        save_clipboard_short_key = Config.get_shortcut('save_clipboard')
        if save_clipboard_short_key:
            keyboard.add_hotkey(save_clipboard_short_key, cls.executor.add, args=(cls.save_clipboard, ()))
        # Open or close kb.xlsx
        open_excel_short_key = Config.get_shortcut('open_excel')
        if open_excel_short_key:
            keyboard.add_hotkey(open_excel_short_key, cls.open_excel)
        close_excel_short_key = Config.get_shortcut('close_excel')
        if close_excel_short_key:
            keyboard.add_hotkey(close_excel_short_key, cls.close_excel)
        # show or move current column position
        move_to_right_short_key = Config.get_shortcut('move_to_right')
        if move_to_right_short_key:
            keyboard.add_hotkey(move_to_right_short_key, cls.executor.add, args=(cls.move_to_left_or_right, (1,)))
        move_to_left_short_key = Config.get_shortcut('move_to_left')
        if move_to_left_short_key:
            keyboard.add_hotkey(move_to_left_short_key, cls.executor.add, args=(cls.move_to_left_or_right, (-1,)))
        home_short_key = Config.get_shortcut('home')
        if home_short_key:
            keyboard.add_hotkey(home_short_key, cls.executor.add, args=(cls.move_to_home, ()))

        reset_column_short_key = Config.get_shortcut('reset_column')
        if reset_column_short_key:
            keyboard.add_hotkey(reset_column_short_key, cls.executor.add, args=(cls.reset_column, ()))

        show_status_short_key = Config.get_shortcut('show_status')
        if show_status_short_key:
            keyboard.add_hotkey(show_status_short_key, cls.executor.add, args=(cls.show_status, ()))

        insert_row_separator_short_key = Config.get_shortcut('insert_row_separator')
        if insert_row_separator_short_key:
            keyboard.add_hotkey(insert_row_separator_short_key, cls.insert_row_sperator, args=(1,))

        one_or_multiple_cells_mode_short_key = Config.get_shortcut('one_or_multiple_cells_mode')
        if one_or_multiple_cells_mode_short_key:
            keyboard.add_hotkey(one_or_multiple_cells_mode_short_key, cls.switch_one_or_multiple_cell_mode)
        # HACK: keyboard caught windows+l pressed event when user is locking screen,
        # but missing the released event.
        keyboard.add_hotkey('windows+l', cls.clear_pressed_events)
        threading.Thread(target=cls.health_check, daemon=True).start()