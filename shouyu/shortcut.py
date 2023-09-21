import keyboard
import logging
import pyperclip
import threading
import time
from PIL import ImageGrab

from shouyu.collector.basic_collector import BasicCollector
from shouyu.collector.chrome import ChromeCollector
from shouyu.config import ConfigManager
from shouyu.context import ExcelContext
from shouyu.excel import KbExcel
from shouyu.exhandler import exception_handler
from shouyu.process import ProcessManager
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
        if BasicCollector.get_process_name() == 'chrome.exe':
            return ChromeCollector()
        else:
            return BasicCollector()

    @classmethod
    @exception_handler
    def one_key_save(cls):
        collector = cls._generate_collector()
        records = collector.collect_records()
        for index, record in enumerate(records):
            if index == 1:
                ExcelContext.steps = 1
            KbExcel().save(record, open_excel_again=False)
        ExcelContext.steps = -1

    @classmethod
    @exception_handler
    def show_column(cls):
        logging.info("show column")
        current_time = time.time()
        if current_time - cls.last_show_time < 3:
            KbExcel().move_column()
        else:
            cls.last_show_time = current_time

    @classmethod
    @exception_handler
    def clear_pressed_events(cls):
        with keyboard._pressed_events_lock:
            keyboard._pressed_events.clear()

    @classmethod
    def register_hot_keys(cls):
        # save clipboard to kb.xlsx
        one_key_save_short_key = ConfigManager.shortcut('one_key_save')
        if one_key_save_short_key:
            keyboard.add_hotkey(one_key_save_short_key, cls.executor.add, args=(cls.one_key_save, ()))

        copy_2_times_short_key = ConfigManager.shortcut('save_clipboard_by_copy')
        if copy_2_times_short_key:
            keyboard.add_hotkey(copy_2_times_short_key, cls.executor.add, args=(cls.save_clipboard_by_copy_2_times, ()))
        save_clipboard_short_key = ConfigManager.shortcut('save_clipboard')
        if save_clipboard_short_key:
            keyboard.add_hotkey(save_clipboard_short_key, cls.executor.add, args=(cls.save_clipboard, ()))
        print_screen_short_key = ConfigManager.shortcut('print_screen')
        if print_screen_short_key:
            keyboard.add_hotkey(print_screen_short_key, cls.executor.add, args=(cls.save_clipboard_by_copy_2_times, ()))
        windows_print_screen_short_key = ConfigManager.shortcut('windows_print_screen')
        if windows_print_screen_short_key:
            keyboard.add_hotkey(
                windows_print_screen_short_key,
                cls.executor.add,
                args=(cls.save_clipboard_by_copy_2_times, ())
            )
        alt_print_screen_short_key = ConfigManager.shortcut('alt_print_screen')
        if alt_print_screen_short_key:
            keyboard.add_hotkey(
                alt_print_screen_short_key,
                cls.executor.add,
                args=(cls.save_clipboard_by_copy_2_times, ())
            )
        # Open or close kb.xlsx
        open_excel_short_key = ConfigManager.shortcut('open_excel')
        if open_excel_short_key:
            keyboard.add_hotkey(open_excel_short_key, ProcessManager.open, args=(ConfigManager.excel_path(),))
        close_excel_short_key = ConfigManager.shortcut('close_excel')
        if close_excel_short_key:
            keyboard.add_hotkey(close_excel_short_key, ProcessManager.terminate_by_path,
                                args=(ConfigManager.excel_path(),))
        # show or move current column position
        move_to_right_short_key = ConfigManager.shortcut('move_to_right')
        if move_to_right_short_key:
            keyboard.add_hotkey(move_to_right_short_key, cls.executor.add,
                                args=(lambda x: KbExcel().move_column(x), (1,)))
        move_to_left_short_key = ConfigManager.shortcut('move_to_left')
        if move_to_left_short_key:
            keyboard.add_hotkey(move_to_left_short_key, cls.executor.add,
                                args=(lambda x: KbExcel().move_column(x), (-1,)))
        home_short_key = ConfigManager.shortcut('home')
        if home_short_key:
            keyboard.add_hotkey(home_short_key, cls.executor.add, args=(lambda: KbExcel().move_to_home(), ()))
        show_status_short_key = ConfigManager.shortcut('show_status')

        reset_column_short_key = ConfigManager.shortcut('reset_column')
        if reset_column_short_key:
            keyboard.add_hotkey(reset_column_short_key, cls.executor.add,
                                args=(lambda: KbExcel().reset_column(), ()))
        show_status_short_key = ConfigManager.shortcut('show_status')

        if show_status_short_key:
            keyboard.add_hotkey(show_status_short_key, cls.show_column)
        insert_row_separator_short_key = ConfigManager.shortcut('insert_row_separator')
        if insert_row_separator_short_key:
            keyboard.add_hotkey(insert_row_separator_short_key, lambda x: KbExcel().insert_row_sperator(x),
                                args=(1,))

        one_or_multiple_cells_mode_short_key = ConfigManager.shortcut('one_or_multiple_cells_mode')
        if one_or_multiple_cells_mode_short_key:
            keyboard.add_hotkey(
                one_or_multiple_cells_mode_short_key, lambda: KbExcel().switch_one_or_multiple_cell_mode()
            )

        keyboard.add_hotkey("windows+l", cls.clear_pressed_events)
