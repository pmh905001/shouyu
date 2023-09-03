import logging
import threading
import time

import keyboard
import pyperclip
from PIL import ImageGrab

from collector.chrome import ChromeCollector
from config import ConfigManager
from excel_context import ExcelContext
from excel_writer import ExcelWriter
from process import ProcessManager
from task_queue import TaskExecutor
from collector.basic_collector import BasicCollector


class Shortcut:
    executor = TaskExecutor()
    last_copy_time = 0
    last_show_time = 0

    @classmethod
    def start(cls):
        threading.Thread(target=cls.executor.run, daemon=True).start()

    @staticmethod
    def save_clipboard():
        try:
            img = ImageGrab.grabclipboard()
            copied_text = pyperclip.paste()
            ExcelWriter().save(img or copied_text)
        except:
            logging.exception('occurred un-expect exception!')

    @classmethod
    def save_clipboard_by_copy_2_times(cls):
        current_time = time.time()
        if current_time - cls.last_copy_time < 3:
            cls.save_clipboard()
        else:
            cls.last_copy_time = current_time

    @classmethod
    def generate_collector(cls):
        if BasicCollector.get_process_name() == 'chrome.exe':
            return ChromeCollector()
        else:
            return BasicCollector()

    @classmethod
    def one_key_save(cls):
        collector = cls.generate_collector()
        records = collector.collect_records()
        for index, record in enumerate(records):
            if index == 1:
                ExcelContext.steps = 1
            ExcelWriter().save(record, open_excel_again=False)
        ExcelContext.steps = -1

    @classmethod
    def show_column(cls):
        current_time = time.time()
        if current_time - cls.last_show_time < 3:
            ExcelWriter().move_column()
        else:
            cls.last_show_time = current_time

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
                                args=(lambda x: ExcelWriter().move_column(x), (1,)))
        move_to_left_short_key = ConfigManager.shortcut('move_to_left')
        if move_to_left_short_key:
            keyboard.add_hotkey(move_to_left_short_key, cls.executor.add,
                                args=(lambda x: ExcelWriter().move_column(x), (-1,)))
        home_short_key = ConfigManager.shortcut('home')
        if home_short_key:
            keyboard.add_hotkey(home_short_key, cls.executor.add, args=(lambda: ExcelWriter().move_to_home(), ()))
        show_status_short_key = ConfigManager.shortcut('show_status')
        if show_status_short_key:
            keyboard.add_hotkey(show_status_short_key, cls.show_column)
        insert_row_separator_short_key = ConfigManager.shortcut('insert_row_separator')
        if insert_row_separator_short_key:
            keyboard.add_hotkey(insert_row_separator_short_key, lambda x: ExcelWriter().insert_row_sperator(x),
                                args=(1,))
