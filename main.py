import logging
import threading
import time

import keyboard
import pyperclip
from PIL import ImageGrab

from config import ConfigManager
from excel_writer import ExcelWriter
from process import ProcessManager
from task_queue import TaskExecutor
from tray import Tray

last_copy_time = 0
last_show_time = 0


def save_clipboard():
    try:
        img = ImageGrab.grabclipboard()
        copied_text = pyperclip.paste()
        ExcelWriter().save(img or copied_text)
    except:
        logging.exception('occurred un-expect exception!')


def save_clipboard_by_copy_2_times():
    global last_copy_time
    current_time = time.time()
    if current_time - last_copy_time < 3:
        save_clipboard()
    else:
        last_copy_time = current_time


def show_column():
    global last_show_time
    current_time = time.time()
    if current_time - last_show_time < 3:
        ExcelWriter().move_column()
    else:
        last_show_time = current_time


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)

    executor = TaskExecutor()
    threading.Thread(target=executor.run, daemon=True).start()

    # save clipboard to kb.xlsx
    copy_2_times_short_key = ConfigManager.shortcut('save_clipboard_by_copy')
    if copy_2_times_short_key:
        keyboard.add_hotkey(copy_2_times_short_key, executor.add, args=(save_clipboard_by_copy_2_times, ()))
    save_clipboard_short_key = ConfigManager.shortcut('save_clipboard')
    if save_clipboard_short_key:
        keyboard.add_hotkey(save_clipboard_short_key, executor.add, args=(save_clipboard, ()))
    print_screen_short_key = ConfigManager.shortcut('print_screen')
    if print_screen_short_key:
        keyboard.add_hotkey(print_screen_short_key, executor.add, args=(save_clipboard_by_copy_2_times, ()))
    windows_print_screen_short_key = ConfigManager.shortcut('windows_print_screen')
    if windows_print_screen_short_key:
        keyboard.add_hotkey(windows_print_screen_short_key, executor.add, args=(save_clipboard_by_copy_2_times, ()))

    # Open or close kb.xlsx
    open_excel_short_key = ConfigManager.shortcut('open_excel')
    if open_excel_short_key:
        keyboard.add_hotkey(open_excel_short_key, ProcessManager.open, args=(ConfigManager.excel_path(),))
    close_excel_short_key = ConfigManager.shortcut('close_excel')
    if close_excel_short_key:
        keyboard.add_hotkey(close_excel_short_key, ProcessManager.terminate_by_path, args=(ConfigManager.excel_path(),))

    # show or move current column position
    move_to_right_short_key = ConfigManager.shortcut('move_to_right')
    if move_to_right_short_key:
        keyboard.add_hotkey(move_to_right_short_key, executor.add, args=(lambda x: ExcelWriter().move_column(x), (1,)))
    move_to_left_short_key = ConfigManager.shortcut('move_to_left')
    if move_to_left_short_key:
        keyboard.add_hotkey(move_to_left_short_key, executor.add, args=(lambda x: ExcelWriter().move_column(x), (-1,)))
    home_short_key = ConfigManager.shortcut('home')
    if home_short_key:
        keyboard.add_hotkey(home_short_key, executor.add, args=(lambda: ExcelWriter().move_to_home(), ()))
    show_status_short_key = ConfigManager.shortcut('show_status')
    if show_status_short_key:
        keyboard.add_hotkey(show_status_short_key, executor.add, args=(show_column, ()))

    insert_row_separator_short_key = ConfigManager.shortcut('insert_row_separator')
    if insert_row_separator_short_key:
        keyboard.add_hotkey(insert_row_separator_short_key, lambda x: ExcelWriter().insert_row_sperator(x), args=(1,))

    icon = Tray.create()
    threading.Thread(target=icon.run, daemon=True).start()
    keyboard.wait()
