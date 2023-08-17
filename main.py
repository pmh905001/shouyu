import logging
import threading
import time

import keyboard
import pyperclip
from PIL import ImageGrab

from excel_context import ExcelContext
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
        ExcelWriter().move_column(0)
    else:
        last_show_time = current_time


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)

    executor = TaskExecutor()
    threading.Thread(target=executor.run, daemon=True).start()

    # save clipboard to kb.xlsx
    keyboard.add_hotkey('ctrl+c', executor.add, args=(save_clipboard_by_copy_2_times, ()))
    keyboard.add_hotkey('ctrl+enter', executor.add, args=(save_clipboard, ()))
    keyboard.add_hotkey('print screen', executor.add, args=(save_clipboard_by_copy_2_times, ()))
    keyboard.add_hotkey('alt+print screen', executor.add, args=(save_clipboard_by_copy_2_times, ()))
    keyboard.add_hotkey('windows+print screen', executor.add, args=(save_clipboard_by_copy_2_times, ()))
    # Open or close kb.xlsx
    keyboard.add_hotkey('ctrl+\\', ProcessManager.open, args=(ExcelContext.excel_path,))
    keyboard.add_hotkey('ctrl+q', ProcessManager.terminate_by_path, args=(ExcelContext.excel_path,))
    # show or move current column position
    keyboard.add_hotkey('tab+right', executor.add, args=(lambda x: ExcelWriter().move_column(x), (1,)))
    keyboard.add_hotkey('tab+left', executor.add, args=(lambda x: ExcelWriter().move_column(x), (-1,)))
    keyboard.add_hotkey('home', executor.add, args=(lambda: ExcelWriter().move_to_home(), ()))
    keyboard.add_hotkey('alt', executor.add, args=(show_column, ()))

    icon = Tray.create()
    threading.Thread(target=icon.run, daemon=True).start()
    keyboard.wait()
