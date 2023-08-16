import logging
import threading
import time

import keyboard
import pyperclip
from PIL import ImageGrab

from excel_writer import ExcelWriter
from process import ProcessManager
from tray import Tray

last_copy_time = 0
last_ctrl_time = 0


def save_clipboard():
    try:
        img = ImageGrab.grabclipboard()
        copied_text = pyperclip.paste()
        ExcelWriter('kb.xlsx').save(img or copied_text)
    except:
        logging.exception('occurred un-expect exception!')


def save_clipboard_by_copy_2_times():
    global last_copy_time
    current_time = time.time()
    if current_time - last_copy_time < 3:
        save_clipboard()
    else:
        last_copy_time = current_time


def show_column_by_shift_2_times():
    global last_ctrl_time
    current_time = time.time()
    if current_time - last_ctrl_time < 3:
        ExcelWriter('kb.xlsx').move_column(0)
    else:
        last_ctrl_time = current_time


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    # save clipboard to kb.xlsx
    keyboard.add_hotkey('ctrl+c', save_clipboard_by_copy_2_times)
    keyboard.add_hotkey('ctrl+enter', save_clipboard)
    # Open or close kb.xlsx
    keyboard.add_hotkey('ctrl+\\', ProcessManager.open, args=('kb.xlsx',))
    keyboard.add_hotkey('ctrl+q', ProcessManager.terminate_by_path, args=('kb.xlsx',))
    # show or move current column position
    keyboard.add_hotkey('ctrl+right', ExcelWriter('kb.xlsx').move_column, args=(1,))
    keyboard.add_hotkey('ctrl+left', ExcelWriter('kb.xlsx').move_column, args=(-1,))
    keyboard.add_hotkey('shift', show_column_by_shift_2_times)

    icon = Tray.create()
    threading.Thread(target=icon.run, daemon=False).start()
    keyboard.wait()
