import logging
import threading
import time

import keyboard
import pyperclip
from PIL import ImageGrab

from excel_context import ExcelContext
from excel_writer import ExcelWriter
from process import ProcessManager
from tray import Tray


def save_clipboard():
    try:
        img = ImageGrab.grabclipboard()
        copied_text = pyperclip.paste()
        ExcelWriter('kb.xlsx').save(img or copied_text)
    except:
        logging.exception('occurred un-expect exception!')


last_time = 0


def do_copy_2_times():
    global last_time
    current_time = time.time()
    if current_time - last_time < 3:
        save_clipboard()
    else:
        last_time = current_time


def move_column(step=0):
    ExcelContext.steps += step
    logging.info(f'move {ExcelContext.steps} steps')


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    keyboard.add_hotkey('ctrl+enter', save_clipboard)
    keyboard.add_hotkey('ctrl+\\', ProcessManager.resume_last_closed_process, args=('kb.xlsx',))
    keyboard.add_hotkey('ctrl+q', ProcessManager.terminate_by_path, args=('kb.xlsx',))
    keyboard.add_hotkey('ctrl+right', move_column, args=(1,))
    keyboard.add_hotkey('ctrl+left', move_column, args=(-1,))
    keyboard.add_hotkey('ctrl+c', do_copy_2_times)
    icon=Tray.create()
    threading.Thread(target=icon.run, daemon=False).start()
    keyboard.wait()
