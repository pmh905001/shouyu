import logging
import time

import keyboard
import pyperclip
from PIL import ImageGrab

from excel_context import ExcelContext
from excel_writer import ExcelWriter


def do_screen_short():
    img = ImageGrab.grabclipboard()
    copied_text = pyperclip.paste()
    ExcelWriter('kb.xlsx').save(img or copied_text)


last_time = 0
def do_copy_2_times():
    global last_time
    current_time = time.time()
    if current_time - last_time < 3:
        do_screen_short()
    else:
        last_time = current_time


def move_column(step=0):
    ExcelContext.steps += step
    logging.info(f'move {ExcelContext.steps} steps')


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    keyboard.add_hotkey('windows+shift+f', do_screen_short)
    keyboard.add_hotkey('windows+ctrl+right', move_column, args=(1,))
    keyboard.add_hotkey('windows+ctrl+left', move_column, args=(-1,))
    keyboard.add_hotkey('ctrl+c', do_copy_2_times)
    keyboard.wait()
