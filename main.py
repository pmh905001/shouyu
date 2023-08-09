import logging

import keyboard
import pyperclip
from PIL import ImageGrab

from excel_context import ExcelContext
from excel_writer import ExcelWriter


def do_screen_short():
    img = ImageGrab.grabclipboard()
    copied_text = pyperclip.paste()
    ExcelWriter('kb.xlsx').save(img or copied_text)


def move_column(step=0):
    ExcelContext.steps += step
    logging.info(f'move {ExcelContext.steps} steps')


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    keyboard.add_hotkey('windows+shift+f', do_screen_short)
    keyboard.add_hotkey('windows+ctrl+right', move_column, args=(1,))
    keyboard.add_hotkey('windows+ctrl+left', move_column, args=(-1,))
    keyboard.wait()
