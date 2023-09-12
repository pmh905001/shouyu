import logging
import threading

import keyboard

from log import setup_log
from shortcut import Shortcut
from tray import Tray

if __name__ == '__main__':
    setup_log()
    logging.error('Started shouyu service!')
    tray = Tray.create()
    threading.Thread(target=tray.run, daemon=True).start()

    Shortcut.start()
    Shortcut.register_hot_keys()
    keyboard.wait()
