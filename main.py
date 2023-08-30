import logging
import threading

import keyboard

from shortcut import Shortcut
from tray import Tray

if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)

    tray = Tray.create()
    threading.Thread(target=tray.run, daemon=True).start()

    Shortcut.start()
    Shortcut.register_hot_keys()
    keyboard.wait()



