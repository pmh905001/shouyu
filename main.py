import logging
import threading

import keyboard

from shouyu.log import setup_log
from shouyu.action.shortcut import Shortcut
from shouyu.util.process import ProcessManager
from shouyu.view.tray import Tray

if __name__ == '__main__':
    setup_log()
    ProcessManager.kill_old_process()
    logging.info('Started service!')
    tray = Tray.create()
    threading.Thread(target=tray.run, daemon=True).start()

    Shortcut.start()
    Shortcut.register_hot_keys()
    keyboard.wait()
