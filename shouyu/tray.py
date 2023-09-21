import logging
import os

import psutil
import pystray
import sys
from PIL import Image
from pystray import MenuItem

from shouyu.config import ConfigManager
from shouyu.package import Package
from shouyu.process import ProcessManager


class Tray:
    _icon: pystray.Icon = None

    @classmethod
    def create(cls):
        menu = (
            MenuItem(text='重启', action=cls.on_restart),
            MenuItem(text='退出', action=cls.on_exit),
            MenuItem(text='显示', action=cls.on_show, default=True, visible=False),
        )
        icon = pystray.Icon("name", Image.open(Package.get_resource_path("shouyu\\fish.jpg")), "授渔", menu)
        cls._icon = icon
        return icon

    @classmethod
    def on_exit(cls, icon, item):
        icon.stop()
        # sys.exit() can stop tray only, but the keyboard is still running.
        psutil.Process().terminate()
        logging.info('Stopped shouyu service!')

    @classmethod
    def on_restart(cls, icon, item):
        python = sys.executable
        os.execl(python, python, *sys.argv)

    @classmethod
    def on_show(cls, icon, item):
        ProcessManager.open(ConfigManager.excel_path())


if __name__ == '__main__':
    Tray.create()
