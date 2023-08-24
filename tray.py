import os
import sys

import psutil
import pystray
from PIL import Image
from pystray import MenuItem

from excel_context import ExcelContext
from process import ProcessManager


class Tray:
    _icon: pystray.Icon = None

    @staticmethod
    def get_resource_path(relative_path):
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, relative_path)
        return os.path.join(os.path.abspath("."), relative_path)

    @classmethod
    def create(cls):
        menu = (
            MenuItem(text='退出', action=cls.on_exit),
            MenuItem(text='显示', action=cls.on_show, default=True, visible=False),

        )
        icon = pystray.Icon("name", Image.open(cls.get_resource_path("fish.jpg")), "shouyu", menu)
        cls._icon = icon
        return icon

    @classmethod
    def on_exit(cls, icon, item):
        icon.stop()
        # sys.exit() can stop tray only, but the keyboard is still running.
        psutil.Process().terminate()

    @classmethod
    def on_show(cls, icon, item):
        ProcessManager.open(ExcelContext.excel_path)


if __name__ == '__main__':
    Tray.create()
