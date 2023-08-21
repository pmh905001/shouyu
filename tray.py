import psutil
import pystray
from PIL import Image
from pystray import MenuItem

from process import ProcessManager


class Tray:
    _icon: pystray.Icon = None

    @classmethod
    def create(cls):
        menu = (
            MenuItem(text='退出', action=cls.on_exit),
            MenuItem(text='显示', action=cls.on_show, default=True, visible=False),

        )
        icon = pystray.Icon("name", Image.open("fish.jpg"), "shouyu", menu)
        cls._icon = icon
        return icon

    @classmethod
    def on_exit(cls, icon, item):
        icon.stop()
        # sys.exit() can stop tray only, but the keyboard is still running.
        psutil.Process().terminate()

    @classmethod
    def on_show(cls, icon, item):
        ProcessManager.open('kb.xlsx')


if __name__ == '__main__':
    Tray.create()
