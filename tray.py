import psutil
import pystray
from PIL import Image
from pystray import MenuItem


class Tray:
    _icon: pystray.Icon = None

    @classmethod
    def create(cls):
        menu = (
            MenuItem(text='退出', action=cls.on_exit),
        )
        icon = pystray.Icon("name", Image.open("fish.png"), "shouyu", menu)
        cls._icon = icon
        return icon

    @classmethod
    def on_exit(cls, icon, item):
        icon.stop()
        psutil.Process().terminate()
        # sys.exit()


if __name__ == '__main__':
    Tray.create()
