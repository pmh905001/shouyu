import psutil
import pystray
from PIL import Image
from pystray import MenuItem
import sys

class Tray:
    @classmethod
    def create(cls):
        menu = (
            MenuItem(text='退出', action=cls.on_exit),
        )
        icon = pystray.Icon("name", Image.open("fish.png"), "shouyu", menu)
        icon.run()

    @classmethod
    def on_exit(cls, icon, item):
        icon.stop()
        psutil.Process().terminate()
        # sys.exit()

if __name__=='__main__':
    Tray.create()