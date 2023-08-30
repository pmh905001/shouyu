import keyboard
import psutil
import pystray
from PIL import Image
from pystray import MenuItem

from config import ConfigManager
from package import Package
from process import ProcessManager


class Tray:
    _icon: pystray.Icon = None

    @classmethod
    def create(cls):
        menu = (
            MenuItem(text='重启', action=cls.on_restart),
            MenuItem(text='退出', action=cls.on_exit),
            MenuItem(text='显示', action=cls.on_show, default=True, visible=False),
        )
        icon = pystray.Icon("name", Image.open(Package.get_resource_path("fish.jpg")), "shouyu", menu)
        cls._icon = icon
        return icon

    @classmethod
    def on_exit(cls, icon, item):
        icon.stop()
        # sys.exit() can stop tray only, but the keyboard is still running.
        psutil.Process().terminate()

    @classmethod
    def on_restart(cls, icon, item):
        keyboard.remove_all_hotkeys()
        from shortcut import Shortcut
        Shortcut.register_hot_keys()

    @classmethod
    def on_show(cls, icon, item):
        ProcessManager.open(ConfigManager.excel_path())


if __name__ == '__main__':
    Tray.create()
