import logging
import sys
import webbrowser

import psutil
import pystray
from PIL import Image
from pystray import MenuItem

from shouyu.config import Config
from shouyu.util.package import Package
from shouyu.util.process import ProcessManager
from shouyu.util.reg import Registry


class Tray:
    _icon: pystray.Icon = None
    APP_NAME = 'shouyu'

    @classmethod
    def create(cls):
        icon = pystray.Icon(
            "name", Image.open(Package.get_resource_path('resources/icons/fish.jpg')), '授渔', cls._menu()
        )
        cls._icon = icon
        return icon

    @classmethod
    def _menu(cls):
        is_auto_run = Registry.is_auto_run(cls.APP_NAME)
        menu = (
            MenuItem(text='帮助', action=cls.on_help),
            MenuItem(text='设置', action=cls.on_config),
            MenuItem(
                text='非开机启动' if is_auto_run else '开机启动',
                action=cls.on_turn_off_auto_run if is_auto_run else cls.on_turn_on_auto_run
            ),
            MenuItem(text='重启', action=cls.on_restart),
            MenuItem(text='退出', action=cls.on_exit),
            MenuItem(text='显示', action=cls.on_show, default=True, visible=False),
        )
        return menu

    @classmethod
    def on_exit(cls, icon, item):
        logging.info('Stopping service!')
        icon.stop()
        # sys.exit() can stop tray only, but the keyboard is still running.
        psutil.Process().terminate()

    @classmethod
    def on_restart(cls, icon, item):
        ProcessManager.retart_myself()

    @classmethod
    def on_show(cls, icon, item):
        ProcessManager.open_file(Config.excel_path())

    @classmethod
    def on_help(cls, icon, item):
        webbrowser.open('https://gitee.com/pmh905001/shouyu/blob/main/README.md')

    @classmethod
    def on_turn_on_auto_run(cls, icon, item):
        cls._turn_auto_run(True)

    @classmethod
    def on_turn_off_auto_run(cls, icon, item):
        cls._turn_auto_run(False)

    @classmethod
    def _turn_auto_run(cls, turn_on: bool):
        Registry.set_auto_run(turn_on, cls.APP_NAME, sys.argv[0])
        cls._icon.menu = cls._menu()
        cls._icon.update_menu()

    @classmethod
    def on_config(cls, icon, item):
        ProcessManager.open_file('kb.ini')


if __name__ == '__main__':
    Tray.create()
