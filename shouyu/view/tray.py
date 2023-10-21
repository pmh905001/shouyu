import logging
import os
import webbrowser
import winreg

import psutil
import pystray
import sys
from PIL import Image
from pystray import MenuItem

from shouyu.config import Config
from shouyu.util.package import Package
from shouyu.util.process import ProcessManager


class Tray:
    _icon: pystray.Icon = None

    @classmethod
    def create(cls):
        menu = (
            MenuItem(text='帮助', action=cls.on_help),
            MenuItem(text='设置&快捷键', action=cls.on_config),
            MenuItem(text='开启开机自启动', action=cls.on_turn_on_auto_run),
            # TODO: not works, not sure what is wrong
            # MenuItem(text='关闭开机自启动', action=cls.on_turn_off_auto_run),
            MenuItem(text='重启', action=cls.on_restart),
            MenuItem(text='退出', action=cls.on_exit),
            MenuItem(text='显示', action=cls.on_show, default=True, visible=False),
        )
        icon = pystray.Icon("name", Image.open(Package.get_resource_path('resources/icons/fish.jpg')), '授渔', menu)
        cls._icon = icon
        return icon

    @classmethod
    def on_exit(cls, icon, item):
        icon.stop()
        # sys.exit() can stop tray only, but the keyboard is still running.
        psutil.Process().terminate()
        logging.info('Stopped service!')

    @classmethod
    def on_restart(cls, icon, item):
        python = sys.executable
        os.execl(python, python, *sys.argv)

    @classmethod
    def on_show(cls, icon, item):
        ProcessManager.open(Config.excel_path())

    @classmethod
    def on_help(cls, icon, item):
        webbrowser.open('https://gitee.com/pmh905001/shouyu/blob/main/README.md')

    @classmethod
    def on_turn_on_auto_run(cls, icon, item):
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_ALL_ACCESS
        )
        winreg.SetValueEx(key, 'shouyu', 0, winreg.REG_SZ, sys.argv[0])
        winreg.CloseKey(key)
        logging.error('registered key')

    @classmethod
    def on_turn_off_auto_run(cls, icon, item):
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_ALL_ACCESS
        )
        winreg.DeleteKey(key, 'shouyu')
        winreg.CloseKey(key)
        logging.error('deleted key')

    @classmethod
    def on_config(cls, icon, item):
        ProcessManager.open('kb.ini')


if __name__ == '__main__':
    Tray.create()
