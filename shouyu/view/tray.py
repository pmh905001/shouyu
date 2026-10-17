import logging
import os
import sys
import tempfile
import webbrowser

import psutil
import pystray
from PIL import Image
from pystray import Menu, MenuItem

from shouyu.config import Config
from shouyu.util.package import Package
from shouyu.util.process import ProcessManager
from shouyu.util.reg import Registry


class Tray:
    _icon: pystray.Icon = None
    APP_NAME = 'shouyu'
    ICON_IMAGE = 'resources/icons/fish.png'
    ICON_TILE = '授渔'

    @classmethod
    def create(cls):
        icon = pystray.Icon(
            cls.APP_NAME, Image.open(Package.get_resource_path(cls.ICON_IMAGE)), cls.ICON_TILE, cls._menu_items()
        )
        cls._icon = icon
        return icon

    @classmethod
    def _menu_items(cls):
        menu = (
            MenuItem(text='帮助', action=cls.on_help),
            MenuItem(text='设置', action=cls.on_config),
            MenuItem(text='打开工作目录', action=cls.on_open_work_dir),
            MenuItem(text='今日任务', action=cls.on_show_todo),
            MenuItem(text='番茄: 启停', action=cls.on_toggle_pomodoro),
            MenuItem(text='番茄: 暂停/继续', action=cls.on_pause_resume_pomodoro),
            MenuItem(text='番茄: 显示/隐藏', action=cls.on_toggle_pomodoro_visible),
            MenuItem(text='从备份恢复…', action=cls.on_restore_backup),
            MenuItem(text='查看待同步队列', action=cls.on_show_queue),
            MenuItem(text='查看错误日志', action=cls.on_show_log),
            MenuItem(text='查看崩溃诊断日志', action=cls.on_show_crash_log),
            MenuItem(text='开机启动', action=cls.on_turn_on_or_off_auto_running, checked=cls.display_checked),
            MenuItem(text='重启', action=cls.on_restart),
            MenuItem(text='显示Excel', action=cls.on_show, default=True, visible=False),
            MenuItem(text='退出', action=cls.on_exit),
        )
        return menu

    @classmethod
    def on_exit(cls, icon, item):
        logging.info('Stopping service!')
        ProcessManager.mark_shutdown_requested()
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
    def on_turn_on_or_off_auto_running(cls, icon, item):
        is_auto_run = Registry.is_auto_run(cls.APP_NAME)
        Registry.set_auto_run(not is_auto_run, cls.APP_NAME, sys.argv[0])

    @classmethod
    def on_config(cls, icon, item):
        ProcessManager.open_file(Config.FILE_NAME)

    @classmethod
    def on_open_work_dir(cls, icon, item):
        ProcessManager.open_folder(Config.work_dir())

    @classmethod
    def on_show_todo(cls, icon, item):
        from shouyu.view.qt_app import QtApp

        QtApp.show_todo_panel()

    @classmethod
    def on_toggle_pomodoro(cls, icon, item):
        from shouyu.service.pomodoro import PomodoroService
        from shouyu.view.pomodoro_window import PomodoroWindow
        from shouyu.view.qt_app import QtApp

        # If the timer is mid-cycle but the window is hidden, the user almost
        # certainly wants to see it again rather than pause/resume the timer.
        # Treat the toggle hotkey as "summon" in that case.
        snap = PomodoroService.instance().snapshot()
        running = snap.get("phase") in ("working", "short_break", "long_break", "paused")
        if running and not PomodoroWindow.is_visible_safe():
            QtApp.show_pomodoro_window()
            return
        PomodoroService.instance().toggle()

    @classmethod
    def on_pause_resume_pomodoro(cls, icon, item):
        """Pause a running cycle, or resume a paused one. Idle is a no-op
        (use 番茄: 启停 to start). Mirrors the state handling in
        PomodoroService.toggle() without the start branch."""
        from shouyu.service.pomodoro import PomodoroService

        service = PomodoroService.instance()
        phase = service.snapshot().get('phase')
        if phase == 'idle':
            logging.info('pomodoro is idle; nothing to pause or resume')
        elif phase == 'paused':
            service.resume()
        else:
            service.pause()

    @classmethod
    def on_toggle_pomodoro_visible(cls, icon, item):
        """Show the floating ball if hidden, hide it if visible. Same action
        as the show/hide hotkey (ctrl+alt+t)."""
        from shouyu.view.qt_app import QtApp

        QtApp.toggle_pomodoro_window()

    @classmethod
    def on_restore_backup(cls, icon, item):
        from shouyu.view.qt_app import QtApp

        QtApp.show_backup_restore()

    @classmethod
    def on_show_queue(cls, icon, item):
        """Dump the durable message-queue contents to a text file and open
        it, for troubleshooting stuck/failed saves. See
        docs/excel-save-resilience.md §8 / §7 item 9."""
        from shouyu.service import message_queue

        path = os.path.join(tempfile.gettempdir(), 'shouyu_queue_dump.txt')
        try:
            with open(path, 'w', encoding='utf-8') as f:
                f.write(message_queue.dump_human_readable())
            ProcessManager.open_file(path)
        except Exception:
            logging.exception('failed to dump/open queue contents')

    @classmethod
    def on_show_log(cls, icon, item):
        path = Config.log_path()
        if not os.path.exists(path):
            logging.info('error log does not exist yet: %s', path)
            return
        ProcessManager.open_log(path)

    @classmethod
    def on_show_crash_log(cls, icon, item):
        path = Config.crash_log_path()
        if not os.path.exists(path):
            # Nothing has crashed yet; create the file so the menu still opens
            # something useful instead of silently doing nothing.
            try:
                os.makedirs(os.path.dirname(path) or '.', exist_ok=True)
                open(path, 'a', encoding='utf-8').close()
            except OSError:
                logging.exception('failed to create crash diagnostics log: %s', path)
                return
        ProcessManager.open_log(path)

    @classmethod
    def display_checked(cls, icon):
        return Registry.is_auto_run(cls.APP_NAME)
