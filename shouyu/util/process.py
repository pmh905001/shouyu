import ctypes
import logging
import os.path
import subprocess
import sys
import threading
import time
from ctypes import wintypes

import psutil
from psutil import AccessDenied, NoSuchProcess


class ProcessManager:
    @staticmethod
    def is_file_path_accepted(file_path, proc):
        try:
            return file_path in ','.join(proc.cmdline())
        except (AccessDenied, NoSuchProcess, OSError):
            return False

    @staticmethod
    def is_process_name_accepted(proc):
        try:
            return proc.name() in ('wps.exe', '7zFM.exe', 'ms-excel.exe', 'EXCEL.EXE')
        except (AccessDenied, NoSuchProcess, OSError):
            logging.warning(f'Ignore pid: {proc.pid}')
            return False

    @classmethod
    def terminate_and_wait(cls, procs):
        for proc in procs:
            try:
                proc.terminate()
            except:
                logging.exception(f'terminate {proc} failed')
        time.sleep(1)

    @classmethod
    def terminate_by_path(cls, file_path: str):
        procs = [proc for proc in psutil.process_iter() if cls.is_file_path_accepted(file_path, proc)]
        if not procs:
            procs = [proc for proc in psutil.process_iter() if cls.is_process_name_accepted(proc)]

        if procs:
            cls.terminate_and_wait(procs)
        return procs

    @classmethod
    def graceful_close_by_path(cls, file_path: str):
        procs = [proc for proc in psutil.process_iter() if cls.is_file_path_accepted(file_path, proc)]
        if not procs:
            procs = [proc for proc in psutil.process_iter() if cls.is_process_name_accepted(proc)]

        if not procs:
            return []

        pids = set()
        for proc in procs:
            try:
                pids.add(proc.pid)
            except (AccessDenied, NoSuchProcess):
                pass

        cls._send_wm_close_to_windows(pids)

        deadline = time.time() + 5
        still_alive = list(procs)
        while time.time() < deadline:
            still_alive = [p for p in still_alive if cls._is_proc_running(p)]
            if not still_alive:
                break
            time.sleep(0.5)

        if still_alive:
            logging.warning('Graceful close timed out, force terminating')
            cls.terminate_and_wait(still_alive)

        cls._cleanup_excel_lock_file(file_path)
        return procs

    @staticmethod
    def _is_proc_running(proc):
        try:
            return proc.is_running() and proc.status() != psutil.STATUS_ZOMBIE
        except (AccessDenied, NoSuchProcess):
            return False

    @staticmethod
    def _send_wm_close_to_windows(pids):
        WM_CLOSE = 0x0010

        def callback(hwnd, _):
            pid = wintypes.DWORD()
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value in pids:
                ctypes.windll.user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)
            return True

        enum_proc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)(callback)
        ctypes.windll.user32.EnumWindows(enum_proc, 0)

    @staticmethod
    def _cleanup_excel_lock_file(file_path: str):
        dir_path = os.path.dirname(file_path) or '.'
        filename = os.path.basename(file_path)
        lock_file = os.path.join(dir_path, f'~${filename}')
        if os.path.exists(lock_file):
            try:
                os.remove(lock_file)
                logging.info(f'Removed lock file: {lock_file}')
            except OSError:
                logging.warning(f'Failed to remove lock file: {lock_file}')

    @staticmethod
    def open_file(cmd: str):
        # If call os.system(excel_path), hot keys of keyboard will be blocked, so create a new thread to execute.
        # os.system(excel_path) would result in a duplicated CMD window displayed, to avoid this issue use
        # subprocess.call() instead of os.system(excel_path)
        threading.Thread(
            target=subprocess.call,
            args=(cmd,),
            kwargs={
                'shell': True,
                'stdin': subprocess.PIPE,
                'stdout': subprocess.PIPE,
                'stderr': subprocess.PIPE
            }
        ).start()

    @staticmethod
    def kill_old_process():
        file_name = 'pid.txt'
        current_pid = psutil.Process().pid
        if os.path.exists(file_name):
            with open(file_name, 'r') as f:
                old_pid = int(f.read())
                if old_pid and old_pid != current_pid:
                    try:
                        proc = psutil.Process(old_pid)
                        proc.kill()
                        logging.info(f'killed pid: {old_pid}')
                    except NoSuchProcess:
                        logging.warning(f'process no longer exists (pid={old_pid})')
                    except:
                        logging.exception(f'failed to kill pid: {old_pid}')

        with open(file_name, 'w') as f:
            f.write(f'{current_pid}')

    @staticmethod
    def retart_myself():
        python = sys.executable
        os.execl(python, python, *sys.argv)


if __name__ == '__main__':
    ProcessManager.terminate_by_path('../../kb.xlsx')
