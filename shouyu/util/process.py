import logging
import os.path
import subprocess
import sys
import threading
import time

import psutil
from psutil import AccessDenied, NoSuchProcess


class ProcessManager:
    @staticmethod
    def is_file_path_accepted(file_path, proc):
        try:
            return file_path in ','.join(proc.cmdline())
        except AccessDenied:
            return False

    @staticmethod
    def is_process_name_accepted(proc):
        try:
            return proc.name() in ('wps.exe', '7zFM.exe', 'ms-excel.exe')
        except AccessDenied:
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
                        logging.error(f'process no longer exists (pid={old_pid})')
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
