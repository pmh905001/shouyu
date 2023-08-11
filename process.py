import logging
import os
import time

import psutil
from psutil import AccessDenied

import threading


class ProcessManager:
    _last_closed_processes: psutil.Process = None

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
    def resume_last_closed_process(excel_path):
        threading.Thread(target=os.system,args=(excel_path,)).start()
        # os.system(excel_path)


if __name__ == '__main__':
    ProcessManager.terminate_by_path('kb.xlsx')
