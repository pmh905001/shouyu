import logging
import winreg


class Registry:
    AUTO_RUN_KEY = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Run'

    @classmethod
    def set_auto_run(cls, generate: bool, program_name: str, program_path: str):
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, cls.AUTO_RUN_KEY, 0, winreg.KEY_ALL_ACCESS)
        if generate:
            winreg.SetValueEx(key, program_name, 0, winreg.REG_SZ, program_path)
            logging.info('turned on auto run')
        else:
            winreg.DeleteValue(key, program_name)
            logging.info('turned off auto run')
        winreg.CloseKey(key)

    @classmethod
    def is_auto_run(cls, program_name: str):
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, cls.AUTO_RUN_KEY, 0, winreg.KEY_ALL_ACCESS)
        try:
            program_cmd, _ = winreg.QueryValueEx(key, program_name)
            return True if program_cmd else False
        except FileNotFoundError:
            logging.warning(f'program name {program_name} not registered')
            return False
