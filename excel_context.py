import logging

from config import ConfigManager


class ExcelContext:
    excel_path = ConfigManager.get('excel_path', 'kb.xlsx')
    steps = 0

    @classmethod
    def get_steps_and_reset(cls):
        tmp = cls.steps
        cls.steps = 0
        return tmp
