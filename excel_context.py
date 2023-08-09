import logging


class ExcelContext:
    steps = 0

    @classmethod
    def get_steps_and_reset(cls):
        tmp = cls.steps
        cls.steps = 0
        return tmp
