import logging

from shouyu.config import ConfigManager
from shouyu.msgbox import MessageBox, MessageType


def exception_handler(func):
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as ex:
            logging.exception('caught unexpected exception!')
            MessageBox.pop_up_message('Failed', str(ex), MessageType.ERROR)
            return None

    return wrapper
