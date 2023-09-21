import logging

from config import ConfigManager
from msg_box import MessageBox, MessageType


def exception_handler(func):
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as ex:
            logging.exception('caught unexpected exception!')
            duration = ConfigManager.shortcut('save_clipboard_popup_duration', '1', lambda x: int(x))
            MessageBox.pop_up_message('Failed', str(ex), MessageType.ERROR, duration)
            return None

    return wrapper
