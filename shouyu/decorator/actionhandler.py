import logging

from shouyu.config import Config
from shouyu.service.context import ExcelContext
from shouyu.ui.msgbox import MessageBox, MessageType
from shouyu.utils.process import ProcessManager


def action_handler(func):
    def wrapper(*args, **kwargs):
        try:
            result = func(*args, **kwargs)
            if ExcelContext.is_terminated_excel_and_reset():
                ProcessManager.open(Config.excel_path())
            return result
        except Exception as ex:
            logging.exception('caught unexpected exception!')
            MessageBox.pop_up_message('Failed', str(ex), MessageType.ERROR)
            return None

    return wrapper
