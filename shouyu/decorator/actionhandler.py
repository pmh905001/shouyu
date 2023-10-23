import logging

from shouyu.config import Config
from shouyu.service.context import ExcelContext
from shouyu.view.msgbox import MessageBox, MessageType
from shouyu.util.process import ProcessManager


def action_handler(func):
    def wrapper(*args, **kwargs):
        try:
            result = func(*args, **kwargs)
            # Already opened Excel file may not be saved information by this program, so this program have to close it,
            # then try to save it again. After saved successfully, open Excel file agin.
            if ExcelContext.is_terminated_excel_and_reset():
                ProcessManager.open_file(Config.excel_path())
            return result
        except Exception as ex:
            logging.exception('caught unexpected exception!')
            MessageBox.pop_up_message('Failed', str(ex), MessageType.ERROR)
            return None

    return wrapper
