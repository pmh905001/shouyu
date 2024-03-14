import threading

from shouyu.service.context import ExcelContext
from shouyu.view.msgbox import MessageBox


def service_handler(func):
    def wrapper(*args, **kwargs):
        result = func(*args, **kwargs)

        from shouyu.service.excel import KbExcel

        kb_excel: KbExcel = args[0] if args else None
        if kb_excel:
            MessageBox.pop_up_message(**kb_excel._pop_up_msgs)
            # from multiprocessing import Process
            # Process(target=MessageBox.pop_up_message, kwargs=kb_excel._pop_up_msgs).start()
            # threading.Thread(target=MessageBox.pop_up_message, kwargs=kb_excel._pop_up_msgs).start()
            kb_excel._save_changed()
            ExcelContext.row_steps = 1

        return result

    return wrapper
