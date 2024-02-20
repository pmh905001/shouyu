import threading

from shouyu.view.msgbox import MessageBox


def service_handler(func):
    def wrapper(*args, **kwargs):
        result = func(*args, **kwargs)

        from shouyu.service.excel import KbExcel

        kb_excel: KbExcel = args[0] if args else None
        if kb_excel:
            threading.Thread(target=MessageBox.pop_up_message, kwargs=kb_excel._pop_up_msgs).start()
            # MessageBox.pop_up_message(**kb_excel._pop_up_msgs)
            kb_excel._save_changed()

        return result

    return wrapper
