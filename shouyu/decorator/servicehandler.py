from shouyu.view.msgbox import MessageBox


def service_handler(func):
    def wrapper(*args, **kwargs):
        result = func(*args, **kwargs)

        from shouyu.service.excel import KbExcel

        kb_excel: KbExcel = args[0] if args else None
        if kb_excel:
            kb_excel._save_changed()
            MessageBox.pop_up_message(**kb_excel._pop_up_msgs)
        return result

    return wrapper
