import logging
import os
from enum import Enum

import time
import win32con
import win32gui
from win11toast import toast


class MessageType(Enum):
    SUCCESS = 0
    ERROR = 1


class MessageBox:
    def __init__(self):
        wc = win32gui.WNDCLASS()
        hinst = wc.hInstance = win32gui.GetModuleHandle(None)
        wc.lpszClassName = f"MessageBox_{time.time()}"
        wc.lpfnWndProc = {win32con.WM_DESTROY: self.OnDestroy, }
        classAtom = win32gui.RegisterClass(wc)
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = win32gui.CreateWindow(
            classAtom, "Message Box", style, 0, 0, win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT, 0, 0, hinst, None
        )
        hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)
        nid = (self.hwnd, 0, win32gui.NIF_ICON, win32con.WM_USER + 20, hicon, "MessageBox")
        win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)

    def showMsg(self, title, msg, level: int = MessageType.SUCCESS):
        nid = (self.hwnd,
               0,
               win32gui.NIF_INFO,
               0,
               0,
               "MessageBox",
               msg,
               0,
               title,
               win32gui.NIIF_INFO + win32gui.NIIF_NOSOUND if level == MessageType.SUCCESS else win32gui.NIIF_ERROR
               )
        win32gui.Shell_NotifyIcon(win32gui.NIM_MODIFY, nid)

    def OnDestroy(self, hwnd, msg, wparam, lparam):
        nid = (self.hwnd, 0)
        win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
        win32gui.PostQuitMessage(0)  # Terminate the app.
        return 0

    @classmethod
    def pop_up_message(cls, title, msg, level: MessageType = MessageType.SUCCESS, duration: int = 5, image_path=None):
        # msg_box = MessageBox()
        # msg_box.showMsg(title, msg, level)
        # time.sleep(duration)
        # win32gui.DestroyWindow(msg_box.hwnd)
        logging.info("pop up message!")
        toast(title, msg, image=image_path, duration='short', audio={'silent': 'true'})


if __name__ == '__main__':
    # MessageBox.pop_up_message("Failed", "test", MessageType.ERROR)
    # MessageBox.pop_up_message("Failed2", "test2", MessageType.ERROR)
    # MessageBox.pop_up_message("Failed3", "test3", MessageType.ERROR)
    # MessageBox.pop_up_message("Failed4", "test4", MessageType.SUCCESS)
    from excel_writer import ExcelWriter

    MessageBox.pop_up_message(
        "Failed4", "test4", MessageType.SUCCESS, image_path=os.path.abspath(ExcelWriter.IMAGE_PATH)
    )
