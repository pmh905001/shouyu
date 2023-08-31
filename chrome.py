import time

import win32gui
import win32process
import psutil


def getProcessName():
    time.sleep(3)
    hwnd = win32gui.GetForegroundWindow()
    print("handleId:", hwnd)
    print("title:" + win32gui.GetWindowText(hwnd))
    print(win32gui.GetClassName(hwnd))
    threadId, processId = win32process.GetWindowThreadProcessId(hwnd)
    print("processId:", processId)
    print("processName:", psutil.Process(processId).name())


if __name__ == '__main__':
    getProcessName()
