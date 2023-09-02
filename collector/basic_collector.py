import psutil
import pyautogui
import pyperclip
import win32gui
import win32process
from PIL import ImageGrab


class BasicCollector:

    @staticmethod
    def get_process_name():
        hwnd = win32gui.GetForegroundWindow()
        threadId, processId = win32process.GetWindowThreadProcessId(hwnd)
        return psutil.Process(processId).name()

    def get_title(self):
        hwnd = win32gui.GetForegroundWindow()
        title = win32gui.GetWindowText(hwnd)
        return title

    @staticmethod
    def get_selected_txt():
        pyperclip.copy("")
        pyautogui.hotkey('ctrl', 'c')
        selected = pyperclip.paste()
        return selected

    def collect_records(self):
        return self.get_title(), self.get_selected_txt(), ImageGrab.grabclipboard()
