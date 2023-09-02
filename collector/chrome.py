import pyautogui
import pyperclip
import time

from collector.basic_collector import BasicCollector


class ChromeCollector(BasicCollector):

    def collect_records(self):
        time.sleep(1)
        # ImageGrab.grabclipboard()
        title = self.get_title()
        selected = self.get_selected_txt()
        url = self.get_url()
        return title, selected,url

    @staticmethod
    def get_url():
        pyautogui.press('f6')
        pyautogui.hotkey('ctrl', 'c')
        url = pyperclip.paste()
        # have to sleep 10 milliseconds
        time.sleep(1)
        return url


if __name__ == '__main__':
    time.sleep(1)

    print(ChromeCollector().collect_records())
