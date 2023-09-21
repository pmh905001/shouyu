import pyautogui
import pyperclip
import time

from shouyu.collector.basic_collector import BasicCollector


class ChromeCollector(BasicCollector):

    def collect_records(self):
        # sleep 500 milliseconds to wait first shortcut (ctrl+alt+enter) released
        time.sleep(0.5)
        title = self.get_title()
        selected = self.get_selected_txt()
        url = self.get_url()
        return title, selected, url

    @staticmethod
    def get_url():
        pyautogui.press('f6')
        pyautogui.hotkey('ctrl', 'c')
        url = pyperclip.paste()
        # sleep 10 milliseconds to get content of clipboard
        time.sleep(0.01)
        return url

    def get_title(self):
        title = super().get_title()
        return title.replace(' - Google Chrome', '') if title else title


if __name__ == '__main__':
    time.sleep(1)

    print(ChromeCollector().collect_records())
