import os
import shutil

import iniconfig

from shouyu.util.package import Package


class Config:
    FILE_NAME = 'kb.ini'
    ini = None

    @classmethod
    def _load(cls):
        if not os.path.exists(cls.FILE_NAME):
            source_path = Package.get_resource_path(cls.FILE_NAME)
            shutil.copy(source_path, cls.FILE_NAME)
        cls.ini = iniconfig.IniConfig(cls.FILE_NAME)
        return cls.ini

    @classmethod
    def get(cls, key, default, section='basic', convert=None):
        ini = cls._load()
        if convert:
            return ini.get(section, key, default, convert)
        else:
            return ini.get(section, key, default)

    @classmethod
    def excel_path(cls):
        return cls.get('excel_path', 'kb.xlsx')

    @classmethod
    def max_backups(cls):
        return int(cls.get('max_backups', '100'))

    @classmethod
    def get_shortcut(cls, key, default=None, convert=None):
        return cls.get(key, default, 'shortcuts', convert)

    @classmethod
    def habits(cls):
        """Return the habit reminder list, ordered by the numeric suffix of habit_<n>."""
        ini = cls._load()
        section = ini.sections.get('habits') or {}

        items = []
        for key, raw_value in section.items():
            value = (raw_value or '').strip()
            if not key.startswith('habit_') or not value:
                continue
            suffix = key[len('habit_'):]
            try:
                order = int(suffix)
            except ValueError:
                order = 0
            items.append((order, value))
        items.sort(key=lambda x: x[0])
        return [text for _, text in items]

    @classmethod
    def pomodoro_enabled(cls):
        return cls.get('enabled', 'false', 'pomodoro').strip().lower() == 'true'

    @classmethod
    def pomodoro_work_minutes(cls):
        return int(cls.get('work_minutes', '25', 'pomodoro'))

    @classmethod
    def pomodoro_short_break_minutes(cls):
        return int(cls.get('short_break_minutes', '5', 'pomodoro'))

    @classmethod
    def pomodoro_long_break_minutes(cls):
        return int(cls.get('long_break_minutes', '20', 'pomodoro'))

    @classmethod
    def pomodoro_cycles_before_long_break(cls):
        return int(cls.get('cycles_before_long_break', '4', 'pomodoro'))

    @classmethod
    def pomodoro_notify_sound(cls):
        return cls.get('notify_sound', 'true', 'pomodoro').strip().lower() == 'true'


if __name__ == '__main__':
    print(Config.get('excel_path', 'kb.xlsx'))
