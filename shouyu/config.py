import os
import re
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
    def save_habits(cls, habits):
        """Rewrite the [habits] section's habit_1..habit_n entries in-place,
        preserving comments and every other section."""
        cls._load()  # ensures the ini file exists on disk
        path = cls.FILE_NAME
        with open(path, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        section_start = None
        section_end = len(lines)
        for i, line in enumerate(lines):
            stripped = line.strip()
            if stripped == '[habits]':
                section_start = i
                continue
            if section_start is not None and stripped.startswith('[') and stripped.endswith(']'):
                section_end = i
                break

        habit_lines = [f'habit_{i}={text}\n' for i, text in enumerate(habits, start=1)]

        if section_start is None:
            new_lines = list(lines)
            if new_lines and not new_lines[-1].endswith('\n'):
                new_lines[-1] += '\n'
            new_lines.append('\n[habits]\n')
            new_lines.extend(habit_lines)
        else:
            head = lines[:section_start + 1]
            tail = lines[section_end:]
            body = [
                line for line in lines[section_start + 1:section_end]
                if not re.match(r'habit_\d+\s*=', line.strip())
            ]
            trailing_blanks = []
            while body and body[-1].strip() == '':
                trailing_blanks.insert(0, body.pop())
            new_lines = head + body + habit_lines + trailing_blanks + tail

        with open(path, 'w', encoding='utf-8') as f:
            f.writelines(new_lines)
        cls.ini = None  # force a fresh parse on next read

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

    @classmethod
    def pomodoro_silence_when_locked(cls):
        """When true, suppress all pomodoro cues (transition beeps AND the
        hard idle alarm) while the workstation is locked — you've stepped away
        so nobody's there to hear it, and it avoids beeping on a shared/office
        machine after you lock and leave. Also stops the idle-drift alarm from
        escalating while locked (a deliberate lock isn't "drifting"). Default:
        true."""
        return cls.get('silence_when_locked', 'true', 'pomodoro').strip().lower() == 'true'

    @classmethod
    def pomodoro_break_reminder(cls):
        """How to announce the START of a break so you don't work straight
        through it:
          center = pop a centered "该休息了" card AND bring the timer window to
                   the front (calm, hard to miss, one-click to skip)
          off    = just quietly show the small floating window in the corner
                   (the old behavior)
        Default: center."""
        value = cls.get('break_reminder', 'center', 'pomodoro').strip().lower()
        return value if value in ('center', 'off') else 'center'

    @classmethod
    def pomodoro_deep_work_minutes(cls):
        """Work duration when running in 'deep' mode (default 90)."""
        return int(cls.get('deep_work_minutes', '90', 'pomodoro'))

    @classmethod
    def pomodoro_deep_short_break_minutes(cls):
        return int(cls.get('deep_short_break_minutes', '15', 'pomodoro'))

    @classmethod
    def pomodoro_deep_long_break_minutes(cls):
        return int(cls.get('deep_long_break_minutes', '30', 'pomodoro'))

    @classmethod
    def pomodoro_idle_warning_seconds(cls):
        """Level-1 idle threshold (seconds): the working-phase 🍅 starts a
        silent blink as a "stop drifting" nudge. Counts both mouse and
        keyboard inactivity. Set to 0 to disable. Default: 60."""
        return int(cls.get('idle_warning_seconds', '60', 'pomodoro'))

    @classmethod
    def pomodoro_idle_alarm_seconds(cls):
        """Level-2 (hard) idle threshold (seconds): escalate to a loud alarm
        — beep repeatedly, force the (possibly hidden) window back to the
        front, and require an explicit "我回来了" click to dismiss (which is
        recorded as a drift). Should be >= idle_warning_seconds. Set to 0 to
        disable escalation. Default: 120."""
        return int(cls.get('idle_alarm_seconds', '120', 'pomodoro'))

    @classmethod
    def pomodoro_idle_drifts_before_break(cls):
        """How many hard-alarm drifts within a single working phase are
        tolerated before we conclude you actually need rest and force a
        short break. Set to 0 to disable the forced break. Default: 3."""
        return int(cls.get('idle_drifts_before_break', '3', 'pomodoro'))

    _WEEKDAY_MAP = {
        'mon': 0, 'tue': 1, 'wed': 2, 'thu': 3, 'fri': 4, 'sat': 5, 'sun': 6,
    }

    @classmethod
    def pomodoro_quiet_days(cls):
        """Weekdays (as a set of ints, Mon=0..Sun=6) on which the 'auto'
        environment mode defaults to the quiet/office profile during the
        quiet window. Accepts mon..sun, numbers 0-6, or the shorthands
        weekday/weekend/all. Default: Mon-Fri."""
        return cls._parse_weekdays(cls.get('quiet_days', 'mon,tue,wed,thu,fri', 'pomodoro'))

    @classmethod
    def pomodoro_quiet_start(cls):
        """Start of the daily quiet window as (hour, minute). Default 09:00."""
        return cls._parse_hhmm(cls.get('quiet_start', '09:00', 'pomodoro'))

    @classmethod
    def pomodoro_quiet_end(cls):
        """End of the daily quiet window as (hour, minute). Default 18:00."""
        return cls._parse_hhmm(cls.get('quiet_end', '18:00', 'pomodoro'))

    @classmethod
    def _parse_weekdays(cls, value):
        result = set()
        for token in str(value).split(','):
            t = token.strip().lower()
            if not t:
                continue
            if t in ('weekday', 'weekdays'):
                result.update({0, 1, 2, 3, 4})
            elif t in ('weekend', 'weekends'):
                result.update({5, 6})
            elif t in ('all', 'everyday', 'daily'):
                result.update({0, 1, 2, 3, 4, 5, 6})
            elif t in cls._WEEKDAY_MAP:
                result.add(cls._WEEKDAY_MAP[t])
            else:
                try:
                    n = int(t)
                    if 0 <= n <= 6:
                        result.add(n)
                except ValueError:
                    pass
        return result

    @classmethod
    def pomodoro_default_mode(cls):
        """Default pomodoro mode used on a fresh launch (before the user has
        ever toggled the 深度/经典 button): 'deep' (90/15) or 'classic'
        (25/5). Default: deep."""
        mode = cls.get('default_mode', 'deep', 'pomodoro').strip().lower()
        return mode if mode in ('deep', 'classic') else 'deep'

    @classmethod
    def pomodoro_planning_enabled(cls):
        """When true, the first pomodoro of the day is a short "plan today's
        tasks" session instead of a normal work block. Default: true."""
        return cls.get('planning_enabled', 'true', 'pomodoro').strip().lower() == 'true'

    @classmethod
    def pomodoro_planning_session_minutes(cls):
        """Duration (minutes) of the morning planning session. Default: 10."""
        return int(cls.get('planning_session_minutes', '10', 'pomodoro'))

    @classmethod
    def pomodoro_planning_break_minutes(cls):
        """Duration (minutes) of the break that follows the morning planning
        session. Default: 5."""
        return int(cls.get('planning_break_minutes', '5', 'pomodoro'))

    @classmethod
    def pomodoro_lunch_enabled(cls):
        """When true, no working/planning phase is allowed to run during the
        configured lunch window; it becomes a "lunch break" instead.
        Default: true."""
        return cls.get('lunch_enabled', 'true', 'pomodoro').strip().lower() == 'true'

    @classmethod
    def pomodoro_lunch_start(cls):
        """Start of the lunch window as (hour, minute), or None if invalid.
        Default: 11:30."""
        return cls._parse_hhmm(cls.get('lunch_start', '11:30', 'pomodoro'))

    @classmethod
    def pomodoro_lunch_end(cls):
        """End of the lunch window as (hour, minute), or None if invalid.
        Default: 13:00."""
        return cls._parse_hhmm(cls.get('lunch_end', '13:00', 'pomodoro'))

    @staticmethod
    def _parse_hhmm(value):
        """Parse an 'HH:MM' string into a (hour, minute) tuple. Returns None
        when the value is missing or malformed."""
        try:
            parts = str(value).strip().split(':')
            hour = int(parts[0])
            minute = int(parts[1]) if len(parts) > 1 else 0
            if 0 <= hour < 24 and 0 <= minute < 60:
                return hour, minute
        except Exception:
            pass
        return None

    @classmethod
    def auto_prompt_duration_for_new_tasks(cls):
        """When you finish typing a brand-new task, pop up the duration picker
        so you commit to a time budget before moving on. Defaults to true."""
        return cls.get('auto_prompt_duration', 'true', 'planning').strip().lower() == 'true'

    @classmethod
    def overload_threshold_minutes(cls):
        """Total estimated minutes per day above which we show the overload warning."""
        return int(cls.get('overload_threshold_minutes', '360', 'planning'))


if __name__ == '__main__':
    print(Config.get('excel_path', 'kb.xlsx'))
