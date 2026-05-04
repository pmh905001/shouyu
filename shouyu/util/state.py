"""Tiny JSON-backed state file for cross-launch persistence.

Used to remember things such as "habit dialog already shown today" so that
the user does not get spammed every time the program restarts.
"""
from __future__ import annotations

import json
import logging
import os
import threading
import time
from typing import Any, Optional


class AppState:
    FILE_NAME = 'shouyu_state.json'
    _lock = threading.Lock()
    _cache: Optional[dict] = None

    @classmethod
    def _path(cls) -> str:
        return os.path.abspath(cls.FILE_NAME)

    @classmethod
    def _load(cls) -> dict:
        if cls._cache is not None:
            return cls._cache
        path = cls._path()
        if not os.path.exists(path):
            cls._cache = {}
            return cls._cache
        try:
            with open(path, 'r', encoding='utf-8') as f:
                cls._cache = json.load(f) or {}
        except Exception:
            logging.exception(f'failed to load state file: {path}')
            cls._cache = {}
        return cls._cache

    @classmethod
    def _save(cls) -> None:
        if cls._cache is None:
            return
        path = cls._path()
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(cls._cache, f, ensure_ascii=False, indent=2)
        except Exception:
            logging.exception(f'failed to save state file: {path}')

    @classmethod
    def get(cls, key: str, default: Any = None) -> Any:
        with cls._lock:
            return cls._load().get(key, default)

    @classmethod
    def set(cls, key: str, value: Any) -> None:
        with cls._lock:
            data = cls._load()
            data[key] = value
            cls._save()

    @classmethod
    def delete(cls, key: str) -> None:
        with cls._lock:
            data = cls._load()
            data.pop(key, None)
            cls._save()

    @staticmethod
    def today_str() -> str:
        return time.strftime('%Y-%m-%d')

    @staticmethod
    def yesterday_str() -> str:
        return time.strftime('%Y-%m-%d', time.localtime(time.time() - 86400))

    @classmethod
    def streak_days(cls) -> int:
        return int(cls.get('streak_days', 0) or 0)

    # ---------- pomodoro mode ----------

    POMODORO_MODE_CLASSIC = 'classic'
    POMODORO_MODE_DEEP = 'deep'

    @classmethod
    def pomodoro_mode(cls) -> str:
        mode = cls.get('pomodoro_mode', cls.POMODORO_MODE_CLASSIC)
        if mode not in (cls.POMODORO_MODE_CLASSIC, cls.POMODORO_MODE_DEEP):
            return cls.POMODORO_MODE_CLASSIC
        return mode

    @classmethod
    def set_pomodoro_mode(cls, mode: str) -> None:
        if mode not in (cls.POMODORO_MODE_CLASSIC, cls.POMODORO_MODE_DEEP):
            mode = cls.POMODORO_MODE_CLASSIC
        cls.set('pomodoro_mode', mode)

    @classmethod
    def increment_today_counter(cls, key: str) -> int:
        """Bump a date-scoped counter. Stored as `<key>:YYYY-MM-DD`."""
        today = cls.today_str()
        full_key = f'{key}:{today}'
        with cls._lock:
            data = cls._load()
            current = int(data.get(full_key, 0) or 0) + 1
            data[full_key] = current
            cls._save()
            return current

    @classmethod
    def get_today_counter(cls, key: str) -> int:
        today = cls.today_str()
        full_key = f'{key}:{today}'
        return int(cls.get(full_key, 0) or 0)

    @classmethod
    def update_ritual_streak(cls) -> int:
        """Mark today's ritual as completed and return the resulting streak length.

        - Same day: idempotent (returns existing streak).
        - Yesterday was the last completion: increments streak.
        - Otherwise: resets streak to 1 (broken chain).
        """
        with cls._lock:
            data = cls._load()
            today = cls.today_str()
            last = data.get('last_ritual_date')
            if last == today:
                return int(data.get('streak_days', 1))
            yesterday = cls.yesterday_str()
            if last == yesterday:
                new_streak = int(data.get('streak_days', 0)) + 1
            else:
                new_streak = 1
            data['last_ritual_date'] = today
            data['streak_days'] = new_streak
            cls._save()
            return new_streak
