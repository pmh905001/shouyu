"""Pomodoro state machine.

Lives outside Qt so the timer keeps ticking even when no UI is visible.
UI updates are dispatched through QtApp.emit_pomodoro_event(...) which is
thread-safe (queued connection into the Qt event loop).

Events emitted (event_name, payload):
    "started"            -> "<phase>:<duration_seconds>:<task_text>"
    "tick"               -> "<remaining_seconds>"
    "phase_changed"      -> "<new_phase>:<duration_seconds>:<task_text>"
    "paused"             -> ""
    "resumed"            -> ""
    "stopped"            -> ""
    "idle_warning_on"    -> "<idle_seconds>"   (working phase only)
    "idle_warning_off"   -> ""
"""
from __future__ import annotations

import io
import logging
import math
import struct
import threading
import time
import wave
from enum import Enum
from typing import Optional

from shouyu.config import Config


class Phase(str, Enum):
    IDLE = "idle"
    PLANNING = "planning"
    WORKING = "working"
    SHORT_BREAK = "short_break"
    LONG_BREAK = "long_break"
    LUNCH_BREAK = "lunch_break"
    PAUSED = "paused"


_SOFT_TONE_WAV: Optional[bytes] = None


def _build_soft_tone_wav() -> bytes:
    """Generate a short, low-frequency, low-amplitude cue as an in-memory WAV.

    `winsound.Beep` can only set frequency/duration (volume follows the
    system), and its default 880/660Hz double beep is piercing and carries
    across a room. This instead builds a soft ~C5 blip at low amplitude with
    a gentle attack/release envelope so it's noticeable up close but stays
    unobtrusive to people nearby.
    """
    framerate = 44100
    duration = 0.16
    freq = 523.25  # C5 — gentler than the old 880Hz
    amplitude = 0.16  # fraction of full scale; deliberately quiet
    n = int(framerate * duration)
    attack = int(framerate * 0.02)
    release = int(framerate * 0.04)
    buf = io.BytesIO()
    with wave.open(buf, 'wb') as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(framerate)
        frames = bytearray()
        for i in range(n):
            env = 1.0
            if i < attack:
                env = i / attack
            elif i > n - release:
                env = max(0.0, (n - i) / release)
            sample = int(amplitude * env * 32767 * math.sin(2 * math.pi * freq * i / framerate))
            frames += struct.pack('<h', sample)
        w.writeframes(bytes(frames))
    return buf.getvalue()


def _soft_tone_wav() -> bytes:
    global _SOFT_TONE_WAV
    if _SOFT_TONE_WAV is None:
        _SOFT_TONE_WAV = _build_soft_tone_wav()
    return _SOFT_TONE_WAV


def _beep_async() -> None:
    # Gate on the environment mode: only the 'home' profile makes sound; the
    # office/quiet profile stays silent so it never disturbs colleagues.
    try:
        if not PomodoroService.instance().sound_allowed():
            return
    except Exception:
        return

    def _play():
        try:
            import winsound

            # NOTE: winsound can't combine SND_MEMORY with SND_ASYNC ("Cannot
            # play asynchronously from memory"). We're already on a throwaway
            # daemon thread, so play synchronously — it just blocks this
            # thread for the tone's ~0.16s.
            winsound.PlaySound(_soft_tone_wav(), winsound.SND_MEMORY)
        except Exception:
            logging.exception("failed to play soft cue")

    threading.Thread(target=_play, daemon=True).start()


class PomodoroService:
    _instance: Optional["PomodoroService"] = None
    _lock = threading.Lock()

    @classmethod
    def instance(cls) -> "PomodoroService":
        with cls._lock:
            if cls._instance is None:
                cls._instance = PomodoroService()
            return cls._instance

    MODE_CLASSIC = 'classic'
    MODE_DEEP = 'deep'

    ENV_MODE_AUTO = 'auto'
    ENV_MODE_HOME = 'home'
    ENV_MODE_OFFICE = 'office'

    def __init__(self) -> None:
        self._phase = Phase.IDLE
        self._phase_before_pause: Optional[Phase] = None
        self._end_time: float = 0.0
        self._remaining_when_paused: float = 0.0
        self._completed_today = 0
        self._current_task_text = ""
        self._current_task_started_at: Optional[float] = None
        self._task_text_override: Optional[str] = None
        self._timer_thread: Optional[threading.Thread] = None
        self._stop_event = threading.Event()
        self._state_lock = threading.RLock()
        # Idle-monitor lives for the whole process lifetime; it self-checks
        # `_phase` on each tick so we don't need to start/stop it on phase
        # transitions.
        #   _idle_level: 0 = clear, 1 = soft blink, 2 = hard alarm.
        #   _idle_ack_needed: while True the hard alarm stays up until the
        #     user explicitly clicks "我回来了" — moving the mouse alone will
        #     NOT silence it (that's the whole point: make ignoring it cost).
        #   _current_phase_drifts: hard-alarm count within the current
        #     working phase; drives the "you actually need rest" forced break.
        self._idle_level = 0
        self._idle_ack_needed = False
        self._current_phase_drifts = 0
        # Restore last-used mode (classic / deep) so it persists across
        # launches, falling back to the kb.ini default_mode when the user has
        # never toggled it.
        try:
            default_mode = Config.pomodoro_default_mode()
        except Exception:
            logging.exception("failed to read default pomodoro mode from config")
            default_mode = self.MODE_CLASSIC
        self._mode: str = default_mode
        try:
            from shouyu.util.state import AppState

            self._mode = AppState.pomodoro_mode(default_mode)
        except Exception:
            logging.exception("failed to read pomodoro mode from state")

        threading.Thread(
            target=self._idle_monitor_loop,
            name="pomodoro-idle-monitor",
            daemon=True,
        ).start()

    # ---------- public API ----------

    def toggle(self) -> None:
        """Convenience binding for the global hotkey: idle->start, running->pause, paused->resume."""
        with self._state_lock:
            phase = self._phase
        if phase == Phase.IDLE:
            self.start_work()
        elif phase == Phase.PAUSED:
            self.resume()
        else:
            self.pause()

    def start_work(self, task_text: Optional[str] = None) -> None:
        if not Config.pomodoro_enabled():
            return
        if task_text:
            with self._state_lock:
                self._task_text_override = task_text
        self._begin_phase(Phase.WORKING)

    def start_planning(self, task_text: Optional[str] = None) -> None:
        """Start the morning planning session: a short block to lay out the
        day's tasks. Followed automatically by a short planning break."""
        if not Config.pomodoro_enabled():
            return
        with self._state_lock:
            self._task_text_override = task_text or "📋 规划今日最重要的 3 件事"
        self._begin_phase(Phase.PLANNING)

    def start_short_break(self) -> None:
        self._begin_phase(Phase.SHORT_BREAK)

    def start_long_break(self) -> None:
        self._begin_phase(Phase.LONG_BREAK)

    def pause(self) -> None:
        with self._state_lock:
            if self._phase in (Phase.IDLE, Phase.PAUSED):
                return
            self._phase_before_pause = self._phase
            self._remaining_when_paused = max(0.0, self._end_time - time.time())
            self._phase = Phase.PAUSED
            self._stop_event.set()
        self._emit("paused", "")

    def resume(self) -> None:
        with self._state_lock:
            if self._phase != Phase.PAUSED or self._phase_before_pause is None:
                return
            phase_to_resume = self._phase_before_pause
            self._phase = phase_to_resume
            self._end_time = time.time() + self._remaining_when_paused
            self._phase_before_pause = None
            self._stop_event.clear()
            self._spawn_timer_thread_locked()
        self._emit("resumed", "")

    def stop(self) -> None:
        with self._state_lock:
            self._phase = Phase.IDLE
            self._end_time = 0.0
            self._stop_event.set()
        self._emit("stopped", "")

    def extend_current_phase(self, minutes: int = 5) -> bool:
        """Add `minutes` to the current phase end time. Useful when in flow."""
        with self._state_lock:
            if self._phase not in (Phase.WORKING, Phase.SHORT_BREAK, Phase.LONG_BREAK):
                return False
            self._end_time += minutes * 60
        self._emit("extended", str(minutes))
        return True

    def mode(self) -> str:
        with self._state_lock:
            return self._mode

    def set_mode(self, mode: str) -> None:
        """Switch between classic (25/5) and deep (90/15) durations.

        Only takes effect on the next phase — we don't retroactively shrink
        a phase the user is already in.
        """
        if mode not in (self.MODE_CLASSIC, self.MODE_DEEP):
            mode = self.MODE_CLASSIC
        with self._state_lock:
            if self._mode == mode:
                return
            self._mode = mode
        try:
            from shouyu.util.state import AppState

            AppState.set_pomodoro_mode(mode)
        except Exception:
            logging.exception("failed to persist pomodoro mode")
        self._emit("mode_changed", mode)

    # ---------- environment mode (home / office / auto) ----------

    def env_mode(self) -> str:
        """The current env-mode *setting* the button reflects: auto/home/office.
        A manual home/office pick decays back to auto the next day."""
        try:
            from shouyu.util.state import AppState

            return AppState.env_mode_setting()
        except Exception:
            logging.exception("failed to read env mode")
            return self.ENV_MODE_AUTO

    def set_env_mode(self, mode: str) -> None:
        if mode not in (self.ENV_MODE_AUTO, self.ENV_MODE_HOME, self.ENV_MODE_OFFICE):
            mode = self.ENV_MODE_AUTO
        try:
            from shouyu.util.state import AppState

            AppState.set_env_mode_setting(mode)
        except Exception:
            logging.exception("failed to persist env mode")
        self._emit("env_mode_changed", mode)

    def resolved_env_mode(self) -> str:
        """Resolve the setting to the profile actually in effect right now:
        'home' or 'office'. Manual picks win; 'auto' follows the schedule."""
        mode = self.env_mode()
        if mode in (self.ENV_MODE_HOME, self.ENV_MODE_OFFICE):
            return mode
        return self._scheduled_env_mode()

    def _scheduled_env_mode(self) -> str:
        try:
            days = Config.pomodoro_quiet_days()
            start = Config.pomodoro_quiet_start()
            end = Config.pomodoro_quiet_end()
        except Exception:
            logging.exception("failed to read quiet schedule")
            return self.ENV_MODE_HOME
        if not start or not end or not days:
            return self.ENV_MODE_HOME
        now = time.localtime()
        if now.tm_wday not in days:
            return self.ENV_MODE_HOME
        now_minutes = now.tm_hour * 60 + now.tm_min
        start_minutes = start[0] * 60 + start[1]
        end_minutes = end[0] * 60 + end[1]
        if start_minutes <= now_minutes < end_minutes:
            return self.ENV_MODE_OFFICE
        return self.ENV_MODE_HOME

    def sound_allowed(self) -> bool:
        """Whether audible cues are allowed right now: only in the home
        profile, and only when notify_sound is on."""
        try:
            if not Config.pomodoro_notify_sound():
                return False
        except Exception:
            return False
        return self.resolved_env_mode() == self.ENV_MODE_HOME

    def acknowledge_idle(self) -> None:
        """Explicit "我回来了" acknowledgement of the hard idle alarm.

        This is the only thing that silences a level-2 alarm — moving the
        mouse won't. Clicking it counts as physically committing to come
        back, and (because the click itself is input) naturally resets the
        idle timer.
        """
        should_emit = False
        with self._state_lock:
            if self._idle_level > 0 or self._idle_ack_needed:
                self._idle_level = 0
                self._idle_ack_needed = False
                should_emit = True
        if should_emit:
            self._emit("idle_warning_off", "")

    def skip_break(self) -> bool:
        """Skip the current break and start a new work phase. Logged in stats."""
        with self._state_lock:
            if self._phase not in (Phase.SHORT_BREAK, Phase.LONG_BREAK):
                return False
        try:
            from shouyu.util.state import AppState

            AppState.increment_today_counter('breaks_skipped')
        except Exception:
            logging.exception("failed to log break skip")
        self._begin_phase(Phase.WORKING)
        return True

    def snapshot(self) -> dict:
        with self._state_lock:
            now = time.time()
            if self._phase == Phase.PAUSED:
                remaining = self._remaining_when_paused
            else:
                remaining = max(0.0, self._end_time - now)
            return {
                "phase": self._phase.value,
                "remaining_seconds": int(remaining),
                "task": self._current_task_text,
                "completed_today": self._completed_today,
            }

    # ---------- internals ----------

    def _begin_phase(self, phase: Phase, duration_override: Optional[int] = None) -> None:
        # Lunch guard: never let a focused (working / planning) phase run
        # during the configured lunch window. Convert it into a lunch break
        # that lasts until lunch ends; work/planning resumes afterwards.
        if phase in (Phase.WORKING, Phase.PLANNING):
            lunch_remaining = self._seconds_until_lunch_end()
            if lunch_remaining > 0:
                phase = Phase.LUNCH_BREAK
                duration_override = lunch_remaining
                with self._state_lock:
                    self._task_text_override = None

        duration = (
            duration_override if duration_override is not None else self._duration_for(phase)
        )
        if duration <= 0:
            return
        with self._state_lock:
            self._phase = phase
            self._end_time = time.time() + duration
            self._stop_event.set()
            self._stop_event = threading.Event()
            self._refresh_current_task_text_locked()
            if phase == Phase.WORKING:
                self._current_task_started_at = time.time()
                # Fresh focus block -> reset the drift budget.
                self._current_phase_drifts = 0
            self._spawn_timer_thread_locked()
        self._emit(
            "started",
            f"{phase.value}:{duration}:{self._current_task_text}",
        )

    def _seconds_until_lunch_end(self) -> int:
        """If we're currently inside the configured lunch window, return the
        seconds remaining until it ends; otherwise 0."""
        try:
            if not Config.pomodoro_lunch_enabled():
                return 0
            start = Config.pomodoro_lunch_start()
            end = Config.pomodoro_lunch_end()
        except Exception:
            logging.exception("failed to read lunch window config")
            return 0
        if not start or not end:
            return 0
        now = time.localtime()
        now_minutes = now.tm_hour * 60 + now.tm_min
        start_minutes = start[0] * 60 + start[1]
        end_minutes = end[0] * 60 + end[1]
        if start_minutes <= now_minutes < end_minutes:
            return (end_minutes - now_minutes) * 60 - now.tm_sec
        return 0

    def _duration_for(self, phase: Phase) -> int:
        deep = self._mode == self.MODE_DEEP
        if phase == Phase.PLANNING:
            return Config.pomodoro_planning_session_minutes() * 60
        if phase == Phase.WORKING:
            mins = Config.pomodoro_deep_work_minutes() if deep else Config.pomodoro_work_minutes()
            return mins * 60
        if phase == Phase.SHORT_BREAK:
            mins = (
                Config.pomodoro_deep_short_break_minutes()
                if deep
                else Config.pomodoro_short_break_minutes()
            )
            return mins * 60
        if phase == Phase.LONG_BREAK:
            mins = (
                Config.pomodoro_deep_long_break_minutes()
                if deep
                else Config.pomodoro_long_break_minutes()
            )
            return mins * 60
        return 0

    def _refresh_current_task_text_locked(self) -> None:
        if self._task_text_override:
            self._current_task_text = self._task_text_override
            self._task_text_override = None
            return

        from shouyu.service.excel import KbExcel

        try:
            excel = KbExcel()
            entry = excel.plan_service().current_in_progress_entry()
            if entry is not None:
                self._current_task_text = entry.text
            else:
                self._current_task_text = ""
        except Exception:
            logging.exception("failed to read current in-progress task")
            self._current_task_text = ""

    def _spawn_timer_thread_locked(self) -> None:
        thread = threading.Thread(target=self._run, name="pomodoro-timer", daemon=True)
        self._timer_thread = thread
        thread.start()

    def _run(self) -> None:
        stop_event = self._stop_event
        while not stop_event.is_set():
            with self._state_lock:
                phase = self._phase
                end_time = self._end_time
            if phase == Phase.PAUSED or phase == Phase.IDLE:
                return
            now = time.time()
            remaining = max(0.0, end_time - now)
            if remaining <= 0:
                self._on_phase_finished(phase)
                return
            self._emit("tick", str(int(remaining)))
            stop_event.wait(min(remaining, 1.0))

    def _on_phase_finished(self, finished_phase: Phase) -> None:
        _beep_async()
        if finished_phase == Phase.PLANNING:
            # Planning done -> short planning break, then normal work.
            self._begin_phase(
                Phase.SHORT_BREAK,
                duration_override=Config.pomodoro_planning_break_minutes() * 60,
            )
        elif finished_phase == Phase.WORKING:
            self._record_pomodoro_completion()
            with self._state_lock:
                self._completed_today += 1
                completed = self._completed_today
            cycles_long = Config.pomodoro_cycles_before_long_break()
            if cycles_long > 0 and completed % cycles_long == 0:
                self._begin_phase(Phase.LONG_BREAK)
            else:
                self._begin_phase(Phase.SHORT_BREAK)
        else:
            # short_break / long_break / lunch_break -> back to work
            self._begin_phase(Phase.WORKING)

    def _record_pomodoro_completion(self) -> None:
        from shouyu.service.excel import KbExcel

        try:
            now = time.time()
            duration_min = (
                Config.pomodoro_deep_work_minutes()
                if self._mode == self.MODE_DEEP
                else Config.pomodoro_work_minutes()
            )
            started_at = self._current_task_started_at or (now - duration_min * 60)
            label = (
                f"🍅 {duration_min}min @"
                f"{time.strftime('%H:%M', time.localtime(started_at))}-"
                f"{time.strftime('%H:%M', time.localtime(now))}"
            )
            KbExcel().append_detail(label)
        except Exception:
            logging.exception("failed to log pomodoro completion to Excel")

    def _emit(self, event: str, payload: str) -> None:
        try:
            from shouyu.view.qt_app import QtApp

            QtApp.emit_pomodoro_event(event, payload)
        except Exception:
            logging.exception("failed to emit pomodoro event")

    # ---------- idle monitor ----------

    _ALARM_BEEP_INTERVAL = 5.0  # seconds between nag-beeps while un-acknowledged

    def _clear_idle_state(self) -> None:
        """Reset any active idle warning/alarm and tell the UI to stop."""
        if self._idle_level > 0 or self._idle_ack_needed:
            self._idle_level = 0
            self._idle_ack_needed = False
            self._emit("idle_warning_off", "")

    def _enter_idle_alarm(self, idle: int) -> None:
        """Escalate to the hard alarm: count the drift, and either force a
        break (if you've drifted too many times this phase) or raise the
        loud, must-acknowledge alarm."""
        try:
            from shouyu.util.state import AppState

            AppState.increment_today_counter('focus_drifts')
        except Exception:
            logging.exception("idle monitor: failed to log focus drift")

        with self._state_lock:
            self._current_phase_drifts += 1
            phase_drifts = self._current_phase_drifts

        try:
            limit = Config.pomodoro_idle_drifts_before_break()
        except Exception:
            limit = 0
        if limit > 0 and phase_drifts >= limit:
            # You've drifted this many times in one block — you don't need
            # more nagging, you need rest. Force a short break.
            self._idle_level = 0
            self._idle_ack_needed = False
            self._emit("idle_warning_off", "")
            with self._state_lock:
                self._task_text_override = "🚑 反复走神，先休息一下再回来"
            self._begin_phase(Phase.SHORT_BREAK)
            return

        self._idle_level = 2
        self._idle_ack_needed = True
        self._emit("idle_alarm_on", str(int(idle)))

    def _idle_monitor_loop(self) -> None:
        """Watch global input idle time and escalate when the user drifts off
        during a working pomodoro.

        Two levels:
          * Level 1 (>= idle_warning_seconds): silent blink; clears by itself
            once the user moves again.
          * Level 2 (>= idle_alarm_seconds): a hard alarm that beeps every
            few seconds and stays up until an explicit "我回来了" click —
            moving the mouse alone will NOT silence it. Each alarm is counted
            as a drift; enough drifts in one phase forces a break.

        Self-contained: checks `_phase` every tick and gates everything on
        `Phase.WORKING`. Pause / stop / break phases clear any active
        warning, so the UI never gets stuck blinking across transitions.
        """
        from shouyu.util.idle import seconds_since_last_input

        last_beep = 0.0
        while True:
            time.sleep(1.0)
            try:
                warn_t = Config.pomodoro_idle_warning_seconds()
                alarm_t = Config.pomodoro_idle_alarm_seconds()
            except Exception:
                logging.exception("idle monitor: failed to read thresholds")
                warn_t, alarm_t = 0, 0

            with self._state_lock:
                phase = self._phase

            # Feature disabled, or not in a working phase -> clear and idle.
            if warn_t <= 0 or phase != Phase.WORKING:
                self._clear_idle_state()
                continue

            # A hard alarm is up and un-acknowledged: keep nagging with sound
            # regardless of whether the mouse has since moved. Only the
            # explicit acknowledge (or a phase change) clears it.
            if self._idle_ack_needed:
                now = time.time()
                if now - last_beep >= self._ALARM_BEEP_INTERVAL:
                    _beep_async()
                    last_beep = now
                continue

            try:
                idle = seconds_since_last_input()
            except Exception:
                logging.exception("idle monitor: failed to read idle time")
                continue

            # Level 2: escalate to the hard alarm.
            if alarm_t > 0 and idle >= alarm_t and self._idle_level < 2:
                self._enter_idle_alarm(int(idle))
                last_beep = time.time()
                continue

            # Level 1: soft blink, auto-clears when the user comes back.
            if self._idle_level == 0 and idle >= warn_t:
                self._idle_level = 1
                self._emit("idle_warning_on", str(int(idle)))
            elif self._idle_level == 1 and idle < warn_t:
                self._idle_level = 0
                self._emit("idle_warning_off", "")
