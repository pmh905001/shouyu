"""System-wide idle detection on Windows.

We use the kernel's `GetLastInputInfo` API instead of running our own
mouse/keyboard listener, for three reasons:

  1. Zero new dependencies — pure ctypes, talking to user32.dll directly.
  2. Zero polling overhead — the OS already tracks "last input time" for
     us; we just read it once per second.
  3. No conflict with the global `keyboard` hook this project already
     uses for hotkeys (adding a second hook tends to break the first).

`GetLastInputInfo` reports the timestamp of the last input event from
ANY device (mouse + keyboard + touch + pen). That matches our actual
intent for the pomodoro idle-warning feature: "user is drifting off"
means *neither* mouse nor keyboard activity, not just an idle mouse —
otherwise typing-heavy work (coding, writing) would falsely trigger.

This module is Windows-only. On other platforms `seconds_since_last_input`
returns 0.0 (i.e. "always active"), which gracefully disables the
feature rather than crashing.
"""
from __future__ import annotations

import logging
import sys

_initialized = False
_user32 = None
_kernel32 = None
_LASTINPUTINFO_cls = None


def _lazy_init() -> bool:
    """Resolve ctypes bindings on first use. Returns True on success."""
    global _initialized, _user32, _kernel32, _LASTINPUTINFO_cls
    if _initialized:
        return _user32 is not None
    _initialized = True
    if not sys.platform.startswith("win"):
        return False
    try:
        import ctypes
        from ctypes import wintypes

        class _LASTINPUTINFO(ctypes.Structure):
            _fields_ = [
                ("cbSize", wintypes.UINT),
                ("dwTime", wintypes.DWORD),
            ]

        _LASTINPUTINFO_cls = _LASTINPUTINFO
        _user32 = ctypes.windll.user32
        _kernel32 = ctypes.windll.kernel32
        # Pin signatures so 64-bit returns aren't truncated to 32-bit ints.
        _user32.GetLastInputInfo.argtypes = [ctypes.POINTER(_LASTINPUTINFO)]
        _user32.GetLastInputInfo.restype = wintypes.BOOL
        _kernel32.GetTickCount.restype = wintypes.DWORD
        # Used by is_workstation_locked(). HDESK is a handle (pointer-sized),
        # so pin the return/arg types to avoid truncation on 64-bit.
        _user32.OpenInputDesktop.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        _user32.OpenInputDesktop.restype = wintypes.HDESK
        _user32.CloseDesktop.argtypes = [wintypes.HDESK]
        _user32.CloseDesktop.restype = wintypes.BOOL
        return True
    except Exception:
        logging.exception("idle: failed to initialize Windows API bindings")
        _user32 = None
        return False


def seconds_since_last_input() -> float:
    """Return seconds elapsed since the last user input (mouse OR keyboard).

    Returns 0.0 on non-Windows platforms or if the API call fails — i.e.
    callers see "user just acted", which safely disables any idle-based
    behavior on unsupported platforms.

    Note: GetTickCount wraps every ~49.7 days. If both `tick_now` and
    `dwTime` straddle the wrap, ctypes' DWORD subtraction in 32-bit
    arithmetic still yields the correct elapsed delta (modulo 2**32),
    so we don't need to special-case it.
    """
    import ctypes

    if not _lazy_init():
        return 0.0
    info = _LASTINPUTINFO_cls()
    info.cbSize = ctypes.sizeof(info)
    if not _user32.GetLastInputInfo(ctypes.byref(info)):
        return 0.0
    tick_now = _kernel32.GetTickCount()
    delta_ms = (tick_now - info.dwTime) & 0xFFFFFFFF
    return delta_ms / 1000.0


def is_workstation_locked() -> bool:
    """Return True if the workstation is currently locked (or otherwise on the
    secure Winlogon desktop, e.g. UAC prompt / lock screen / secure screensaver).

    Implementation: try to open the *input* desktop. When the session is
    locked, the interactive input desktop switches to the secure desktop and a
    normal-privilege process can no longer open it, so OpenInputDesktop returns
    NULL. We treat that NULL as "locked".

    Returns False on non-Windows platforms or if the API call fails — i.e. we
    default to "unlocked / allowed to make sound", so a detection failure can
    never leave the pomodoro permanently silent.
    """
    if not _lazy_init():
        return False
    # DESKTOP_SWITCHDESKTOP — cheapest access right that still forces the
    # "can I actually reach the input desktop?" check we care about.
    DESKTOP_SWITCHDESKTOP = 0x0100
    try:
        handle = _user32.OpenInputDesktop(0, False, DESKTOP_SWITCHDESKTOP)
    except Exception:
        logging.exception("idle: OpenInputDesktop call failed")
        return False
    if not handle:
        return True
    try:
        _user32.CloseDesktop(handle)
    except Exception:
        logging.exception("idle: CloseDesktop call failed")
    return False
