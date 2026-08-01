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
_wtsapi32 = None
_LASTINPUTINFO_cls = None
_WTSINFOEXW_cls = None

# WTSQuerySessionInformationW selectors / constants.
_WTS_CURRENT_SESSION = 0xFFFFFFFF  # (DWORD)-1
_WTSSessionInfoEx = 25             # WTS_INFO_CLASS.WTSSessionInfoEx
# WTSINFOEX_LEVEL1_W.SessionFlags values.
_WTS_SESSIONSTATE_UNKNOWN = 0xFFFFFFFF
_WTS_SESSIONSTATE_LOCK = 0
_WTS_SESSIONSTATE_UNLOCK = 1


def _lazy_init() -> bool:
    """Resolve ctypes bindings on first use. Returns True on success."""
    global _initialized, _user32, _kernel32, _wtsapi32
    global _LASTINPUTINFO_cls, _WTSINFOEXW_cls
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

        # --- WTS session info (reliable lock detection on Win8+) ------------
        LARGE_INTEGER = ctypes.c_longlong
        WCHAR = ctypes.c_wchar

        class _WTSINFOEX_LEVEL1_W(ctypes.Structure):
            _fields_ = [
                ("SessionId", wintypes.DWORD),
                ("SessionState", ctypes.c_int),  # WTS_CONNECTSTATE_CLASS
                ("SessionFlags", ctypes.c_long),  # LONG: 0=lock,1=unlock,-1=unknown
                ("WinStationName", WCHAR * 33),
                ("UserName", WCHAR * 21),
                ("DomainName", WCHAR * 18),
                ("LogonTime", LARGE_INTEGER),
                ("ConnectTime", LARGE_INTEGER),
                ("DisconnectTime", LARGE_INTEGER),
                ("LastInputTime", LARGE_INTEGER),
                ("CurrentTime", LARGE_INTEGER),
                ("IncomingBytes", wintypes.DWORD),
                ("OutgoingBytes", wintypes.DWORD),
                ("IncomingFrames", wintypes.DWORD),
                ("OutgoingFrames", wintypes.DWORD),
                ("IncomingCompressedBytes", wintypes.DWORD),
                ("OutgoingCompressedBytes", wintypes.DWORD),
            ]

        class _WTSINFOEX_LEVEL_W(ctypes.Union):
            _fields_ = [("WTSInfoExLevel1", _WTSINFOEX_LEVEL1_W)]

        class _WTSINFOEXW(ctypes.Structure):
            _fields_ = [
                ("Level", wintypes.DWORD),
                ("Data", _WTSINFOEX_LEVEL_W),
            ]

        _LASTINPUTINFO_cls = _LASTINPUTINFO
        _WTSINFOEXW_cls = _WTSINFOEXW
        _user32 = ctypes.windll.user32
        _kernel32 = ctypes.windll.kernel32
        # Pin signatures so 64-bit returns aren't truncated to 32-bit ints.
        _user32.GetLastInputInfo.argtypes = [ctypes.POINTER(_LASTINPUTINFO)]
        _user32.GetLastInputInfo.restype = wintypes.BOOL
        _kernel32.GetTickCount.restype = wintypes.DWORD
        # Fallback lock check. HDESK is a handle (pointer-sized), so pin the
        # return/arg types to avoid truncation on 64-bit.
        _user32.OpenInputDesktop.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        _user32.OpenInputDesktop.restype = wintypes.HDESK
        _user32.CloseDesktop.argtypes = [wintypes.HDESK]
        _user32.CloseDesktop.restype = wintypes.BOOL
        # Primary lock check via WTS. Not fatal if it fails to load.
        try:
            _wtsapi32 = ctypes.windll.wtsapi32
            _wtsapi32.WTSQuerySessionInformationW.argtypes = [
                wintypes.HANDLE,
                wintypes.DWORD,
                ctypes.c_int,
                ctypes.POINTER(ctypes.c_void_p),
                ctypes.POINTER(wintypes.DWORD),
            ]
            _wtsapi32.WTSQuerySessionInformationW.restype = wintypes.BOOL
            _wtsapi32.WTSFreeMemory.argtypes = [ctypes.c_void_p]
            _wtsapi32.WTSFreeMemory.restype = None
        except Exception:
            logging.exception("idle: wtsapi32 unavailable; using desktop fallback")
            _wtsapi32 = None
        return True
    except Exception:
        logging.exception("idle: failed to initialize Windows API bindings")
        _user32 = None
        return False


def _session_locked_via_wts():
    """Query the current session's lock state via WTSSessionInfoEx.

    Returns True/False when the API gives a definite answer, or None when the
    state is unknown / the call fails (so the caller can fall back).
    """
    if _wtsapi32 is None:
        return None
    import ctypes

    buf = ctypes.c_void_p()
    returned = None
    try:
        from ctypes import wintypes

        returned = wintypes.DWORD(0)
        ok = _wtsapi32.WTSQuerySessionInformationW(
            None,
            _WTS_CURRENT_SESSION,
            _WTSSessionInfoEx,
            ctypes.byref(buf),
            ctypes.byref(returned),
        )
        if not ok or not buf:
            return None
        info = ctypes.cast(buf, ctypes.POINTER(_WTSINFOEXW_cls)).contents
        flags = info.Data.WTSInfoExLevel1.SessionFlags & 0xFFFFFFFF
    except Exception:
        logging.exception("idle: WTSQuerySessionInformationW failed")
        return None
    finally:
        if buf:
            try:
                _wtsapi32.WTSFreeMemory(buf)
            except Exception:
                logging.exception("idle: WTSFreeMemory failed")

    if flags == _WTS_SESSIONSTATE_UNKNOWN:
        return None
    locked = flags == _WTS_SESSIONSTATE_LOCK
    # Windows 7 / Server 2008 R2 shipped the LOCK/UNLOCK flags reversed.
    try:
        ver = sys.getwindowsversion()
        if ver.major < 6 or (ver.major == 6 and ver.minor <= 1):
            locked = flags == _WTS_SESSIONSTATE_UNLOCK
    except Exception:
        pass
    return locked


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
    """Return True if the workstation is currently locked (lock screen / secure
    desktop / UAC prompt / secure screensaver).

    Primary path: WTSSessionInfoEx.SessionFlags — the reliable way to detect a
    lock on Windows 8/10/11. Falls back to OpenInputDesktop only when WTS can't
    give a definite answer (older Windows, RDP quirks), because the desktop
    trick misfires on modern Windows (it often still succeeds while locked).

    Returns False on non-Windows platforms or if everything fails — i.e. we
    default to "unlocked / allowed to make sound", so a detection failure can
    never leave the pomodoro permanently silent.
    """
    if not _lazy_init():
        return False

    wts = _session_locked_via_wts()
    if wts is not None:
        return wts

    # Fallback: try to open the *input* desktop. When locked, input switches to
    # the secure desktop which a normal-privilege process can't open (NULL).
    # DESKTOP_SWITCHDESKTOP is the cheapest access right for this probe.
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
