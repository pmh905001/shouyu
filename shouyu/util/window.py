import ctypes
import logging
import threading
from contextlib import contextmanager
from ctypes import wintypes

user32 = ctypes.windll.user32
dwmapi = ctypes.windll.dwmapi

_DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = ctypes.c_void_p(-4)
_MONITORINFOF_PRIMARY = 0x1
_DWMWA_EXTENDED_FRAME_BOUNDS = 9
_SW_RESTORE = 9
_SW_MAXIMIZE = 3
_SWP_NOZORDER = 0x0004
_SWP_NOACTIVATE = 0x0010

user32.SetThreadDpiAwarenessContext.restype = ctypes.c_void_p
user32.SetThreadDpiAwarenessContext.argtypes = [ctypes.c_void_p]


class _MONITORINFO(ctypes.Structure):
    _fields_ = [
        ('cbSize', wintypes.DWORD),
        ('rcMonitor', wintypes.RECT),
        ('rcWork', wintypes.RECT),
        ('dwFlags', wintypes.DWORD),
    ]


_MonitorEnumProc = ctypes.WINFUNCTYPE(
    wintypes.BOOL, wintypes.HMONITOR, wintypes.HDC, ctypes.POINTER(wintypes.RECT), wintypes.LPARAM
)


@contextmanager
def _physical_pixels():
    """Mixed-DPI monitors only share one coordinate space when the calling
    thread is per-monitor DPI aware; otherwise Windows hands back scaled
    rectangles and the window lands in the wrong place."""
    old = user32.SetThreadDpiAwarenessContext(_DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2)
    try:
        yield
    finally:
        if old:
            user32.SetThreadDpiAwarenessContext(old)


class WindowManager:
    # hwnd -> (left, top, width, height, was_maximized) before spanning
    _saved = {}
    _lock = threading.Lock()

    @staticmethod
    def _monitors():
        """[(work_area, is_primary, monitor_area)] for every connected monitor."""
        result = []

        def callback(hmonitor, _hdc, _rect, _data):
            info = _MONITORINFO()
            info.cbSize = ctypes.sizeof(_MONITORINFO)
            if user32.GetMonitorInfoW(hmonitor, ctypes.byref(info)):
                result.append((info.rcWork, bool(info.dwFlags & _MONITORINFOF_PRIMARY), info.rcMonitor))
            return True

        user32.EnumDisplayMonitors(None, None, _MonitorEnumProc(callback), 0)
        return result

    @staticmethod
    def _common_span(work_areas):
        """Rectangle for a window that must fit on every given work area:
        horizontally the union (the screens sit side by side), vertically the
        intersection, so the height is the shortest screen's. Returns
        (left, top, right, bottom), or None when the screens share no
        vertical range."""
        left = min(w.left for w in work_areas)
        right = max(w.right for w in work_areas)
        top = max(w.top for w in work_areas)
        bottom = min(w.bottom for w in work_areas)
        if bottom <= top:
            return None
        return left, top, right, bottom

    @classmethod
    def _span_area(cls):
        """Span area for the monitors above the primary one (every non-primary
        monitor if none sits above it). See _common_span for the geometry."""
        monitors = cls._monitors()
        primary = next((m for m in monitors if m[1]), None)
        if primary is None:
            return None
        others = [m for m in monitors if not m[1]]
        above = [m for m in others if m[2].top < primary[2].top]
        targets = above or others
        if not targets:
            return None
        return cls._common_span([m[0] for m in targets])

    @staticmethod
    def _invisible_border(hwnd):
        """Win10/11 windows carry an invisible resize border; return how far
        the real window rect extends past the visible frame on each side."""
        rect = wintypes.RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
        frame = wintypes.RECT()
        hr = dwmapi.DwmGetWindowAttribute(
            hwnd, _DWMWA_EXTENDED_FRAME_BOUNDS, ctypes.byref(frame), ctypes.sizeof(frame)
        )
        if hr != 0:
            return 0, 0, 0, 0
        return frame.left - rect.left, frame.top - rect.top, rect.right - frame.right, rect.bottom - frame.bottom

    @staticmethod
    def _move(hwnd, left, top, width, height):
        # Crossing into a monitor with a different scale makes the app resize
        # itself on WM_DPICHANGED; the second call wins.
        for _ in range(2):
            user32.SetWindowPos(hwnd, None, left, top, width, height, _SWP_NOZORDER | _SWP_NOACTIVATE)

    @classmethod
    def toggle_span_monitors(cls, hotkey: str = '') -> None:
        """Stretch the foreground window across the monitors above the
        primary one; run again on the same window to put it back."""
        try:
            with cls._lock, _physical_pixels():
                hwnd = user32.GetForegroundWindow()
                if not hwnd:
                    return

                saved = cls._saved.pop(hwnd, None)
                if saved is not None:
                    left, top, width, height, was_max = saved
                    cls._move(hwnd, left, top, width, height)
                    if was_max:
                        user32.ShowWindow(hwnd, _SW_MAXIMIZE)
                    logging.info(f'{hotkey}: restored window hwnd={hwnd}')
                    return

                area = cls._span_area()
                if area is None:
                    logging.info(f'{hotkey}: no span area (no secondary monitor, or no shared height)')
                    return

                was_max = bool(user32.IsZoomed(hwnd))
                if was_max:
                    user32.ShowWindow(hwnd, _SW_RESTORE)
                rect = wintypes.RECT()
                user32.GetWindowRect(hwnd, ctypes.byref(rect))
                cls._saved[hwnd] = (rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top, was_max)

                left, top, right, bottom = area
                cls._move(hwnd, left, top, right - left, bottom - top)
                # The invisible border scales with DPI, so measure it on the target monitor.
                bl, bt, br, bb = cls._invisible_border(hwnd)
                cls._move(hwnd, left - bl, top - bt, right - left + bl + br, bottom - top + bt + bb)
                logging.info(f'{hotkey}: spanned window hwnd={hwnd} to {area}')
        except Exception:
            logging.exception('failed to span window across monitors')
