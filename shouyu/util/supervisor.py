"""Parent watchdog that restarts the shouyu worker after abnormal exits."""
from __future__ import annotations

import logging
import os
import subprocess
import sys
import time
from typing import List

from shouyu.config import Config
from shouyu.util.state import AppState


CHILD_ARGUMENT = "--shouyu-child"
SHUTDOWN_MARKER = "shouyu_shutdown.request"
MAX_RESTARTS_IN_WINDOW = 5
RESTART_WINDOW_SECONDS = 60


def shutdown_marker_path() -> str:
    return os.path.abspath(SHUTDOWN_MARKER)


def mark_shutdown_requested() -> None:
    try:
        with open(shutdown_marker_path(), "w", encoding="utf-8") as f:
            f.write(str(os.getpid()))
    except Exception:
        logging.exception("failed to write normal-shutdown marker")


def run() -> int:
    """Run the worker and restart it after a crash.

    A small restart limit prevents a broken installation from spinning at
    100% CPU forever. A normal tray Exit creates a marker and is never
    restarted.
    """
    _clear_shutdown_marker()
    restarts: List[float] = []
    while True:
        command = _child_command()
        started_at = time.time()
        # faulthandler appends to crash.log only while this worker runs, so
        # everything past this offset belongs to this worker session.
        crash_offset = _crash_log_size()
        logging.info("starting supervised worker: %s", command)
        child = None
        try:
            child = subprocess.Popen(command)
            code = child.wait()
        except Exception:
            logging.exception("supervisor failed to start or wait for worker")
            return 1

        if os.path.exists(shutdown_marker_path()):
            _clear_shutdown_marker()
            logging.info("worker exited normally by user request (code=%s)", code)
            return 0

        # A newer shouyu instance owns pid.txt. This supervisor was replaced
        # by a later launch and must not resurrect a competing worker.
        if _pid_file_belongs_to_another_process(child.pid):
            logging.info("another shouyu instance took ownership; supervisor exits")
            return 0

        now = time.time()
        restarts = [t for t in restarts if now - t < RESTART_WINDOW_SECONDS]
        restarts.append(now)
        logging.critical(
            "shouyu worker exited abnormally: code=%s, runtime=%.1fs",
            code,
            now - started_at,
        )
        _archive_native_crash(crash_offset, code, now - started_at, now)
        _record_crash_notice(code)
        if len(restarts) > MAX_RESTARTS_IN_WINDOW:
            logging.critical(
                "worker crashed %s times in %ss; automatic restart paused",
                len(restarts),
                RESTART_WINDOW_SECONDS,
            )
            return 1
        time.sleep(0.5)


def _child_command() -> List[str]:
    if getattr(sys, "frozen", False):
        return [sys.executable, CHILD_ARGUMENT]
    return [sys.executable, os.path.abspath(sys.argv[0]), CHILD_ARGUMENT]


def _clear_shutdown_marker() -> None:
    try:
        os.remove(shutdown_marker_path())
    except FileNotFoundError:
        pass
    except OSError:
        logging.exception("failed to clear normal-shutdown marker")


def _pid_file_belongs_to_another_process(child_pid: int) -> bool:
    try:
        with open("pid.txt", "r", encoding="utf-8") as f:
            owner = int(f.read().strip())
        return owner not in (0, child_pid)
    except (FileNotFoundError, ValueError):
        return False
    except OSError:
        logging.exception("failed to inspect pid.txt after worker exit")
        return False


def _crash_log_size() -> int:
    try:
        return os.path.getsize(Config.crash_log_path())
    except OSError:
        return 0


def _archive_native_crash(offset: int, code: int, runtime: float, when: float) -> None:
    """Timestamp this worker's native crash dump and mirror it into kb.log.

    faulthandler writes its dump to crash.log with no timestamp. The supervisor
    knows exactly when the worker died, so it prepends a timestamped header to
    the slice of crash.log written by that worker, and logs the same dump into
    kb.log, which is the single entry point for all logs.
    """
    path = Config.crash_log_path()
    stamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(when))
    header = (
        f"\n=== [{stamp}] worker crashed: exit code {code} "
        f"(0x{code & 0xFFFFFFFF:08X}), runtime {runtime:.1f}s ===\n"
    )
    try:
        with open(path, "r+b") as f:
            f.seek(0, os.SEEK_END)
            start = min(offset, f.tell())  # file may have been truncated meanwhile
            f.seek(start)
            raw = f.read()
            # Keep this worker's "enabled" marker first, then the crash header,
            # then the faulthandler dump itself.
            marker = raw.find(b"crash diagnostics enabled")
            split = raw.find(b"\n", marker) + 1 if marker >= 0 else 0
            stamped = raw[:split] + header.encode("utf-8") + raw[split:]
            f.seek(start)
            f.write(stamped)
    except OSError:
        logging.exception("failed to timestamp native crash dump in %s", path)
        return
    logging.critical(
        "native crash dump (also stamped in %s):\n%s",
        path,
        stamped.decode("utf-8", errors="replace").strip() or "<no faulthandler output>",
    )


def _record_crash_notice(code: int) -> None:
    try:
        AppState.set(
            "crash_notice",
            {
                "time": time.strftime("%Y-%m-%d %H:%M:%S"),
                "exit_code": code,
                "log_path": Config.log_path(),
                "crash_log_path": Config.crash_log_path(),
            },
        )
    except Exception:
        logging.exception("failed to persist crash notice")
