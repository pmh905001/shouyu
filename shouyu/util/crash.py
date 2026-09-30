"""Process-level crash diagnostics for the desktop application."""
from __future__ import annotations

import faulthandler
import logging
import os
import sys
import threading
from types import TracebackType
from typing import Optional, Type

from shouyu.config import Config


_fault_log = None


def install() -> None:
    """Install traceback and native-crash diagnostics once per process."""
    global _fault_log
    if _fault_log is not None:
        return

    crash_path = Config.crash_log_path()
    try:
        os.makedirs(os.path.dirname(crash_path) or ".", exist_ok=True)
        _fault_log = open(crash_path, "a", encoding="utf-8", buffering=1)
        faulthandler.enable(file=_fault_log, all_threads=True)
        _fault_log.write("\n--- crash diagnostics enabled ---\n")
    except Exception:
        logging.exception("failed to enable faulthandler")

    sys.excepthook = _unhandled_exception
    threading.excepthook = _unhandled_thread_exception


def _unhandled_exception(
    exc_type: Type[BaseException],
    exc_value: BaseException,
    traceback: Optional[TracebackType],
) -> None:
    logging.critical(
        "uncaught exception in main thread",
        exc_info=(exc_type, exc_value, traceback),
    )


def _unhandled_thread_exception(args: threading.ExceptHookArgs) -> None:
    logging.critical(
        "uncaught exception in thread %s",
        args.thread.name if args.thread else "<unknown>",
        exc_info=(args.exc_type, args.exc_value, args.exc_traceback),
    )
