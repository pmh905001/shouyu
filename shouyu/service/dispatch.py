"""Background consumer that drains `message_queue` into kb.xlsx.

Design reference: docs/excel-save-resilience.md (§4, §8). Every write path
(habit_dialog / todo_panel plan-save, clipboard/screenshot save) enqueues a
message and returns immediately; this module is the single place that
actually opens Excel and retries when it's locked by another program.

One drain cycle:
  1. Pull the next batch of pending messages (after this consumer's cursor).
  2. Stage each one into a single shared `KbExcel` instance. A message whose
     payload can't even be staged (a bug, not a lock issue) is dead-lettered
     immediately so it doesn't block everything behind it.
  3. Save once for the whole batch.
     - PermissionError (locked) -> bump attempts, leave pending, retry later.
     - anything else -> same retry bookkeeping, plus a rate-limited blocking
       popup, since this is not the "just wait for Excel to close" case.
     - success -> mark done, advance the cursor.
"""
from __future__ import annotations

import io
import logging
import os
import tempfile
import threading
import time
from typing import List, Optional

from shouyu.config import Config
from shouyu.service import message_queue

CONSUMER = 'excel_dispatch'
DRAIN_INTERVAL_SECONDS = 20
BATCH_LIMIT = 200
_STUCK_ALERT_RATE_LIMIT_SECONDS = 3600  # at most once/hour, per docs §4.4
_BROKEN_SAVE_ALERT_RATE_LIMIT_SECONDS = 3600

_kick_event = threading.Event()
_start_lock = threading.Lock()
_started = False
_last_stuck_alert_at = 0.0
_last_broken_alert_at = 0.0


def start() -> None:
    """Idempotent: safe to call once from main.py's startup path."""
    global _started
    with _start_lock:
        if _started:
            return
        _started = True
    message_queue.init_db()
    threading.Thread(target=_run_loop, name='shouyu-queue-dispatch', daemon=True).start()


def kick() -> None:
    """Wake the dispatcher immediately instead of waiting for the next
    periodic tick. Safe to call from any thread."""
    _kick_event.set()


def notify_recorded(title: str, msg: str) -> None:
    """Fired by callers right after `enqueue()` - this IS the fast feedback;
    the dispatcher itself stays quiet on the ordinary happy path so a normal
    save doesn't double-toast."""
    try:
        from shouyu.view.msgbox import MessageBox, MessageType

        MessageBox.pop_up_message(title=title, msg=msg, level=MessageType.SUCCESS)
    except Exception:
        logging.exception('failed to show "recorded" toast')


def _run_loop() -> None:
    while True:
        try:
            _drain_once()
        except Exception:
            logging.exception('dispatch drain cycle crashed')
        _kick_event.wait(timeout=DRAIN_INTERVAL_SECONDS)
        _kick_event.clear()


def _temp_image_path(message_id: int) -> str:
    return os.path.join(tempfile.gettempdir(), f'shouyu_queue_img_{message_id}.png')


def _dispatch_plan_save(kb_excel, payload: dict, attachment_bytes: Optional[bytes], image_path: str) -> None:
    from shouyu.service.plan import PlanTask, TaskCategory, TaskStatus

    tasks = [PlanTask.from_dict(d) for d in (payload.get('tasks') or [])]
    non_empty = [t for t in tasks if (t.text or '').strip()]
    plan = kb_excel.plan_service()
    plan.write_plan_tasks(non_empty)

    original_in_progress = payload.get('original_in_progress')
    new_in_progress = next((t for t in non_empty if t.status == TaskStatus.IN_PROGRESS), None)
    if new_in_progress is not None and new_in_progress.text != original_in_progress:
        plan.switch_in_progress(new_in_progress)

    yesterday_done_rows = payload.get('yesterday_done_rows') or []
    yesterday_date = payload.get('yesterday_date')
    if yesterday_done_rows and yesterday_date:
        yesterday_plan = kb_excel.plan_service_for(yesterday_date)
        if yesterday_plan is not None:
            for row in yesterday_done_rows:
                if row:
                    yesterday_plan.mark_plan_done(row)

    work_tasks = payload.get('work_tasks')
    if work_tasks is not None:
        kb_excel.backlog_service(TaskCategory.WORK).write(
            [PlanTask.from_dict(d) for d in work_tasks if (d.get('text') or '').strip()]
        )
    life_tasks = payload.get('life_tasks')
    if life_tasks is not None:
        kb_excel.backlog_service(TaskCategory.LIFE).write(
            [PlanTask.from_dict(d) for d in life_tasks if (d.get('text') or '').strip()]
        )
    kb_excel.mark_changed()


def _dispatch_clipboard_append(kb_excel, payload: dict, attachment_bytes: Optional[bytes], image_path: str) -> None:
    from PIL import Image as PILImageModule

    column = payload.get('column')
    if attachment_bytes is not None:
        data = PILImageModule.open(io.BytesIO(attachment_bytes))
        kb_excel.append_one_record(data, target_column=column, image_path=image_path)
    else:
        kb_excel.append_one_record(payload.get('text'), target_column=column)


_DISPATCHERS = {
    'plan_save': _dispatch_plan_save,
    'clipboard_append': _dispatch_clipboard_append,
}


def _stage_message(kb_excel, message: 'message_queue.QueueMessage', image_path: str) -> None:
    dispatcher = _DISPATCHERS.get(message.kind)
    if dispatcher is None:
        raise ValueError(f'unknown queue kind: {message.kind}')
    attachment_bytes = None
    if message.attachment_id is not None:
        attachment_bytes = message_queue.get_attachment_bytes(message.attachment_id)
    dispatcher(kb_excel, message.payload, attachment_bytes, image_path)


def _drain_once() -> None:
    batch = message_queue.fetch_pending_batch(CONSUMER, limit=BATCH_LIMIT)
    if not batch:
        _maybe_alert_stuck()
        return

    message_queue.mark_processing([m.id for m in batch])

    from shouyu.service.excel import KbExcel

    kb_excel = KbExcel()
    staged: List['message_queue.QueueMessage'] = []
    temp_paths: List[str] = []
    for m in batch:
        image_path = _temp_image_path(m.id)
        try:
            _stage_message(kb_excel, m, image_path)
            staged.append(m)
            if os.path.exists(image_path):
                temp_paths.append(image_path)
        except Exception as e:
            logging.exception(f'failed to stage queue message #{m.id} ({m.kind}); dead-lettering it')
            message_queue.mark_dead(m.id, f'stage error: {e}')

    try:
        if staged:
            kb_excel.force_save()
        message_queue.mark_done([m.id for m in staged])
        retried = [m for m in staged if m.attempts > 0]
        if retried:
            _notify_recovered(len(retried))
        _clear_lock_alert_state()
    except PermissionError as e:
        logging.warning(f'batch save locked ({len(staged)} messages), will retry next cycle: {e}')
        message_queue.mark_pending_retry([m.id for m in staged], str(e))
    except Exception as e:
        logging.exception('batch save failed with a non-lock error')
        message_queue.mark_pending_retry([m.id for m in staged], str(e))
        _maybe_alert_broken_save(kb_excel, e)
    finally:
        message_queue.advance_cursor(CONSUMER)
        for p in temp_paths:
            _safe_remove_temp(p)

    _maybe_alert_stuck()


def _safe_remove_temp(path: str) -> None:
    try:
        if os.path.exists(path):
            os.remove(path)
    except Exception:
        logging.exception(f'failed to remove temp image: {path}')


def _maybe_alert_stuck() -> None:
    global _last_stuck_alert_at
    info = message_queue.stuck_info(CONSUMER, Config.stuck_alert_minutes())
    if info is None:
        return
    now = time.time()
    if now - _last_stuck_alert_at < _STUCK_ALERT_RATE_LIMIT_SECONDS:
        return
    _last_stuck_alert_at = now
    try:
        from shouyu.view.msgbox import MessageBox, MessageType

        MessageBox.pop_up_message(
            title='保存排队中',
            msg=f"有 {info['count']} 条还没同步到 Excel，可能是文件被占用；会持续在后台重试，不会丢失。",
            level=MessageType.ERROR,
        )
    except Exception:
        logging.exception('failed to show stuck-queue toast')


def _clear_lock_alert_state() -> None:
    global _last_stuck_alert_at
    _last_stuck_alert_at = 0.0


def _notify_recovered(count: int) -> None:
    try:
        from shouyu.view.msgbox import MessageBox, MessageType

        MessageBox.pop_up_message(
            title='已同步',
            msg=f'{count} 条之前卡住的记录已成功写入 Excel',
            level=MessageType.SUCCESS,
        )
    except Exception:
        logging.exception('failed to show recovered toast')


def _maybe_alert_broken_save(kb_excel, err: Exception) -> None:
    global _last_broken_alert_at
    now = time.time()
    if now - _last_broken_alert_at < _BROKEN_SAVE_ALERT_RATE_LIMIT_SECONDS:
        return
    _last_broken_alert_at = now
    preserved = ''
    try:
        preserved = kb_excel.preserve_unsaved() or ''
    except Exception:
        logging.exception('failed to preserve unsaved changes after a non-lock save error')
    msg = f'保存失败：{err}'
    if preserved:
        msg += f'\n\n你的改动已经写入备用文件，不会丢失：\n{preserved}'
    else:
        msg += '\n\n（备用文件也写入失败；请手动核对最近的备份。）'
    try:
        from shouyu.view.qt_app import QtApp

        QtApp.show_save_status('error', '保存失败', msg)
    except Exception:
        logging.exception('failed to show blocking save-failure dialog')
