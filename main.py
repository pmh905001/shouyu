import http.server
import json
import logging
import socketserver
import sys
import threading

import keyboard

from shouyu.action.shortcut import Shortcut
from shouyu.config import Config
from shouyu.log import Log
from shouyu.service.excel import KbExcel
from shouyu.util.package import Package
from shouyu.util.process import ProcessManager
from shouyu.util.state import AppState
from shouyu.view.msgbox import MessageBox, MessageType
from shouyu.view.qt_app import QtApp
from shouyu.view.tray import Tray


HTTP_HOST = "127.0.0.1"
HTTP_PORT = 19823


class TitleRequestHandler(http.server.BaseHTTPRequestHandler):
    def do_POST(self):
        if self.path == "/add-title":
            content_length = int(self.headers["Content-Length"])
            post_data = self.rfile.read(content_length)
            try:
                data = json.loads(post_data.decode("utf-8"))
                title = data.get("title", "").strip()
                if title:
                    # 使用队列执行，避免阻塞 HTTP 服务
                    def do_write():
                        KbExcel.append_title_to_next_row(title)
                    Shortcut.executor.add(do_write, ())
                    self.send_response(200)
                    self.send_header("Content-Type", "application/json")
                    self.end_headers()
                    self.wfile.write(json.dumps({"status": "ok", "message": "queued"}).encode())
                else:
                    self.send_response(400)
                    self.send_header("Content-Type", "application/json")
                    self.end_headers()
                    self.wfile.write(json.dumps({"status": "error", "message": "title required"}).encode())
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"status": "error", "message": str(e)}).encode())
        else:
            self.send_response(404)
            self.end_headers()

    def log_message(self, format, *args):
        logging.info(f"[HTTP] {args[0]}")


def _run_http_server(port=HTTP_PORT):
    with socketserver.TCPServer(("", port), TitleRequestHandler) as httpd:
        logging.info(f"[HTTP] Server started on port {port}")
        httpd.serve_forever()


def _show_habit_dialog_if_needed():
    """Pop up the habit reminder dialog at most once per day (per user)."""
    habits = Config.habits()
    if not habits:
        return
    today = AppState.today_str()
    if AppState.get('last_habit_shown_date') == today:
        logging.info('habit dialog already shown today, skip')
        return
    AppState.set('last_habit_shown_date', today)
    QtApp.show_habit_dialog(habits)


def _alert_if_excel_was_recovered():
    """If the canonical Excel was corrupt at startup and we auto-recovered
    from a backup, surface that to the user immediately so they can pick a
    different backup if they don't trust the one we picked."""
    try:
        excel = KbExcel()
        if excel.recovered_from_backup:
            logging.warning(
                f"main Excel was corrupt at startup; auto-recovered from "
                f"{excel.recovered_from_backup}"
            )
            QtApp.show_backup_restore(excel.recovered_from_backup)
    except Exception:
        logging.exception("failed to check Excel recovery status")


def _start_pomodoro_if_enabled():
    if not Config.pomodoro_enabled():
        return
    from shouyu.service.pomodoro import PomodoroService

    PomodoroService.instance().start_work()


def _is_daemon_up() -> bool:
    """Check whether a shouyu daemon is already listening on HTTP_PORT."""
    import socket

    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(0.5)
    try:
        return sock.connect_ex((HTTP_HOST, HTTP_PORT)) == 0
    except Exception:
        return False
    finally:
        sock.close()


def _cli_add_title(title: str) -> int:
    """One-shot CLI: add `title` as a new row, then return.

    Triggered when shouyu.exe is invoked with positional args (e.g. via
    Win+R: `shouyu hello`). Stays out of the daemon's way:

      * NEVER calls `kill_old_process` — that would kill the running
        daemon and orphan the tray / hotkeys / Qt session.
      * NEVER starts the tray, Qt, hotkeys, or the local HTTP server.

    Forwards to the running daemon when possible so behavior matches the
    hotkey path (queue, retries, backups). Falls back to writing Excel
    directly when the daemon is offline.

    Stays silent on success — Win+R users want fire-and-forget. Only
    surfaces a Windows toast when something fails.

    Returns: process exit code (0 == success).
    """
    import http.client

    if _is_daemon_up():
        try:
            payload = json.dumps({"title": title}).encode("utf-8")
            conn = http.client.HTTPConnection(HTTP_HOST, HTTP_PORT, timeout=5)
            try:
                conn.request(
                    "POST",
                    "/add-title",
                    body=payload,
                    headers={"Content-Type": "application/json"},
                )
                resp = conn.getresponse()
                if resp.status == 200:
                    logging.info(f"CLI: queued via daemon: {title!r}")
                    return 0
                body = resp.read().decode("utf-8", errors="replace")
                logging.error(f"daemon rejected (HTTP {resp.status}): {body}")
                MessageBox.pop_up_message(
                    "添加失败",
                    f"daemon 拒绝：HTTP {resp.status}",
                    level=MessageType.ERROR,
                )
                return 2
            finally:
                conn.close()
        except Exception as e:
            logging.exception("failed to forward to daemon")
            MessageBox.pop_up_message(
                "添加失败",
                f"无法连接 daemon：{e}",
                level=MessageType.ERROR,
            )
            return 3

    # Daemon offline — write straight to Excel.
    try:
        KbExcel.append_title_to_next_row(title)
        logging.info(f"CLI: wrote directly (daemon offline): {title!r}")
        return 0
    except Exception as e:
        logging.exception("direct Excel write failed")
        MessageBox.pop_up_message(
            "添加失败",
            f"写入失败：{e}",
            level=MessageType.ERROR,
        )
        return 4


def _run_daemon():
    """Original startup path — long-running tray + Qt + hotkeys + HTTP."""
    ProcessManager.kill_old_process()
    logging.info('Started service!')
    tray = Tray.create()
    threading.Thread(target=tray.run, daemon=True).start()

    # 启动本地 HTTP 服务，用于接收 shouyu.exe 的 CLI 调用 / run.py 的请求
    threading.Thread(target=_run_http_server, daemon=True).start()

    QtApp.start()
    Shortcut.start()
    Shortcut.register_hot_keys()

    # 在 Qt 就绪后弹出当日习惯提醒；如果开启了番茄工作法也一并启动。
    # 启动健康检查放在最前 — 如果主 Excel 损坏，先让用户决定如何恢复。
    threading.Timer(0.5, _alert_if_excel_was_recovered).start()
    threading.Timer(1.0, _show_habit_dialog_if_needed).start()
    threading.Timer(2.0, _start_pomodoro_if_enabled).start()

    keyboard.wait()


if __name__ == '__main__':
    Package.set_cwd()
    Log.setup()

    # CLI mode: any positional arg means "add this as a new row, then exit".
    # Triggered by Win+R / cmd: `shouyu 我的任务`. Must run BEFORE any
    # daemon-side init (especially kill_old_process) so we don't disturb
    # the long-running shouyu.exe instance.
    cli_args = [a for a in sys.argv[1:] if a.strip()]
    if cli_args:
        title = " ".join(cli_args).strip()
        sys.exit(_cli_add_title(title) if title else 0)

    _run_daemon()
