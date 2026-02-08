import http.server
import json
import logging
import socketserver
import threading

import keyboard

from shouyu.action.shortcut import Shortcut
from shouyu.log import Log
from shouyu.service.excel import KbExcel
from shouyu.util.package import Package
from shouyu.util.process import ProcessManager
from shouyu.view.msgbox import MessageBox
from shouyu.view.tray import Tray


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


def _run_http_server(port=19823):
    with socketserver.TCPServer(("", port), TitleRequestHandler) as httpd:
        logging.info(f"[HTTP] Server started on port {port}")
        httpd.serve_forever()


if __name__ == '__main__':
    Package.set_cwd()
    Log.setup()
    ProcessManager.kill_old_process()
    logging.info('Started service!')
    tray = Tray.create()
    threading.Thread(target=tray.run, daemon=True).start()

    # 启动本地 HTTP 服务，用于接收 run.py 的请求
    threading.Thread(target=_run_http_server, daemon=True).start()

    Shortcut.start()
    Shortcut.register_hot_keys()
    keyboard.wait()
