import json
import socket
import sys
import time

from shouyu.log import Log
from shouyu.util.package import Package
from shouyu.view.msgbox import MessageBox


HTTP_HOST = "127.0.0.1"
HTTP_PORT = 19823


def _is_http_server_up(host, port, timeout=0.5):
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(timeout)
    try:
        result = sock.connect_ex((host, port))
        return result == 0
    except Exception:
        return False
    finally:
        sock.close()


def _send_via_http(title):
    payload = json.dumps({"title": title}).encode("utf-8")
    import http.client
    conn = http.client.HTTPConnection(HTTP_HOST, HTTP_PORT, timeout=5)
    try:
        conn.request("POST", "/add-title", body=payload, headers={"Content-Type": "application/json"})
        response = conn.getresponse()
        if response.status == 200:
            data = json.loads(response.read().decode("utf-8"))
            MessageBox.pop_up_message("Queued", data.get("message", "请求已加入队列"))
        else:
            MessageBox.pop_up_message("Error", f"HTTP {response.status}: {response.read().decode()}")
    except Exception as e:
        MessageBox.pop_up_message("Failed", f"请求失败: {e}")
    finally:
        conn.close()


def _build_title_from_args(argv):
    return " ".join(argv).strip()


def main(argv):
    Package.set_cwd()
    Log.setup()

    title = _build_title_from_args(argv)
    if not title:
        MessageBox.pop_up_message("Missing Title", "请传入标题：python cmd.py \"this is a new plan title\"")
        return

    # 优先尝试通过 main.py 的 HTTP 服务写入（走队列，保持与快捷键一致的行为）
    if _is_http_server_up(HTTP_HOST, HTTP_PORT):
        _send_via_http(title)
    else:
        # 如果 main.py 未启动，则回退到直接写入
        from shouyu.service.excel import KbExcel
        KbExcel.append_title_to_next_row(title)


if __name__ == '__main__':
    main(sys.argv[1:])
