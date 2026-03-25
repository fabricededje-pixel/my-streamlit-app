import os
import socket
import sys
import threading
import time
import webbrowser

from streamlit.web.bootstrap import run


def resource_path(relative_path: str) -> str:
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def find_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


def open_browser_later(url: str) -> None:
    def _open() -> None:
        time.sleep(2)
        webbrowser.open(url)

    threading.Thread(target=_open, daemon=True).start()


def main() -> None:
    app_path = resource_path("app.py")
    app_dir = os.path.dirname(app_path)
    os.chdir(app_dir)

    port = find_free_port()
    url = f"http://127.0.0.1:{port}"
    open_browser_later(url)

    run(
        app_path,
        False,
        [],
        {
            "server.headless": True,
            "server.port": port,
            "server.address": "127.0.0.1",
            "browser.serverAddress": "127.0.0.1",
            "browser.gatherUsageStats": False,
        },
    )


if __name__ == "__main__":
    main()
