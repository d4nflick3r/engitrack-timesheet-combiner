import sys
import os
import socket
import time
import multiprocessing
import requests
import webview
from streamlit.web import cli as stcli

sys.setrecursionlimit(5000)


def find_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("", 0))
        s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        return s.getsockname()[1]


def run_streamlit(script_path, port):
    if getattr(sys, "frozen", False):
        os.chdir(sys._MEIPASS)
    sys.argv = [
        "streamlit", "run", script_path,
        "--server.address=localhost",
        f"--server.port={port}",
        "--server.headless=true",
        "--global.developmentMode=false",
    ]
    stcli.main()


def wait_for_server(port, timeout=120):
    start = time.time()
    url = f"http://localhost:{port}"
    while True:
        try:
            requests.get(url, timeout=2)
            return
        except Exception:
            if time.time() - start > timeout:
                raise TimeoutError(
                    f"Streamlit server did not start within {timeout}s."
                )
            time.sleep(0.5)


if __name__ == "__main__":
    multiprocessing.freeze_support()

    if getattr(sys, "frozen", False):
        os.chdir(sys._MEIPASS)

    port = find_free_port()

    proc = multiprocessing.Process(target=run_streamlit, args=("app.py", port))
    proc.start()

    try:
        wait_for_server(port, timeout=120)
        webview.create_window(
            "EngiTrack Timesheet Combiner",
            f"http://localhost:{port}",
            width=1400,
            height=900,
        )
        webview.start()
    finally:
        proc.terminate()
        proc.join()
