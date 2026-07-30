"""
用 pywebview 把本機跑的 Streamlit App 包成一個看起來像原生視窗的程式：
沒有網址列、沒有瀏覽器分頁，也不會跳出黑色終端機視窗。

Streamlit 伺服器只監聽 127.0.0.1（設定見 .streamlit/config.toml），資料不會被
同一個網路上的其他電腦存取到；這支腳本另外把 Streamlit 子行程的主控台視窗也
關閉（CREATE_NO_WINDOW），輸出改寫進 logs\\app.log，出問題時可以把那份記錄
內容貼出來排查。
"""
import os
import socket
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

HERE = Path(__file__).resolve().parent
LOG_DIR = HERE / "logs"
LOG_FILE = LOG_DIR / "app.log"


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _ensure_streamlit_credentials():
    """第一次在這台電腦/這個帳號執行 streamlit 時，官方會互動式詢問 email，
    子行程沒有畫面可以輸入時會卡住，這裡先建立一份空白設定跳過詢問。"""
    cfg_dir = Path.home() / ".streamlit"
    cfg_file = cfg_dir / "credentials.toml"
    if not cfg_file.exists():
        cfg_dir.mkdir(parents=True, exist_ok=True)
        cfg_file.write_text('[general]\nemail = ""\n', encoding="utf-8")


def _wait_for_server(port: int, timeout: float = 30.0) -> bool:
    deadline = time.time() + timeout
    url = f"http://127.0.0.1:{port}"
    while time.time() < deadline:
        try:
            urllib.request.urlopen(url, timeout=1)
            return True
        except Exception:
            time.sleep(0.5)
    return False


def main():
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    _ensure_streamlit_credentials()

    env = dict(os.environ)
    playwright_browsers = HERE.parent / "ms-playwright"
    if playwright_browsers.exists():
        env["PLAYWRIGHT_BROWSERS_PATH"] = str(playwright_browsers)

    port = _free_port()
    log_fh = open(LOG_FILE, "a", encoding="utf-8")
    creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
    proc = subprocess.Popen(
        [sys.executable, "-m", "streamlit", "run", str(HERE / "app.py"),
         "--server.port", str(port)],
        cwd=str(HERE),
        env=env,
        stdout=log_fh,
        stderr=subprocess.STDOUT,
        creationflags=creationflags,
    )

    try:
        if not _wait_for_server(port):
            raise RuntimeError(f"Streamlit 伺服器啟動逾時，請查看 {LOG_FILE}")

        import webview
        webview.create_window("出貨自動化工具", f"http://127.0.0.1:{port}", width=1440, height=960)
        webview.start()
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except Exception:
            proc.kill()
        log_fh.close()


if __name__ == "__main__":
    main()
