import os
import socket
import sys
import threading
import time
import webbrowser


def _app_root() -> str:
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
        if os.path.basename(exe_dir).lower() == "kintaichecker":
            dist_dir = os.path.dirname(exe_dir)
            if os.path.basename(dist_dir).lower() == "dist":
                return os.path.dirname(dist_dir)
        return exe_dir
    return os.path.dirname(os.path.abspath(__file__))


ROOT_DIR = _app_root()
os.chdir(ROOT_DIR)


def _open_browser():
    if os.environ.get("KINTAI_CHECKER_NO_BROWSER") == "1":
        return
    # サーバーのポートが開いたら即ブラウザを開く（固定3秒待ちをやめて起動を体感短縮）
    deadline = time.time() + 30
    while time.time() < deadline:
        try:
            with socket.create_connection(("127.0.0.1", 5000), timeout=0.5):
                break
        except OSError:
            time.sleep(0.2)
    webbrowser.open("http://localhost:5000")


if __name__ == "__main__":
    from app import app, cleanup_uploads_on_start

    # 前回までの実行で uploads に残った入力ファイルの写しを掃除する。
    # 氏名や勤怠を含むファイルが共有フォルダに残り続けないようにするため。
    cleanup_uploads_on_start()

    threading.Thread(target=_open_browser, daemon=True).start()
    # 各自のPCでローカル起動する運用のため、待ち受けは自分のPC内のみ（LANに公開しない）
    app.run(debug=False, host="127.0.0.1", port=5000)
