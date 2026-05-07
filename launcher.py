import os
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
    time.sleep(3)
    webbrowser.open("http://localhost:5000")


if __name__ == "__main__":
    from app import app

    threading.Thread(target=_open_browser, daemon=True).start()
    app.run(debug=False, host="0.0.0.0", port=5000)
