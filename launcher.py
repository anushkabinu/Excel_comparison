"""
launcher.py
-----------
Entry point for the PyInstaller EXE.
Starts the Streamlit server and opens the browser automatically.
"""

import os
import sys
import threading
import time
import traceback
import webbrowser


def _open_browser(port: int = 8501, delay: float = 4.0) -> None:
    """Wait for the server to start, then open the browser."""
    time.sleep(delay)
    webbrowser.open_new(f"http://localhost:{port}")


def _get_app_path() -> str:
    """
    Return the absolute path to app.py.
    When frozen by PyInstaller, files are extracted to sys._MEIPASS.
    """
    if getattr(sys, "frozen", False):
        base = sys._MEIPASS          # type: ignore[attr-defined]
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "app.py")


def main() -> None:
    port = 8501
    app_path = _get_app_path()

    print(f"[Launcher] App path : {app_path}")
    print(f"[Launcher] App exists: {os.path.exists(app_path)}")
    print(f"[Launcher] Starting Streamlit on port {port}...")

    # Open browser in background after server starts
    threading.Thread(target=_open_browser, args=(port, 4.0), daemon=True).start()

    # Start Streamlit via its internal CLI
    from streamlit.web import cli as stcli
    sys.argv = [
        "streamlit",
        "run",
        app_path,
        f"--server.port={port}",
        "--server.headless=true",
        "--global.developmentMode=false",   # required when frozen by PyInstaller
        "--browser.gatherUsageStats=false",
    ]
    stcli.main()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        print("\n" + "=" * 60)
        print("ERROR — the app failed to start. Details:")
        print("=" * 60)
        traceback.print_exc()
        print("=" * 60)
        input("\nPress ENTER to close this window...")
        sys.exit(1)
