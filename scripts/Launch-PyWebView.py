#!/usr/bin/env python3
"""
PSDocling window launcher.

Opens the PSDocling UI (served by the API server) in a native pywebview
window. Started by Start-DoclingSystem with pythonw.exe, so there is no
console: diagnostics go to <home>/logs/window.log.

Usage: Launch-PyWebView.py <api_port> [run_dir] [icon_path]

Closing the window shuts PSDocling down (POST /api/shutdown with the per-run
token from run_dir/token.txt), so no server is left listening.
"""

import logging
import os
import sys
import time
from logging.handlers import RotatingFileHandler
from pathlib import Path

import requests
import webview

log = logging.getLogger("psdocling.window")


def setup_logging(run_dir):
    home = Path(run_dir).parent if run_dir else Path(os.environ.get("LOCALAPPDATA", ".")) / "PSDocling"
    log_dir = home / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    handler = RotatingFileHandler(log_dir / "window.log", maxBytes=1_000_000, backupCount=1, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    log.addHandler(handler)
    log.setLevel(logging.INFO)


def read_token(run_dir):
    try:
        return (Path(run_dir) / "token.txt").read_text(encoding="ascii").strip()
    except OSError:
        return ""


def wait_for_backend(api_url, attempts=30, delay=1.0):
    for _ in range(attempts):
        try:
            if requests.get(f"{api_url}/api/health", timeout=2).status_code == 200:
                return True
        except requests.exceptions.RequestException:
            pass
        time.sleep(delay)
    return False


class WindowAPI:
    """Functions exposed to the page as window.pywebview.api.*"""

    def __init__(self, api_url, run_dir):
        self.api_url = api_url
        self.run_dir = run_dir
        self._window = None
        self._shutdown_sent = False

    def set_window(self, window):
        self._window = window

    def shutdown_backend(self):
        """Ask the API to stop all PSDocling processes (once)."""
        if self._shutdown_sent:
            return
        self._shutdown_sent = True
        try:
            requests.post(
                f"{self.api_url}/api/shutdown",
                headers={"X-PSDocling-Token": read_token(self.run_dir)},
                timeout=5,
            )
            log.info("Shutdown requested")
        except requests.exceptions.RequestException as exc:
            log.warning("Shutdown request failed: %s", exc)

    def quit(self):
        """Quit from the UI: stop the backend, then close the window."""
        self.shutdown_backend()
        if self._window:
            self._window.destroy()

    def download_file(self, doc_id, filename):
        """Download a processed document through a native save dialog."""
        try:
            response = requests.get(f"{self.api_url}/api/download/{doc_id}", timeout=60)
            if response.status_code != 200:
                log.warning("Download %s failed with HTTP %s", doc_id, response.status_code)
                return None
            suggested = filename if filename.endswith(".zip") else f"{doc_id}.zip"
            save_path = self._window.create_file_dialog(
                webview.SAVE_DIALOG,
                directory=str(Path.home() / "Downloads"),
                save_filename=suggested,
                file_types=("Zip Files (*.zip)", "All files (*.*)"),
            )
            if isinstance(save_path, (tuple, list)):
                save_path = save_path[0] if save_path else None
            if not save_path:
                return None
            with open(save_path, "wb") as handle:
                handle.write(response.content)
            log.info("Saved download %s", doc_id)
            return str(save_path)
        except Exception:
            log.exception("Download %s failed", doc_id)
            return None


def set_window_icon(window, icon_path):
    """WinForms backend: give the window and taskbar button the PSDocling icon."""
    if not icon_path or not Path(icon_path).is_file():
        return
    try:
        from System.Drawing import Icon  # pythonnet, present with the WinForms backend

        form = window.native
        form.Invoke(lambda: setattr(form, "Icon", Icon(icon_path)))
    except Exception as exc:
        log.warning("Could not set window icon: %s", exc)


def main():
    api_port = int(sys.argv[1]) if len(sys.argv) > 1 else 8080
    run_dir = sys.argv[2] if len(sys.argv) > 2 else ""
    icon_path = sys.argv[3] if len(sys.argv) > 3 else ""
    setup_logging(run_dir)

    if sys.platform == "win32":
        # Own taskbar group instead of "Python"
        try:
            import ctypes

            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("PSDocling.App")
        except Exception:
            pass

    api_url = f"http://localhost:{api_port}"
    if not wait_for_backend(api_url):
        log.error("Backend did not respond at %s", api_url)
        return 1

    api = WindowAPI(api_url, run_dir)
    window = webview.create_window(
        title="PSDocling",
        url=api_url,
        width=1400,
        height=900,
        min_size=(800, 600),
        background_color="#0f1115",
        js_api=api,
    )
    api.set_window(window)
    window.events.shown += lambda: set_window_icon(window, icon_path)
    window.events.closed += api.shutdown_backend

    log.info("Window opening at %s", api_url)
    webview.start(debug=False, http_server=False, gui="edgechromium")
    api.shutdown_backend()
    log.info("Window closed")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        log.exception("Window launcher crashed")
        sys.exit(1)
