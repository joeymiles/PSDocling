#!/usr/bin/env python3
"""
PSDocling PyWebView Launcher
Launches the PSDocling web interface in a native desktop window using pywebview
"""

import sys
import time
import requests
import webview
import threading
import os
from pathlib import Path

def check_backend_ready(url, max_attempts=30, delay=1):
    """Check if the backend server is ready"""
    print(f"Waiting for backend at {url}...")
    for attempt in range(max_attempts):
        try:
            response = requests.get(f"{url}/api/health", timeout=2)
            if response.status_code == 200:
                print(f"Backend ready after {attempt + 1} attempts")
                return True
        except requests.exceptions.RequestException:
            pass
        time.sleep(delay)
    print(f"Backend not ready after {max_attempts} attempts")
    return False

class DownloadAPI:
    """API class to expose download functionality to JavaScript"""
    def __init__(self, api_url, window):
        self.api_url = api_url
        self.window = window

    def download_file(self, doc_id, filename):
        """Download a file from the API with native save dialog"""
        try:
            print(f"Downloading document: {doc_id}")
            # Construct the API URL
            download_url = f"{self.api_url}/api/download/{doc_id}"

            # Fetch the file
            response = requests.get(download_url, timeout=60)
            if response.status_code == 200:
                # Suggest filename with .zip extension if not present
                suggested_name = filename if filename.endswith('.zip') else f"{filename}.zip"

                # Show native save file dialog
                file_types = ('Zip Files (*.zip)', 'All files (*.*)')
                save_path = self.window.create_file_dialog(
                    webview.SAVE_DIALOG,
                    directory=str(Path.home() / "Downloads"),
                    save_filename=suggested_name,
                    file_types=file_types
                )

                if save_path:
                    # User selected a location, save the file
                    with open(save_path, 'wb') as f:
                        f.write(response.content)
                    print(f"Downloaded to: {save_path}")
                    return str(save_path)
                else:
                    print("Download cancelled by user")
                    return None
            else:
                print(f"Download failed with status: {response.status_code}")
                return None
        except Exception as e:
            print(f"Download error: {e}")
            return None

def main():
    # Configuration
    api_port = 8080
    web_port = 8081

    # Parse command line arguments for custom ports
    if len(sys.argv) > 1:
        try:
            api_port = int(sys.argv[1])
        except ValueError:
            print(f"Invalid API port: {sys.argv[1]}, using default 8080")

    if len(sys.argv) > 2:
        try:
            web_port = int(sys.argv[2])
        except ValueError:
            print(f"Invalid Web port: {sys.argv[2]}, using default 8081")

    # URLs
    api_url = f"http://localhost:{api_port}"
    web_url = f"http://localhost:{web_port}"

    # Check if backend is ready
    if not check_backend_ready(api_url):
        print("ERROR: Backend API is not responding. Please start the backend services first.")
        sys.exit(1)

    # Create and configure the webview window
    print(f"Launching PSDocling interface at {web_url}")

    # Window configuration with file download support
    window = webview.create_window(
        title='PSDocling - Document Processor',
        url=web_url,
        width=1400,
        height=900,
        resizable=True,
        fullscreen=False,
        min_size=(800, 600),
        background_color='#0f1115'
    )

    # Create API instance and pass window reference for file dialogs
    download_api = DownloadAPI(api_url, window)
    window.expose(download_api)

    # Configure download handler
    def on_download(download_path):
        """Handle downloads"""
        print(f"Download started: {download_path}")

    # Start the webview with download support and proper configuration
    # On Windows, this uses Edge/Chromium which supports blob downloads natively
    webview.start(debug=False, http_server=False, gui='edgechromium')

    print("PyWebView window closed")

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\nInterrupted by user")
        sys.exit(0)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)
