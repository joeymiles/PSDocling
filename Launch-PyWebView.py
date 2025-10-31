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

    # Window configuration
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

    # Start the webview
    webview.start(debug=False)

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
