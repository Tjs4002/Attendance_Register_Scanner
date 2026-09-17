#!/usr/bin/env python3
"""
Lightweight HTTP Server for Maha AI Olympiad Register Scanner.
Usage:
    python server.py [port]
"""

import http.server
import socketserver
import os
import sys
import webbrowser

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 3000
DIRECTORY = os.path.dirname(os.path.abspath(__file__))

class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def end_headers(self):
        # Enable CORS and disable cache during development
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate')
        super().end_headers()

def run():
    # Try finding an open port starting from PORT
    port = PORT
    max_tries = 10
    httpd = None
    
    for p in range(port, port + max_tries):
        try:
            httpd = socketserver.TCPServer(("", p), Handler)
            port = p
            break
        except OSError:
            continue

    if not httpd:
        print(f"Error: Could not bind to any port between {PORT} and {PORT + max_tries - 1}")
        sys.exit(1)

    url = f"http://localhost:{port}/index.html"
    print("=" * 60)
    print("  Maha AI Olympiad - Attendance Register Scanner")
    print("=" * 60)
    print(f"  Server running at: {url}")
    print("  Press Ctrl+C to stop.")
    print("=" * 60)

    try:
        webbrowser.open(url)
    except Exception:
        pass

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nServer stopped.")
        httpd.server_close()

if __name__ == '__main__':
    run()
