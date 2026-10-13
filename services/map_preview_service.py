"""Serve map previews over HTTP, without exposing workspace files."""

import atexit
import secrets
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer


class MapPreviewServer:
    def __init__(self):
        self.pages = {}
        self.lock = threading.Lock()
        owner = self

        class Handler(BaseHTTPRequestHandler):
            def do_GET(self):
                with owner.lock:
                    page = owner.pages.get(self.path)
                if page is None:
                    self.send_error(404)
                    return
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(page)))
                self.send_header("Cache-Control", "no-store")
                self.send_header("Referrer-Policy", "strict-origin-when-cross-origin")
                self.end_headers()
                self.wfile.write(page)

            def log_message(self, format, *args):
                pass

        self.server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self.server.daemon_threads = True
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        atexit.register(self.close)

    def publish(self, html):
        path = "/" + secrets.token_urlsafe(24)
        with self.lock:
            self.pages[path] = html.encode("utf-8")
            # Retain recent previews so reopening the map does not grow memory forever.
            while len(self.pages) > 32:
                del self.pages[next(iter(self.pages))]
        return f"http://127.0.0.1:{self.server.server_port}{path}"

    def close(self):
        self.server.shutdown()
        self.server.server_close()


_preview_server = None


def publish_map_preview(html):
    global _preview_server
    if _preview_server is None:
        _preview_server = MapPreviewServer()
    return _preview_server.publish(html)
