# -*- coding: utf-8 -*-
"""
Tiny local-only HTTP listener that lets the "Update data now" button on the
public tracker page trigger a real update on this machine.

Binds to 127.0.0.1 only -- never reachable from the network or the internet,
only from a browser running on this same machine. A browser can only get a
response from it (thanks to CORS) if the page making the request comes from
one of ALLOWED_ORIGINS below; every other page's JS is blocked by the browser
itself before the request is even sent (the POST body is JSON, so it always
triggers a CORS preflight -- there is no "simple request" bypass here).

Run this once, in the background, for the lifetime of the machine session.
It's registered as a Windows Scheduled Task ("LSE-CTUR-Tracker-UpdateServer",
trigger: at log on) so it comes back automatically after a restart.
"""
import json
import subprocess
import sys
import threading
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

HOST = "127.0.0.1"
PORT = 8930
REPO_ROOT = Path(__file__).resolve().parent.parent
TRACKER_JSON = REPO_ROOT / "data" / "tracker.json"

ALLOWED_ORIGINS = {
    "https://qizhou-0910.github.io",
    # local testing only:
    "http://localhost:8931",
    "http://127.0.0.1:8931",
}

update_lock = threading.Lock()


def run_update():
    """Runs the same generator + commit/push logic as update_and_push.ps1,
    reimplemented here so this one process can report a structured result
    back to the button instead of just logging to a file."""
    result = subprocess.run(
        [sys.executable, str(REPO_ROOT / "scripts" / "generate_data.py")],
        cwd=REPO_ROOT, capture_output=True, text=True,
    )
    if result.returncode != 0:
        return dict(status="error", message=f"generate_data.py failed:\n{result.stderr[-2000:]}")

    status = subprocess.run(
        ["git", "status", "--porcelain", "data/tracker.json"],
        cwd=REPO_ROOT, capture_output=True, text=True,
    )
    if not status.stdout.strip():
        generated_at = json.loads(TRACKER_JSON.read_text(encoding="utf-8"))["generated_at"]
        return dict(status="unchanged", message="Regenerated -- no change since last update.",
                    generated_at=generated_at)

    subprocess.run(["git", "add", "data/tracker.json"], cwd=REPO_ROOT, check=True)
    commit = subprocess.run(
        ["git", "commit", "-m", f"Manual update via tracker page button ({datetime.now():%Y-%m-%d %H:%M})"],
        cwd=REPO_ROOT, capture_output=True, text=True,
    )
    if commit.returncode != 0:
        return dict(status="error", message=f"git commit failed:\n{commit.stderr[-2000:]}")

    push = subprocess.run(["git", "push"], cwd=REPO_ROOT, capture_output=True, text=True)
    if push.returncode != 0:
        return dict(status="error", message=f"git push failed:\n{push.stderr[-2000:]}")

    generated_at = json.loads(TRACKER_JSON.read_text(encoding="utf-8"))["generated_at"]
    return dict(status="updated", message="Data refreshed and pushed.", generated_at=generated_at)


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        print(f"[{datetime.now().isoformat(timespec='seconds')}] {fmt % args}")

    def _origin_ok(self):
        return self.headers.get("Origin") in ALLOWED_ORIGINS

    def _cors_headers(self):
        origin = self.headers.get("Origin")
        if origin in ALLOWED_ORIGINS:
            self.send_header("Access-Control-Allow-Origin", origin)
            self.send_header("Vary", "Origin")
            self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
            self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors_headers()
        self.end_headers()

    def _send_json(self, code, payload):
        body = json.dumps(payload).encode("utf-8")
        self.send_response(code)
        self._cors_headers()
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path == "/health":
            if not self._origin_ok():
                self._send_json(403, dict(status="error", message="origin not allowed"))
                return
            self._send_json(200, dict(status="ok"))
        else:
            self._send_json(404, dict(status="error", message="not found"))

    def do_POST(self):
        if self.path != "/update":
            self._send_json(404, dict(status="error", message="not found"))
            return
        if not self._origin_ok():
            self._send_json(403, dict(status="error", message="origin not allowed"))
            return
        if not update_lock.acquire(blocking=False):
            self._send_json(409, dict(status="busy", message="An update is already running."))
            return
        try:
            result = run_update()
            self._send_json(200 if result["status"] != "error" else 500, result)
        finally:
            update_lock.release()


if __name__ == "__main__":
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"LSE-CTUR tracker update listener running at http://{HOST}:{PORT} "
          f"(local-only; allowed origins: {sorted(ALLOWED_ORIGINS)})")
    server.serve_forever()
