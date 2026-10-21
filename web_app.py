#!/usr/bin/env python3
"""Local-only browser interface for the Outlook cleaner.

All mailbox operations run through the existing CLI in a child process. The
server never stores credentials in browser storage or returns review JSON.
"""

from __future__ import annotations

import argparse
import html
import json
import os
import re
import secrets
import subprocess
import sys
import threading
from collections import Counter
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlsplit

PROJECT = Path(__file__).resolve().parent
REVIEWS = PROJECT / "reviews"
REVIEW_NAME = re.compile(r"review-\d{8}-\d{6}-\d+\.json\Z")
HTML_NAME = re.compile(r"review-\d{8}-\d{6}-\d+\.html\Z")
MAX_REQUEST = 100_000


def review_path(name: str) -> Path:
    if not REVIEW_NAME.fullmatch(name):
        raise ValueError("Invalid review filename")
    path = REVIEWS / name
    if not path.is_file() or path.is_symlink():
        raise ValueError("Review does not exist")
    return path


def public_counts(path: Path) -> dict[str, int]:
    data = json.loads(path.read_text(encoding="utf-8"))
    counts = Counter(item.get("decision") for item in data.get("items", []))
    return {
        "total": len(data.get("items", [])),
        "keep": counts["keep"],
        "review": counts["review"],
        "delete_candidate": counts["delete_candidate"],
    }


class Jobs:
    def __init__(self) -> None:
        self.lock = threading.Lock()
        self.active = False
        self.status: dict[str, object] = {"state": "idle"}

    def start(self, command: str, argv: list[str]) -> bool:
        with self.lock:
            if self.active:
                return False
            self.active = True
            self.status = {"state": "running", "operation": command}
        thread = threading.Thread(target=self._run, args=(command, argv), daemon=True)
        thread.start()
        return True

    def _run(self, command: str, argv: list[str]) -> None:
        REVIEWS.mkdir(mode=0o700, parents=True, exist_ok=True)
        log_path = REVIEWS / f"web-run-{datetime.now().strftime('%Y%m%d-%H%M%S-%f')}.log"
        try:
            fd = os.open(log_path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
            with os.fdopen(fd, "w", encoding="utf-8") as log:
                result = subprocess.run(
                    [sys.executable, str(PROJECT / "jev_outlook_cleaner.py"), *argv],
                    cwd=PROJECT, stdout=log, stderr=subprocess.STDOUT, check=False,
                    env={**os.environ, "JEV_NONINTERACTIVE": "1"},
                )
            lines = log_path.read_text(encoding="utf-8", errors="replace").splitlines()
            summary: dict[str, object] = {}
            if command == "scan" and result.returncode == 0:
                names = [line.partition("Review saved to: ")[2] for line in lines
                         if line.startswith("Review saved to: ")]
                if names:
                    name = Path(names[-1]).name
                    path = review_path(name)
                    summary = {"review": name, **public_counts(path)}
            elif command == "auto-apply":
                match = re.search(
                    r"Automatically selected (\d+); moved (\d+); skipped (\d+); failed (\d+)",
                    "\n".join(lines),
                )
                if match:
                    summary = dict(zip(("selected", "moved", "skipped", "failed"),
                                       map(int, match.groups())))
            elif command == "apply-selection":
                match = re.search(r"保留 (\d+)；移動 (\d+)；跳過 (\d+)；失敗 (\d+)", "\n".join(lines))
                if match:
                    summary = dict(zip(("kept", "moved", "skipped", "failed"),
                                       map(int, match.groups())))
            with self.lock:
                self.status = {"state": "done" if result.returncode == 0 else "failed",
                               "operation": command, **summary}
        except Exception:
            with self.lock:
                self.status = {"state": "failed", "operation": command}
        finally:
            with self.lock:
                self.active = False

    def snapshot(self) -> dict[str, object]:
        with self.lock:
            return dict(self.status)


JOBS = Jobs()
CSRF = secrets.token_urlsafe(32)


def dashboard() -> str:
    cards = []
    for path in sorted(REVIEWS.glob("review-*.json"), reverse=True)[:10] if REVIEWS.exists() else []:
        if not REVIEW_NAME.fullmatch(path.name) or path.is_symlink():
            continue
        try:
            count = public_counts(path)
        except (OSError, ValueError, KeyError):
            continue
        name = html.escape(path.name)
        cards.append(
            f'<article><h3>{name}</h3><p>共 {count["total"]} 封 · 保留 {count["keep"]} · '
            f'待確認 {count["review"]} · 刪除候選 {count["delete_candidate"]}</p>'
            f'<p><a href="/review/{name[:-5]}.html">查看預覽</a></p>'
            f'<button data-auto="{name}">自動套用高信心候選</button> '
            f'<label>上傳預覽下載的選擇檔 <input type="file" data-selection="{name}" '
            f'accept="application/json,.json"></label></article>'
        )
    return f'''<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>Outlook Jev 清理</title>
<style>body{{font:16px system-ui;max-width:940px;margin:40px auto;padding:0 20px;color:#17212e;background:#f5f7fa}}
h1{{font-size:30px}}article,.panel{{background:white;padding:22px;border:1px solid #d9e0e8;border-radius:14px;margin:16px 0}}
button,input{{font:inherit}}button{{background:#1769aa;color:white;border:0;border-radius:8px;padding:10px 14px;cursor:pointer}}
button[data-auto]{{background:#9d2c23}}input[type=number]{{width:90px;padding:8px}}label{{margin-right:12px}}
small{{color:#526071}}#status{{white-space:pre-wrap}}</style></head><body>
<h1>Outlook Jev 清理</h1><p>僅在這台電腦上運作。保留與待確認郵件留在 Inbox；移動只會到可復原的「刪除的郵件」。</p>
<p><small>首次使用請先在終端執行 <code>.venv/bin/python jev_outlook_cleaner.py login</code> 完成 Microsoft 登入。</small></p>
<section class="panel"><h2>掃描 Inbox</h2><label>超過 <input id="days" type="number" min="1" max="3650" value="30"> 天</label>
<label>最多 <input id="limit" type="number" min="1" max="200" value="200"> 封</label>
<button id="scan">開始掃描與分類</button><p><small>掃描會略過保留清單和已處理郵件。分類失敗即停止。</small></p></section>
<section class="panel"><h2>執行狀態</h2><p id="status">讀取中…</p><button onclick="location.reload()">更新預覽清單</button></section>
<h2>最近預覽</h2>{''.join(cards) or '<p>尚無預覽。</p>'}
<script>
const token = {json.dumps(CSRF)};
async function post(path, value) {{
  const response = await fetch(path, {{method:'POST',headers:{{'Content-Type':'application/json','X-CSRF-Token':token}},body:JSON.stringify(value)}});
  if (!response.ok) throw new Error('操作未啟動（HTTP ' + response.status + '）');
  await refresh();
}}
async function refresh() {{
  try {{
    const value = await (await fetch('/api/status',{{cache:'no-store'}})).json();
    const labels = {{idle:'待命',running:'執行中',done:'已完成',failed:'失敗，已停止。請查看本機終端紀錄。'}};
    document.getElementById('status').textContent = (labels[value.state] || value.state) + '\\n' +
      Object.entries(value).filter(([key])=>!['state','operation'].includes(key)).map(([key,val])=>key+': '+val).join(' · ');
    if (value.state==='running') setTimeout(refresh, 1500);
  }} catch {{ document.getElementById('status').textContent='無法讀取執行狀態'; }}
}}
document.getElementById('scan').onclick=async()=>{{
  try {{await post('/api/scan',{{days:Number(document.getElementById('days').value),limit:Number(document.getElementById('limit').value)}})}}
  catch(e){{alert(e.message)}}
}};
for(const button of document.querySelectorAll('[data-auto]')) button.onclick=async()=>{{
  if(!confirm('只移動已核准規則或高信心群發刪除候選；繼續？')) return;
  try{{await post('/api/auto-apply',{{review:button.dataset.auto}})}}catch(e){{alert(e.message)}}
}};
for(const input of document.querySelectorAll('[data-selection]')) input.onchange=async()=>{{
  if(!input.files.length) return;
  if(!confirm('依選擇檔保留勾選郵件，將未勾選郵件移至「刪除的郵件」？')) return;
  try{{await post('/api/apply-selection',{{review:input.dataset.selection,selection:JSON.parse(await input.files[0].text())}})}}
  catch(e){{alert(e.message)}}
}};
refresh();
</script></body></html>'''


class Handler(BaseHTTPRequestHandler):
    server_version = "JevLocalWeb/1"

    def log_message(self, _format: str, *args: object) -> None:
        # URLs and review names may reveal mailbox data; do not log requests.
        pass

    def send_bytes(self, code: int, content: bytes, content_type: str) -> None:
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(content)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        self.send_header("Content-Security-Policy", "default-src 'none'; style-src 'unsafe-inline'; script-src 'unsafe-inline'; connect-src 'self'; form-action 'self'; base-uri 'none'; frame-ancestors 'none'")
        self.end_headers()
        self.wfile.write(content)

    def allowed_host(self) -> bool:
        return self.headers.get("Host", "") in {
            f"127.0.0.1:{self.server.server_port}", f"localhost:{self.server.server_port}"
        }

    def do_GET(self) -> None:
        if not self.allowed_host():
            return self.send_bytes(403, b"Forbidden", "text/plain")
        path = urlsplit(self.path).path
        if path == "/":
            return self.send_bytes(200, dashboard().encode(), "text/html; charset=utf-8")
        if path == "/api/status":
            return self.send_bytes(200, json.dumps(JOBS.snapshot()).encode(), "application/json")
        if path.startswith("/review/"):
            name = path.removeprefix("/review/")
            if HTML_NAME.fullmatch(name):
                try:
                    file = review_path(name[:-5] + ".json").with_suffix(".html")
                    if file.is_file() and not file.is_symlink():
                        return self.send_bytes(200, file.read_bytes(), "text/html; charset=utf-8")
                except ValueError:
                    pass
        return self.send_bytes(404, b"Not found", "text/plain")

    def do_POST(self) -> None:
        if not self.allowed_host():
            return self.send_bytes(403, b"Forbidden", "text/plain")
        origin = self.headers.get("Origin", "")
        allowed_origins = {f"http://127.0.0.1:{self.server.server_port}",
                           f"http://localhost:{self.server.server_port}"}
        if origin not in allowed_origins or self.headers.get("X-CSRF-Token") != CSRF:
            return self.send_bytes(403, b"Forbidden", "text/plain")
        if self.headers.get("Content-Type", "").split(";", 1)[0] != "application/json":
            return self.send_bytes(415, b"JSON required", "text/plain")
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if not 0 < length <= MAX_REQUEST:
                raise ValueError("Invalid request length")
            payload = json.loads(self.rfile.read(length))
            if not isinstance(payload, dict):
                raise ValueError("Invalid request")
            path = urlsplit(self.path).path
            if path == "/api/scan":
                days, limit = payload.get("days"), payload.get("limit")
                if type(days) is not int or type(limit) is not int or not 1 <= days <= 3650 or not 1 <= limit <= 200:
                    raise ValueError("Invalid scan limits")
                started = JOBS.start("scan", ["scan", "--older-than-days", str(days), "--limit", str(limit)])
            elif path == "/api/auto-apply":
                path_to_review = review_path(payload.get("review", ""))
                started = JOBS.start("auto-apply", ["auto-apply", str(path_to_review)])
            elif path == "/api/apply-selection":
                path_to_review = review_path(payload.get("review", ""))
                selection = payload.get("selection")
                if not isinstance(selection, dict) or selection.get("schema_version") != 1:
                    raise ValueError("Invalid selection")
                REVIEWS.mkdir(mode=0o700, parents=True, exist_ok=True)
                selection_path = REVIEWS / f"web-selection-{secrets.token_hex(12)}.json"
                fd = os.open(selection_path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
                with os.fdopen(fd, "w", encoding="utf-8") as file:
                    json.dump(selection, file)
                started = JOBS.start("apply-selection", ["apply-selection", str(path_to_review), str(selection_path)])
                if not started:
                    selection_path.unlink(missing_ok=True)
            else:
                return self.send_bytes(404, b"Not found", "text/plain")
            return self.send_bytes(202 if started else 409, b"Accepted" if started else b"Busy", "text/plain")
        except (ValueError, TypeError, json.JSONDecodeError):
            return self.send_bytes(400, b"Invalid request", "text/plain")


def main() -> None:
    parser = argparse.ArgumentParser(description="Local-only web interface for Outlook Jev cleanup")
    parser.add_argument("--port", type=int, default=8765)
    args = parser.parse_args()
    if not 1 <= args.port <= 65535:
        parser.error("port must be between 1 and 65535")
    server = ThreadingHTTPServer(("127.0.0.1", args.port), Handler)
    print(f"Open http://127.0.0.1:{server.server_port}/ in this computer's browser")
    server.serve_forever()


if __name__ == "__main__":
    main()
