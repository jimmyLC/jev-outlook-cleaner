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

from ui_theme import BASE_CSS, DASHBOARD_CSS

PROJECT = Path(__file__).resolve().parent
REVIEWS = PROJECT / "reviews"
REVIEW_NAME = re.compile(r"review-\d{8}-\d{6}-\d+\.json\Z")
HTML_NAME = re.compile(r"review-\d{8}-\d{6}-\d+\.html\Z")
MAX_REQUEST = 100_000
LOGIN_TIMEOUT = 16 * 60  # Microsoft device codes expire after about 15 minutes
DEVICE_PREFIX = "JEV_DEVICE_LOGIN "
DEVICE_URL = re.compile(r"https://(?:www\.)?(?:microsoft\.com|login\.microsoftonline\.com)/[\w/.-]+\Z")
DEVICE_CODE = re.compile(r"[A-Z0-9]{6,12}\Z")


def parse_device_login(line: str) -> tuple[str, str] | None:
    """Return (url, code) from the CLI's structured line, or None if invalid."""
    if not line.startswith(DEVICE_PREFIX):
        return None
    try:
        info = json.loads(line[len(DEVICE_PREFIX):])
        url, code = info["url"], info["code"]
    except (ValueError, KeyError, TypeError):
        return None
    if isinstance(url, str) and isinstance(code, str) and DEVICE_URL.match(url) and DEVICE_CODE.match(code):
        return url, code
    return None


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
                    summary = {"review_file": name, **public_counts(path)}
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
            if result.returncode != 0 and "需要互動登入" in "\n".join(lines):
                summary = {"reason": "need_login"}
            with self.lock:
                self.status = {"state": "done" if result.returncode == 0 else "failed",
                               "operation": command, **summary}
        except Exception:
            with self.lock:
                self.status = {"state": "failed", "operation": command}
        finally:
            with self.lock:
                self.active = False

    def start_login(self) -> bool:
        """Run the CLI `login` command and surface its Microsoft device code."""
        with self.lock:
            if self.active:
                return False
            self.active = True
            self.status = {"state": "running", "operation": "login"}
        threading.Thread(target=self._run_login, daemon=True).start()
        return True

    def _run_login(self) -> None:
        timer = None
        try:
            env = {key: value for key, value in os.environ.items() if key != "JEV_NONINTERACTIVE"}
            proc = subprocess.Popen(
                [sys.executable, "-u", str(PROJECT / "jev_outlook_cleaner.py"), "login"],
                cwd=PROJECT, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, env=env,
            )
            timer = threading.Timer(LOGIN_TIMEOUT, proc.kill)
            timer.start()
            for line in proc.stdout:
                device = parse_device_login(line)
                if device:
                    with self.lock:
                        self.status = {"state": "running", "operation": "login",
                                       "login_url": device[0], "login_code": device[1]}
            proc.wait()
            with self.lock:
                self.status = {"state": "done" if proc.returncode == 0 else "failed", "operation": "login"}
        except Exception:
            with self.lock:
                self.status = {"state": "failed", "operation": "login"}
        finally:
            if timer:
                timer.cancel()
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
        stamp = html.escape(f"{path.name[7:11]}-{path.name[11:13]}-{path.name[13:15]} "
                            f"{path.name[16:18]}:{path.name[18:20]}")
        cards.append(
            f'<article class="card review-row"><div class="meta"><span class="name">{stamp}</span>'
            f'<div class="chips"><span class="chip total">共 {count["total"]} 封</span>'
            f'<span class="chip keep">保留 {count["keep"]}</span>'
            f'<span class="chip review">待確認 {count["review"]}</span>'
            f'<span class="chip delete_candidate">刪除候選 {count["delete_candidate"]}</span></div></div>'
            f'<div class="actions"><a class="btn" href="/review/{name[:-5]}.html">查看預覽</a>'
            f'<label class="btn ghost file-btn">上傳選擇檔<input type="file" data-selection="{name}" '
            f'accept="application/json,.json" aria-label="上傳選擇檔"></label>'
            f'<button class="danger-ghost" data-auto="{name}" title="只移動已核准規則或高信心群發候選">自動套用</button>'
            f'</div></article>'
        )
    empty = '<div class="card empty-state">尚無預覽。先在上方開始掃描。</div>'
    return f'''<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>Outlook Jev 清理</title>
<style>{BASE_CSS}{DASHBOARD_CSS}</style></head><body><div class="wrap">
<div class="top"><div><h1>Outlook Jev 清理</h1>
<p class="sub">保留與待確認郵件留在 Inbox；移動只會到可復原的「刪除的郵件」。</p></div>
<span class="badge-local">僅限本機 127.0.0.1</span></div>
<div class="steps">
<section class="panel"><h2><span class="step-no">1</span>登入 Microsoft</h2>
<div id="login-idle"><p class="small">首次使用或登入過期時，按下按鈕取得登入代碼。登入只讀取你的帳號資料，不會掃描或移動郵件。</p>
<p style="margin-top:12px"><button id="login">登入 Outlook</button></p></div>
<div id="login-wait" hidden><p>① <a id="login-link" target="_blank" rel="noopener noreferrer">開啟 Microsoft 登入頁面</a></p>
<p style="margin-top:6px">② 輸入這組代碼：</p><div id="login-code" class="code"></div>
<p class="small">完成後此處會自動更新；代碼約 15 分鐘內有效。</p></div></section>
<section class="panel"><h2><span class="step-no">2</span>掃描 Inbox</h2>
<div class="fields"><label>超過幾天<input id="days" type="number" min="1" max="3650" value="30"></label>
<label>最多幾封<input id="limit" type="number" min="1" max="200" value="200"></label>
<button id="scan">開始掃描與分類</button></div>
<p class="small">會略過保留清單與已處理郵件；分類失敗即整批停止。</p></section></div>
<section class="panel" aria-live="polite"><div id="status-box"><span class="dot" id="dot"></span>
<div><strong id="status-label">讀取中…</strong><div id="status-detail"></div></div>
<button class="ghost" style="margin-left:auto" onclick="location.reload()">更新清單</button></div></section>
<h2>最近預覽</h2>{''.join(cards) or empty}
</div><script>
const token = {json.dumps(CSRF)};
async function post(path, value) {{
  const response = await fetch(path, {{method:'POST',headers:{{'Content-Type':'application/json','X-CSRF-Token':token}},body:JSON.stringify(value)}});
  if (!response.ok) throw new Error(response.status === 409 ? '已有作業進行中，請稍候' : '操作未啟動（HTTP ' + response.status + '）');
  await refresh();
}}
const names = {{login:'登入',scan:'掃描',"auto-apply":'自動套用',"apply-selection":'套用選擇檔'}};
const keys = {{total:'共',keep:'保留',review:'待確認',delete_candidate:'刪除候選',selected:'符合',moved:'已移動',skipped:'跳過',failed:'失敗',kept:'保留'}};
async function refresh() {{
  try {{
    const value = await (await fetch('/api/status',{{cache:'no-store'}})).json();
    const labels = {{idle:'待命',running:'執行中…',done:'已完成',failed:'失敗，已停止；請查看 reviews 內的紀錄'}};
    const op = names[value.operation] ? names[value.operation] + '：' : '';
    const waiting = value.operation === 'login' && value.state === 'running';
    document.getElementById('login-idle').hidden = waiting && !!value.login_code;
    document.getElementById('login-wait').hidden = !(waiting && value.login_code);
    if (waiting && value.login_code && value.login_url.startsWith('https://')) {{
      document.getElementById('login-link').href = value.login_url;
      document.getElementById('login-code').textContent = value.login_code;
    }}
    document.getElementById('login').disabled = value.state === 'running';
    document.getElementById('dot').className = 'dot ' + value.state;
    document.getElementById('status-label').textContent = op + (value.reason === 'need_login' ? '需要先登入 Microsoft，請按左側「登入 Outlook」' : (labels[value.state] || value.state));
    document.getElementById('status-detail').textContent = Object.entries(value)
      .filter(([key]) => !['state','operation','review_file','reason','login_url','login_code'].includes(key))
      .map(([key,val]) => (keys[key] || key) + ' ' + val).join(' · ');
    document.getElementById('scan').disabled = value.state === 'running';
    if (value.state === 'running') setTimeout(refresh, 1500);
    else if (value.state === 'done' && value.operation !== 'login' && !sessionStorage.getItem('reloaded-' + value.operation + JSON.stringify(value))) {{
      sessionStorage.setItem('reloaded-' + value.operation + JSON.stringify(value), '1'); location.reload();
    }}
  }} catch {{ document.getElementById('status-label').textContent='無法讀取執行狀態'; }}
}}
document.getElementById('login').onclick=async()=>{{
  try {{await post('/api/login',{{}})}} catch(e){{alert(e.message)}}
}};
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
  if(!confirm('依選擇檔保留勾選郵件，將未勾選郵件移至「刪除的郵件」？')) {{ input.value=''; return; }}
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
            elif path == "/api/login":
                started = JOBS.start_login()
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
