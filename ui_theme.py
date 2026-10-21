"""Shared stylesheet for the dashboard and review pages.

Static strings only: no mailbox data passes through this module.
"""

BASE_CSS = """
:root{--bg:#f4f6f9;--surface:#fff;--surface-2:#f8fafc;--text:#172033;--muted:#5b6678;--line:#dfe4ec;
--accent:#2563eb;--accent-soft:#e8f0ff;--keep:#15803d;--keep-soft:#e6f6ec;--review:#b45309;--review-soft:#fff3dc;
--delete:#b42318;--delete-soft:#fde8e6;--radius:14px;--shadow:0 1px 2px rgba(16,24,40,.06),0 4px 14px rgba(16,24,40,.05)}
@media (prefers-color-scheme:dark){:root{--bg:#0f141c;--surface:#171e29;--surface-2:#1c2431;--text:#e8ecf3;
--muted:#9aa6b8;--line:#2a3445;--accent:#6ea0ff;--accent-soft:#1b2a47;--keep:#4ade80;--keep-soft:#12301f;
--review:#fbbf24;--review-soft:#3a2c0e;--delete:#f87171;--delete-soft:#3b1a18;--shadow:none}}
*{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--bg);color:var(--text);line-height:1.55;
font:15px/1.55 -apple-system,BlinkMacSystemFont,"SF Pro Text","PingFang TC","Noto Sans TC","Microsoft JhengHei",system-ui,sans-serif}
.wrap{max-width:980px;margin:0 auto;padding:28px 20px 80px}
h1{font-size:26px;line-height:1.25;margin:0 0 4px}
h2{font-size:17px;margin:0 0 12px}
h3{font-size:16px;line-height:1.4;margin:0;overflow-wrap:anywhere}
p{margin:0}
a{color:var(--accent)}
code{background:var(--surface-2);border:1px solid var(--line);border-radius:6px;padding:1px 6px;font-size:13px}
.sub{color:var(--muted)}
.small{font-size:13px;color:var(--muted)}
.panel,.card{background:var(--surface);border:1px solid var(--line);border-radius:var(--radius);box-shadow:var(--shadow)}
.panel{padding:20px;margin:16px 0}
button,.btn{font:inherit;font-weight:600;border-radius:10px;border:1px solid transparent;padding:9px 16px;cursor:pointer;
background:var(--accent);color:#fff;text-decoration:none;display:inline-flex;align-items:center;gap:6px;min-height:40px}
button:hover,.btn:hover{filter:brightness(1.07)}
button:disabled{opacity:.5;cursor:not-allowed}
button.ghost,.btn.ghost{background:var(--surface);color:var(--text);border-color:var(--line)}
button.danger{background:var(--delete);color:#fff}
button.danger-ghost{background:var(--surface);color:var(--delete);border-color:var(--delete)}
:focus-visible{outline:3px solid var(--accent);outline-offset:2px}
input[type=number],input[type=search],select{font:inherit;color:var(--text);background:var(--surface);border:1px solid var(--line);
border-radius:10px;padding:9px 12px;min-height:40px}
.chip{display:inline-flex;align-items:center;gap:6px;border-radius:999px;padding:2px 10px;font-size:13px;font-weight:600;white-space:nowrap}
.chip.keep{background:var(--keep-soft);color:var(--keep)}
.chip.review{background:var(--review-soft);color:var(--review)}
.chip.delete_candidate{background:var(--delete-soft);color:var(--delete)}
.chip.total{background:var(--accent-soft);color:var(--accent)}
[hidden]{display:none!important}
@media (max-width:640px){.wrap{padding:18px 14px 60px}h1{font-size:22px}}
@media (prefers-reduced-motion:reduce){*{transition:none!important;animation:none!important}}
"""

DASHBOARD_CSS = """
.top{display:flex;justify-content:space-between;align-items:flex-start;gap:16px;flex-wrap:wrap;margin-bottom:8px}
.badge-local{display:inline-flex;gap:6px;align-items:center;background:var(--keep-soft);color:var(--keep);
border-radius:999px;padding:4px 12px;font-size:13px;font-weight:600}
.steps{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:14px;margin:16px 0}
.steps .panel{margin:0}
.step-no{display:inline-grid;place-items:center;width:24px;height:24px;border-radius:50%;background:var(--accent-soft);
color:var(--accent);font-weight:700;font-size:13px;margin-right:8px}
.fields{display:flex;gap:16px;flex-wrap:wrap;align-items:flex-end;margin:12px 0}
.fields label{display:flex;flex-direction:column;gap:4px;font-size:13px;color:var(--muted)}
.fields input{width:110px}
#status-box{display:flex;gap:12px;align-items:center}
.dot{width:12px;height:12px;border-radius:50%;background:var(--muted);flex:none}
.dot.running{background:var(--accent);animation:pulse 1.1s ease-in-out infinite}
.dot.done{background:var(--keep)}.dot.failed{background:var(--delete)}
@keyframes pulse{50%{opacity:.25}}
#status-detail{color:var(--muted);font-size:13px}
.review-row{padding:18px 20px;margin:12px 0;display:flex;gap:16px;justify-content:space-between;align-items:center;flex-wrap:wrap}
.review-row .meta{display:flex;flex-direction:column;gap:8px;min-width:0}
.review-row .name{font-weight:650;font-variant-numeric:tabular-nums}
.chips{display:flex;gap:8px;flex-wrap:wrap}
.actions{display:flex;gap:10px;align-items:center;flex-wrap:wrap}
.file-btn{position:relative;overflow:hidden}
.file-btn input{position:absolute;inset:0;opacity:0;cursor:pointer;font-size:0}
.code{font:700 30px/1.2 ui-monospace,SFMono-Regular,Menlo,monospace;letter-spacing:.15em;margin:8px 0;padding:12px 16px;background:var(--accent-soft);color:var(--accent);border-radius:10px;text-align:center;user-select:all}
.empty-state{text-align:center;padding:36px 16px;color:var(--muted)}
"""

REVIEW_CSS = """
.hero{display:flex;justify-content:space-between;gap:12px;align-items:flex-end;flex-wrap:wrap;margin-bottom:14px}
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin:14px 0}
.stat{padding:14px 16px;border:1px solid var(--line);border-radius:var(--radius);background:var(--surface);
text-align:left;cursor:pointer;color:var(--text);font-weight:400;display:block;box-shadow:var(--shadow)}
.stat b{display:block;font-size:26px;line-height:1.2;font-variant-numeric:tabular-nums}
.stat span{font-size:13px;color:var(--muted)}
.stat[aria-pressed=true]{outline:2px solid var(--accent)}
.stat.keep b{color:var(--keep)}.stat.review b{color:var(--review)}.stat.delete_candidate b{color:var(--delete)}
.toolbar{position:sticky;top:0;z-index:5;background:var(--bg);padding:10px 0;border-bottom:1px solid var(--line);
display:flex;gap:10px;flex-wrap:wrap;align-items:center}
.toolbar input[type=search]{flex:1 1 220px;min-width:0}
#result-count{color:var(--muted);font-size:13px}
.summary{margin:14px 0;padding:14px 16px;border-radius:var(--radius);border:1px solid var(--line);background:var(--surface);
display:flex;gap:12px;flex-wrap:wrap;align-items:center}
.summary strong{font-size:15px}
.summary .grow{flex:1 1 280px}
.notice{color:var(--muted);font-size:13px;flex-basis:100%}
.group{margin:28px 0 8px}
.group h2{display:flex;align-items:center;gap:10px}
article{position:relative;background:var(--surface);border:1px solid var(--line);border-left:5px solid var(--line);
border-radius:var(--radius);padding:14px 16px;margin:10px 0;display:grid;grid-template-columns:auto 1fr;gap:4px 14px;
box-shadow:var(--shadow)}
article[data-decision=keep]{border-left-color:var(--keep)}
article[data-decision=review]{border-left-color:var(--review)}
article[data-decision=delete_candidate]{border-left-color:var(--delete)}
article>*:not(.pick){grid-column:2}
.pick{grid-column:1;grid-row:1 / span 4;display:flex;align-items:flex-start;padding-top:2px}
.pick input{width:22px;height:22px;accent-color:var(--keep);cursor:pointer}
.pick input:disabled{cursor:not-allowed}
.line1{display:flex;gap:10px;justify-content:space-between;align-items:flex-start;flex-wrap:wrap}
.who{color:var(--muted);font-size:13px;overflow-wrap:anywhere}
.why{font-size:13px;color:var(--muted)}
.fate{font-size:12px;font-weight:700;border-radius:999px;padding:2px 10px;white-space:nowrap}
.fate.keeps{background:var(--keep-soft);color:var(--keep)}
.fate.trash{background:var(--delete-soft);color:var(--delete)}
.fate.moved,.fate.locked{background:var(--accent-soft);color:var(--accent)}
.empty{color:var(--muted)}
#toast{position:fixed;left:50%;bottom:24px;transform:translateX(-50%);background:var(--text);color:var(--bg);
padding:10px 18px;border-radius:999px;font-weight:600;box-shadow:0 8px 24px rgba(0,0,0,.25)}
"""
