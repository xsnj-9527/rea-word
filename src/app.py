# -*- coding: utf-8 -*-
"""DocPreview - a small read-only .docx viewer with dark mode.

The Word installation is never touched: this program only reads .docx files,
which are ordinary zip archives, and renders them with the system WebView2.
"""

import json
import os
import sys

import webview

from docx2html import convert

APP_TITLE = "DocPreview"
APP_ID = "DocPreview"
VERSION = "1.0.1"


def _base_dir():
    """Directory holding bundled resources.

    PyInstaller unpacks a one-file build into a temp folder recorded in
    sys._MEIPASS; running from source uses this file's directory.
    """
    return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))


def _icon_path():
    for name in ("icon.ico", "icon.png"):
        p = os.path.join(_base_dir(), name)
        if os.path.isfile(p):
            return p
    return None

# The whole UI is one HTML document; Python injects the converted body.
SHELL = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>DocPreview</title>
<style>
  :root {
    --bg: #1e1e1e;
    --bg-elev: #252526;
    --fg: #e6e6e6;
    --fg-dim: #9d9d9d;
    --accent: #4a9eff;
    --border: #333;
    --paper: #1e1e1e;
    --paper-fg: #e6e6e6;
    --paper-shadow: 0 2px 24px rgba(0,0,0,.55);
  }
  html[data-theme="light"] {
    --bg: #f3f3f3;
    --bg-elev: #fafafa;
    --fg: #1f1f1f;
    --fg-dim: #666;
    --accent: #0b6bcb;
    --border: #ddd;
    --paper: #ffffff;
    --paper-fg: #1a1a1a;
    --paper-shadow: 0 2px 20px rgba(0,0,0,.12);
  }
  * { box-sizing: border-box; }
  html, body { height: 100%; margin: 0; }
  body {
    background: var(--bg);
    color: var(--fg);
    font: 14px/1.5 "Segoe UI", "Microsoft YaHei", system-ui, sans-serif;
    display: flex;
    flex-direction: column;
    overflow: hidden;
  }

  /* ---- toolbar ---- */
  header {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 8px 14px;
    background: var(--bg-elev);
    border-bottom: 1px solid var(--border);
    flex: 0 0 auto;
    user-select: none;
  }
  header .title {
    font-weight: 600;
    font-size: 13px;
    margin-right: 6px;
    white-space: nowrap;
  }
  header .fname {
    color: var(--fg-dim);
    font-size: 12px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    flex: 1 1 auto;
  }
  button {
    font: inherit;
    font-size: 12px;
    color: var(--fg);
    background: transparent;
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 5px 11px;
    cursor: pointer;
    white-space: nowrap;
  }
  button:hover { border-color: var(--accent); color: var(--accent); }
  button:active { transform: translateY(1px); }
  .spacer { flex: 1 1 auto; }

  /* ---- reading area ---- */
  main {
    flex: 1 1 auto;
    overflow-y: auto;
    padding: 26px 20px 70px;
    scroll-behavior: smooth;
  }
  #paper {
    max-width: 780px;
    margin: 0 auto;
    background: var(--paper);
    color: var(--paper-fg);
    padding: 56px 64px 72px;
    border-radius: 3px;
    box-shadow: var(--paper-shadow);
    min-height: 60vh;
    transition: background .18s ease, color .18s ease;
  }
  #paper.empty {
    display: flex;
    align-items: center;
    justify-content: center;
    color: var(--fg-dim);
    font-size: 14px;
    text-align: center;
    box-shadow: none;
    background: transparent;
    border: 1px dashed var(--border);
  }

  /* ---- document typography ---- */
  #paper h1, #paper h2, #paper h3, #paper h4, #paper h5, #paper h6 {
    line-height: 1.3;
    margin: 1.4em 0 .55em;
    font-weight: 650;
  }
  #paper h1 { font-size: 1.85em; margin-top: .2em; }
  #paper h2 { font-size: 1.5em; }
  #paper h3 { font-size: 1.25em; }
  #paper h4 { font-size: 1.1em; }
  #paper h5, #paper h6 { font-size: 1em; }
  #paper p { margin: 0 0 .72em; }
  #paper p.empty { margin: 0 0 .35em; }
  #paper p.li { padding-left: 1.4em; position: relative; }
  #paper p.li::before {
    content: "•";
    position: absolute;
    left: .45em;
    color: var(--fg-dim);
  }
  #paper .doc-table {
    border-collapse: collapse;
    margin: .9em 0 1.2em;
    width: 100%;
    font-size: .95em;
  }
  #paper .doc-table td {
    border: 1px solid var(--border);
    padding: 6px 10px;
    vertical-align: top;
  }
  #paper .doc-table tr:first-child td { background: rgba(127,127,127,.10); font-weight: 600; }
  #paper sup, #paper sub { font-size: .72em; }
  #paper img { max-width: 100%; height: auto; }
  #paper a { color: var(--accent); }

  /* Text must stay selectable so the reader can copy quotes out. */
  #paper, #paper * { user-select: text; -webkit-user-select: text; }
  #paper ::selection { background: rgba(74,158,255,.35); }

  /* Word lets authors set near-black text, which vanishes on a dark page.
     The converter tags those runs so the theme can flip them. */
  html[data-theme="dark"] #paper .dark-ink { color: #e6e6e6; }
  html[data-theme="light"] #paper .dark-ink { color: #1a1a1a; }

  /* ---- figures & lightbox ---- */
  #paper p.fig {
    margin: 1.1em 0;
    text-align: center;
  }
  #paper img.doc-img {
    cursor: zoom-in;
    max-width: 100%;
    border-radius: 3px;
    transition: opacity .15s ease;
  }
  #paper img.doc-img:hover { opacity: .88; }
  #lightbox {
    position: fixed;
    inset: 0;
    display: none;
    align-items: center;
    justify-content: center;
    background: rgba(0,0,0,.88);
    z-index: 200;
    cursor: zoom-out;
    padding: 24px;
  }
  #lightbox.on { display: flex; }
  #lightbox img {
    max-width: 100%;
    max-height: 100%;
    object-fit: contain;
    box-shadow: 0 8px 48px rgba(0,0,0,.7);
    border-radius: 4px;
  }

  /* ---- status bar ---- */
  footer {
    flex: 0 0 auto;
    display: flex;
    align-items: center;
    gap: 14px;
    padding: 6px 14px;
    background: var(--bg-elev);
    border-top: 1px solid var(--border);
    font-size: 11.5px;
    color: var(--fg-dim);
    user-select: none;
  }
  #warn { color: #e0a030; cursor: help; }

  /* ---- drop overlay ---- */
  #drop {
    position: fixed;
    inset: 0;
    display: none;
    align-items: center;
    justify-content: center;
    background: rgba(0,0,0,.6);
    color: #fff;
    font-size: 18px;
    z-index: 99;
    pointer-events: none;
  }
  #drop.on { display: flex; }
</style>
</head>
<body>
<header>
  <span class="title">DocPreview</span>
  <button id="btn-open">打开文档</button>
  <button id="btn-theme">切换主题</button>
  <button id="btn-zoom-out" title="缩小">−</button>
  <button id="btn-zoom-in" title="放大">+</button>
  <span class="fname" id="fname">未打开文件</span>
</header>

<main id="main">
  <div id="paper" class="empty">把 .docx 文件拖到这里，或点击「打开文档」</div>
</main>

<footer>
  <span id="stat">就绪</span>
  <span class="spacer"></span>
  <span id="warn"></span>
</footer>

<div id="drop">松开以打开文档</div>
<div id="lightbox"><img id="lightbox-img" alt=""></div>

<script>
const paper  = document.getElementById('paper');
const fname  = document.getElementById('fname');
const stat   = document.getElementById('stat');
const warnEl = document.getElementById('warn');
let zoom = 1;
let currentPath = null;

function applyTheme(t) {
  document.documentElement.dataset.theme = t;
  try { localStorage.setItem('dp-theme', t); } catch (e) {}
}
applyTheme((() => { try { return localStorage.getItem('dp-theme') || 'dark'; } catch(e){ return 'dark'; } })());

document.getElementById('btn-theme').onclick = () => {
  applyTheme(document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark');
};
function setZoom(z) {
  zoom = Math.min(2.5, Math.max(0.5, z));
  // CSS zoom scales the whole paper, including the pt-based font sizes that
  // Word emits. Changing font-size alone would only reflow the text.
  paper.style.zoom = zoom;
  stat.textContent = '缩放 ' + Math.round(zoom * 100) + '%';
}
document.getElementById('btn-zoom-in').onclick  = () => setZoom(zoom + 0.1);
document.getElementById('btn-zoom-out').onclick = () => setZoom(zoom - 0.1);

document.getElementById('btn-open').onclick = () => window.pywebview.api.open_dialog();

// Called from Python after a document is converted.
window.renderDoc = function (payload) {
  const data = typeof payload === 'string' ? JSON.parse(payload) : payload;
  currentPath = data.path || null;
  fname.textContent = data.name || '未打开文件';
  if (data.error) {
    paper.className = 'empty';
    paper.textContent = data.error;
    stat.textContent = '打开失败';
    warnEl.textContent = '';
    return;
  }
  paper.className = '';
  paper.innerHTML = data.html || '<p class="empty"></p>';
  const n = paper.querySelectorAll('p,h1,h2,h3,h4,h5,h6,table').length;
  stat.textContent = n + ' 个段落 · ' + (data.words || 0) + ' 字';
  warnEl.textContent = (data.warnings && data.warnings.length)
      ? '⚠ ' + data.warnings.length + ' 项限制'
      : '';
  warnEl.title = (data.warnings || []).join('\n');
  paper.scrollTop = 0;
  document.getElementById('main').scrollTop = 0;
};

// ---- image lightbox ----
const lightbox = document.getElementById('lightbox');
const lightboxImg = document.getElementById('lightbox-img');
paper.addEventListener('click', e => {
  const img = e.target.closest('img.doc-img');
  if (!img) return;
  lightboxImg.src = img.src;
  lightbox.classList.add('on');
});
lightbox.addEventListener('click', () => {
  lightbox.classList.remove('on');
  lightboxImg.src = '';
});

// ---- drag & drop ----
let dragDepth = 0;
const dropEl = document.getElementById('drop');
window.addEventListener('dragenter', e => { e.preventDefault(); dragDepth++; dropEl.classList.add('on'); });
window.addEventListener('dragover',  e => { e.preventDefault(); });
window.addEventListener('dragleave', e => { e.preventDefault(); if (--dragDepth <= 0) { dragDepth = 0; dropEl.classList.remove('on'); } });
window.addEventListener('drop', e => {
  e.preventDefault(); dragDepth = 0; dropEl.classList.remove('on');
  const f = e.dataTransfer.files && e.dataTransfer.files[0];
  if (f && window.pywebview && window.pywebview.api) {
    window.pywebview.api.open_path(f.name);
  }
});

// ---- keyboard ----
document.addEventListener('keydown', e => {
  if (e.key === 'Escape' && lightbox.classList.contains('on')) {
    lightbox.classList.remove('on'); lightboxImg.src = ''; return;
  }
  if (e.ctrlKey && e.key === 'o') { e.preventDefault(); window.pywebview.api.open_dialog(); }
  if (e.ctrlKey && (e.key === '=' || e.key === '+')) { e.preventDefault(); setZoom(zoom + 0.1); }
  if (e.ctrlKey && e.key === '-') { e.preventDefault(); setZoom(zoom - 0.1); }
  if (e.ctrlKey && e.key === '0') { e.preventDefault(); setZoom(1); }
});
</script>
</body>
</html>
"""


def _load_document(path):
    """Convert one .docx and return the JSON-serialisable payload."""
    name = os.path.basename(path)
    try:
        html, warnings, meta = convert(path)
    except Exception as exc:  # noqa: BLE001 - surfaced to the UI verbatim
        return {"error": "无法读取：%s" % exc, "name": name, "path": path}

    # Rough word count so the status bar is informative without being exact.
    text = "".join(html.split("<"))
    words = len([w for w in text.replace("\n", " ").split() if w.strip()])

    payload = {
        "name": name,
        "path": path,
        "html": html,
        "warnings": warnings,
        "words": words,
    }
    if meta.get("title"):
        payload["name"] = "%s — %s" % (meta["title"], name)
    return payload


class Api:
    """Methods reachable from JavaScript as window.pywebview.api.*"""

    def __init__(self):
        self.window = None

    def open_dialog(self):
        win = self.window
        if win is None:
            return
        result = win.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Word 文档 (*.docx;*.docm)", "所有文件 (*.*)"),
        )
        if result:
            self._show(result[0])

    def open_path(self, path):
        """Drag-and-drop entry point: the browser only gives us a filename."""
        if os.path.isfile(path):
            self._show(path)
            return
        # WebView2 hands us just the basename; look for it in common folders.
        for base in (os.path.expanduser("~/Desktop"),
                     os.path.expanduser("~/Documents"),
                     os.path.expanduser("~/Downloads"),
                     os.getcwd()):
            cand = os.path.join(base, path)
            if os.path.isfile(cand):
                self._show(cand)
                return
        self._push({"error": "找不到文件：%s（请用「打开文档」按钮）" % path,
                    "name": path})

    def _show(self, path):
        self._push(_load_document(path))

    def _push(self, payload):
        if self.window is None:
            return
        js = "window.renderDoc(%s)" % json.dumps(json.dumps(payload, ensure_ascii=False))
        self.window.evaluate_js(js)


def main():
    api = Api()

    # Accept a path on the command line so the app can be used as a
    # file-association target ("Open with" passes the file as argv[1]).
    initial = None
    if len(sys.argv) > 1 and os.path.isfile(sys.argv[1]):
        initial = sys.argv[1]

    window = webview.create_window(
        APP_TITLE,
        html=SHELL,
        width=1040,
        height=820,
        min_size=(560, 420),
        background_color="#1e1e1e",
        js_api=api,
    )
    api.window = window

    def on_loaded():
        if initial:
            api._show(initial)

    window.events.loaded += on_loaded

    icon = _icon_path()
    webview.start(debug=False, icon=icon) if icon else webview.start(debug=False)


if __name__ == "__main__":
    main()
