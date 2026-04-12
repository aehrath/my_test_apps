#!/usr/bin/env python3
"""CSV Editor — browser-based, zero external dependencies.

Usage:
  python3 csv_editor.py              # start with blank sheet
  python3 csv_editor.py data.csv     # open a file directly
"""

import base64
import csv
import io
import json
import sys
import threading
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

try:
    import openpyxl
    _XLSX_OK = True
except ImportError:
    _XLSX_OK = False


# ── Shared state ──────────────────────────────────────────────────────────────

class State:
    def __init__(self):
        self.sheets       = [{'name': 'Sheet1', 'headers': [], 'rows': []}]
        self.active_sheet = 0
        self.filepath     = None
        self.filetype     = 'csv'

    @property
    def headers(self):
        return self.sheets[self.active_sheet]['headers']

    @headers.setter
    def headers(self, v):
        self.sheets[self.active_sheet]['headers'] = v

    @property
    def rows(self):
        return self.sheets[self.active_sheet]['rows']

    @rows.setter
    def rows(self, v):
        self.sheets[self.active_sheet]['rows'] = v

state = State()


# ── Excel helpers ─────────────────────────────────────────────────────────────

import datetime as _dt
import re as _re

def _cell_to_str(value):
    """Convert an openpyxl cell value to a clean string."""
    if value is None:
        return ''
    if isinstance(value, bool):
        return 'TRUE' if value else 'FALSE'
    if isinstance(value, float):
        # Avoid ugly "984.0" for whole numbers
        return str(int(value)) if value == int(value) else str(value)
    if isinstance(value, _dt.datetime):
        # Strip the time portion when it is midnight (date-only cells)
        if value.hour == 0 and value.minute == 0 and value.second == 0:
            return value.strftime('%Y-%m-%d')
        return value.strftime('%Y-%m-%d %H:%M:%S')
    if isinstance(value, _dt.date):
        return value.strftime('%Y-%m-%d')
    return str(value)

_INT_RE  = _re.compile(r'^-?\d+$')
_FLOAT_RE = _re.compile(r'^-?\d+\.?\d*([eE][+-]?\d+)?$')

def _str_to_cell(s):
    """Convert a string cell value back to an appropriate Python type for Excel."""
    if not isinstance(s, str):
        return s
    if _INT_RE.match(s):
        return int(s)
    if _FLOAT_RE.match(s):
        return float(s)
    return s


def _parse_xlsx(raw_bytes):
    """Parse an .xlsx binary into (headers, rows). Requires openpyxl."""
    if not _XLSX_OK:
        raise RuntimeError('openpyxl not installed — run: pip install openpyxl')
    wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
    ws = wb.active
    if ws is None:
        # Workbook has no active sheet — try first sheet
        if not wb.sheetnames:
            return [], []
        ws = wb[wb.sheetnames[0]]
    all_rows = []
    for row in ws.iter_rows(values_only=False):
        cells = [_cell_to_str(cell.value) for cell in row]
        # Skip rows that are entirely blank
        if any(v != '' for v in cells):
            all_rows.append(cells)
    if not all_rows:
        return [], []
    n_cols = max(len(r) for r in all_rows)
    for r in all_rows:
        while len(r) < n_cols:
            r.append('')
    return all_rows[0], all_rows[1:]


def _write_xlsx(headers, rows):
    """Serialise headers + rows to .xlsx bytes. Requires openpyxl.
    Numeric-looking strings are written as numbers so Excel treats them correctly."""
    if not _XLSX_OK:
        raise RuntimeError('openpyxl not installed — run: pip install openpyxl')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(headers)  # headers always as strings
    for row in rows:
        ws.append([_str_to_cell(c) for c in row])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── GitHub state & helpers ────────────────────────────────────────────────────

class GitHubConfig:
    def __init__(self):
        self.token  = ''
        self.repo   = ''    # owner/repo
        self.branch = 'main'
        self.path   = ''    # file path inside repo

gh = GitHubConfig()

_GH_CONFIG_FILE = Path.home() / '.csv_editor_github.json'

def load_gh_config():
    """Load persisted GitHub config from disk, if it exists."""
    try:
        data = json.loads(_GH_CONFIG_FILE.read_text(encoding='utf-8'))
        gh.token  = data.get('token',  '')
        gh.repo   = data.get('repo',   '')
        gh.branch = data.get('branch', 'main') or 'main'
        gh.path   = data.get('path',   '')
    except (FileNotFoundError, json.JSONDecodeError):
        pass

def save_gh_config():
    """Persist current GitHub config to disk."""
    _GH_CONFIG_FILE.write_text(
        json.dumps({'token': gh.token, 'repo': gh.repo, 'branch': gh.branch, 'path': gh.path}),
        encoding='utf-8',
    )


def _gh_api(method, endpoint, data=None):
    """Call GitHub REST API. Returns (result_dict, error_str)."""
    url = 'https://api.github.com' + endpoint
    hdrs = {
        'Authorization': f'token {gh.token}',
        'Accept': 'application/vnd.github.v3+json',
        'User-Agent': 'CSVEditor/1.0',
    }
    body = json.dumps(data).encode() if data is not None else None
    if body:
        hdrs['Content-Type'] = 'application/json'
    req = urllib.request.Request(url, data=body, headers=hdrs, method=method)
    try:
        with urllib.request.urlopen(req) as r:
            return json.loads(r.read()), None
    except urllib.error.HTTPError as e:
        try:
            msg = json.loads(e.read()).get('message', str(e))
        except Exception:
            msg = str(e)
        return None, msg
    except Exception as e:
        return None, str(e)


# ── Embedded HTML/CSS/JS ──────────────────────────────────────────────────────

HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>CSV Editor</title>
<script src="https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js"></script>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    font-family: Arial, sans-serif;
    font-size: 13px;
    display: flex;
    flex-direction: column;
    height: 100vh;
    overflow: hidden;
    background: #f0f2f5;
    color: #222;
  }

  /* ── Toolbar ── */
  #toolbar {
    display: flex;
    align-items: center;
    gap: 4px;
    padding: 6px 10px;
    background: #2c3e50;
    color: #fff;
    flex-shrink: 0;
    flex-wrap: wrap;
  }
  #toolbar button {
    background: #3d5166;
    color: #fff;
    border: 1px solid #4a6278;
    border-radius: 4px;
    padding: 4px 10px;
    cursor: pointer;
    font-size: 12px;
    white-space: nowrap;
  }
  #toolbar button:hover { background: #4e6a84; }
  #toolbar button:active { background: #5a7a96; }

  .sep { width: 1px; height: 24px; background: #4a6278; margin: 0 4px; }

  #search-input {
    padding: 4px 8px;
    border: 1px solid #4a6278;
    border-radius: 4px;
    background: #3d5166;
    color: #fff;
    font-size: 12px;
    width: 160px;
  }
  #search-input::placeholder { color: #8aa; }
  #search-input:focus { outline: none; background: #46607a; }

  #filepath-display {
    font-size: 11px;
    color: #aac;
    margin-left: auto;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    max-width: 300px;
  }

  /* ── Table container ── */
  #table-wrap {
    flex: 1;
    overflow: auto;
    background: #fff;
    margin: 6px;
    border-radius: 6px;
    box-shadow: 0 1px 4px rgba(0,0,0,.15);
  }

  table {
    border-collapse: collapse;
    min-width: 100%;
    table-layout: fixed;
  }

  /* Header */
  thead th {
    position: sticky;
    top: 0;
    z-index: 2;
    background: #dce8f7;
    border: 1px solid #b8cfe8;
    padding: 6px 8px;
    font-weight: 700;
    text-align: left;
    cursor: pointer;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    user-select: none;
    min-width: 100px;
  }
  thead th:hover { background: #c8d8ee; }
  thead th.row-num { min-width: 42px; width: 42px; cursor: default; text-align: center; }

  /* Rows */
  tbody tr.even { background: #fff; }
  tbody tr.odd  { background: #f6f9ff; }
  tbody tr:hover { background: #edf3fc; }
  tbody tr.selected-row { background: #d6e8ff !important; }

  thead th.selected-col { background: #a8c4e8; }
  tbody td.selected-col { background: rgba(100,160,255,.10); }
  tbody tr.selected-row td.selected-col { background: rgba(80,140,255,.30) !important; }

  tbody td {
    border: 1px solid #dde;
    padding: 0;
    height: 28px;
  }
  tbody td.row-num {
    text-align: center;
    color: #888;
    font-size: 11px;
    background: #f0f0f0;
    border-right: 2px solid #ccc;
    cursor: pointer;
    user-select: none;
  }
  tbody td.row-num:hover { background: #e0e0e0; }

  tbody td input {
    width: 100%;
    height: 100%;
    border: none;
    background: transparent;
    padding: 4px 8px;
    font-family: inherit;
    font-size: 13px;
    color: #222;
    outline: none;
    display: block;
  }
  tbody td input:focus {
    background: #e8f0fe;
    outline: 2px solid #4a90d9;
    outline-offset: -2px;
    border-radius: 2px;
  }
  tbody td.highlight input { background: #fff3a0; }
  tbody td.highlight input:focus { background: #ffe060; }

  /* cells that differ from the last-committed / last-loaded baseline */
  tbody td.cell-modified { box-shadow: inset 3px 0 0 #e6a817; background: #fffae8; }
  tbody td.cell-modified input { background: transparent; }
  tbody td.cell-modified input:focus { background: #e8f0fe; }

  /* ── Inline diff ── */
  #diff-banner { display:none; align-items:center; gap:10px; padding:5px 12px;
    background:#1e3a5f; color:#9fc8f0; font-size:12px; flex-shrink:0; }
  #diff-banner.open { display:flex; }
  #diff-banner strong { color:#7ab8e8; font-family:monospace; }
  #diff-banner .diff-stats { color:#aad; }
  #diff-banner button { margin-left:auto; background:transparent; border:1px solid #4a6278;
    color:#9fc8f0; border-radius:3px; padding:1px 10px; cursor:pointer; font-size:11px; }
  #diff-banner button:hover { background:#2a4a6e; }
  tbody tr.row-diff-added > td { background:#efffef !important; }
  tbody tr.row-diff-removed > td { background:#f8f8f8 !important; color:#bbb; pointer-events:none; }
  tbody tr.row-diff-removed > td.row-num { color:#ddd; }
  tbody td.cell-diff-changed { background:#fffbe6 !important; box-shadow:inset 3px 0 0 #f0c000; }
  tbody td.cell-diff-changed input { background:transparent; color:inherit; }
  tbody td.cell-diff-changed input:focus { background:#fff8d6; }
  tbody td.cell-diff-changed .old-val {
    display:block; font-size:10px; color:#aaa; text-decoration:line-through;
    white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
    line-height:1.2; margin-bottom:1px; pointer-events:none;
  }

  /* ── Status bar ── */
  #statusbar {
    flex-shrink: 0;
    padding: 3px 12px;
    background: #2c3e50;
    color: #aac;
    font-size: 11px;
    display: flex;
    justify-content: space-between;
  }

  /* ── GitHub panel ── */
  #gh-panel {
    position: fixed; top: 0; right: -380px; bottom: 0; width: 380px;
    background: #1a2535; color: #c8d8e8;
    display: flex; flex-direction: column;
    box-shadow: -4px 0 20px rgba(0,0,0,.4);
    transition: right .22s ease; z-index: 60;
    font-size: 12px;
  }
  #gh-panel.open { right: 0; }
  #gh-panel h2 { font-size: 13px; padding: 12px 14px 8px; background: #111d2b;
    border-bottom: 1px solid #2a3d52; display: flex; align-items: center; gap: 8px; }
  #gh-panel h2 span { flex: 1; }
  #gh-panel h2 button { background: none; border: none; color: #888; cursor: pointer; font-size: 16px; }
  #gh-panel h2 button:hover { color: #fff; }

  .gh-tabs { display: flex; border-bottom: 1px solid #2a3d52; }
  .gh-tab { flex: 1; padding: 7px; text-align: center; cursor: pointer;
    color: #778899; border-bottom: 2px solid transparent; }
  .gh-tab.active { color: #7ab8e8; border-bottom-color: #7ab8e8; }

  .gh-body { flex: 1; overflow-y: auto; padding: 12px; }

  .gh-field { margin-bottom: 10px; }
  .gh-field label { display: block; font-size: 11px; color: #778899; margin-bottom: 3px; }
  .gh-field input { width: 100%; padding: 5px 8px; background: #243447;
    border: 1px solid #2e4460; border-radius: 4px; color: #c8d8e8; font-size: 12px; }
  .gh-field input:focus { outline: none; border-color: #4a90d9; }

  .gh-btn { display: block; width: 100%; padding: 7px; margin-top: 8px;
    background: #2d6a9f; color: #fff; border: none; border-radius: 4px;
    cursor: pointer; font-size: 12px; }
  .gh-btn:hover { background: #3a7db5; }
  .gh-btn.danger { background: #7a2d2d; }
  .gh-btn.danger:hover { background: #9a3d3d; }
  .gh-btn:disabled { background: #2a3d52; color: #556677; cursor: default; }

  #gh-status { font-size: 11px; padding: 4px 0; }
  #gh-status.ok  { color: #5cb85c; }
  #gh-status.err { color: #e07070; }

  .commit-item { padding: 8px 0; border-bottom: 1px solid #2a3d52; }
  .commit-item:last-child { border: none; }
  .commit-sha  { font-family: monospace; font-size: 11px; color: #7ab8e8; }
  .commit-msg  { font-weight: 600; margin: 2px 0; line-height: 1.3; }
  .commit-meta { font-size: 10px; color: #556677; }
  .commit-actions { display: flex; gap: 5px; margin-top: 5px; }
  .commit-actions button { flex: 1; padding: 3px 6px; font-size: 11px;
    border: 1px solid #2e4460; background: #243447; color: #c8d8e8;
    border-radius: 3px; cursor: pointer; }
  .commit-actions button:hover { background: #2d5270; }

  /* ── Diff overlay ── */
  #diff-overlay {
    display: none; position: fixed; inset: 0; z-index: 100;
    flex-direction: column; background: #f0f2f5;
  }
  #diff-overlay.open { display: flex; }

  #diff-toolbar {
    display: flex; align-items: center; gap: 10px;
    padding: 8px 14px; background: #2c3e50; color: #fff; flex-shrink: 0;
  }
  #diff-toolbar h3 { flex: 1; font-size: 13px; font-weight: 600; }
  #diff-toolbar button { background: #3d5166; color: #fff; border: 1px solid #4a6278;
    border-radius: 4px; padding: 4px 12px; cursor: pointer; font-size: 12px; }
  #diff-toolbar button:hover { background: #4e6a84; }

  #diff-legend { display: flex; gap: 16px; padding: 6px 14px;
    background: #e8edf2; border-bottom: 1px solid #ccc; font-size: 11px; flex-shrink: 0; }
  .legend-dot { display: inline-block; width: 12px; height: 12px;
    border-radius: 2px; margin-right: 4px; vertical-align: middle; }

  #diff-wrap { flex: 1; overflow: auto; }

  #diff-table { border-collapse: collapse; min-width: 100%; font-size: 12px; }
  #diff-table th { position: sticky; top: 0; background: #dce8f7;
    border: 1px solid #b8cfe8; padding: 5px 8px; font-weight: 700;
    text-align: left; white-space: nowrap; }
  #diff-table td { border: 1px solid #dde; padding: 4px 8px;
    white-space: pre-wrap; max-width: 300px; overflow: hidden;
    text-overflow: ellipsis; }
  #diff-table td.diff-marker { font-weight: 700; text-align: center;
    width: 28px; font-family: monospace; }

  tr.diff-added   td                { background: #d4edda; }
  tr.diff-added   td.diff-marker    { background: #b8ddc2; color: #1a7a1a; }
  tr.diff-removed td                { background: #fde0e0; text-decoration: line-through; color: #888; }
  tr.diff-removed td.diff-marker    { background: #f5b8b8; color: #a00; }
  tr.diff-modified-old td           { background: #fde0e0; color: #888; }
  tr.diff-modified-old td.diff-marker { background: #f5b8b8; color: #a00; }
  tr.diff-modified-new td           { background: #d4edda; }
  tr.diff-modified-new td.diff-marker { background: #b8ddc2; color: #1a7a1a; }
  td.cell-changed                   { background: #fff3a0 !important; font-weight: 600; color: #333 !important; }

  /* ── Modal ── */
  #modal-overlay {
    display: none;
    position: fixed; inset: 0;
    background: rgba(0,0,0,.45);
    z-index: 100;
    align-items: center;
    justify-content: center;
  }
  #modal-overlay.open { display: flex; }
  #modal-box {
    background: #fff;
    border-radius: 8px;
    padding: 24px;
    min-width: 320px;
    box-shadow: 0 8px 32px rgba(0,0,0,.25);
  }
  #modal-box h3 { margin-bottom: 12px; }
  #modal-box input[type=text] {
    width: 100%;
    padding: 7px 10px;
    border: 1px solid #ccc;
    border-radius: 4px;
    font-size: 13px;
    margin-bottom: 14px;
  }
  #modal-box input[type=text]:focus { outline: 2px solid #4a90d9; }
  .modal-btns { display: flex; gap: 8px; justify-content: flex-end; }
  .modal-btns button {
    padding: 6px 16px;
    border-radius: 4px;
    border: 1px solid #ccc;
    cursor: pointer;
    font-size: 13px;
  }
  .modal-btns button.primary {
    background: #2c3e50;
    color: #fff;
    border-color: #2c3e50;
  }
  .modal-btns button.primary:hover { background: #3d5166; }

/* ── Sheet tabs ── */
#sheet-tabs {
  display: flex;
  align-items: center;
  gap: 2px;
  padding: 3px 8px 0;
  background: #e8eaed;
  border-top: 1px solid #ccc;
  flex-shrink: 0;
  overflow-x: auto;
  min-height: 30px;
}
.sheet-tab {
  padding: 4px 14px;
  background: #d0d3d8;
  border: 1px solid #bbb;
  border-bottom: none;
  border-radius: 4px 4px 0 0;
  cursor: pointer;
  font-size: 12px;
  white-space: nowrap;
  user-select: none;
  color: #444;
}
.sheet-tab:hover { background: #e2e5ea; }
.sheet-tab.active { background: #fff; color: #111; font-weight: 600; border-color: #aaa; }
.sheet-tab-add {
  padding: 3px 10px;
  background: none;
  border: 1px dashed #bbb;
  border-radius: 4px 4px 0 0;
  cursor: pointer;
  font-size: 14px;
  color: #888;
  line-height: 1;
}
.sheet-tab-add:hover { background: #d8dbe0; color: #333; }
</style>
</head>
<body>

<div id="toolbar">
  <button onclick="openFile()">📂 Open</button>
  <button onclick="saveFile()">💾 Save</button>
  <button onclick="saveAs()">Save As…</button>
  <button onclick="downloadFile('csv')"  title="Download as CSV">⬇ CSV</button>
  <button onclick="downloadFile('xlsx')" title="Download as Excel">⬇ XLSX</button>
  <div class="sep"></div>
  <button onclick="addRow()">+ Row</button>
  <button onclick="deleteSelectedRow()">− Row</button>
  <button onclick="addColumn()">+ Col</button>
  <button onclick="deleteColumn()">− Col</button>
  <button onclick="renameColumn()">Rename Col</button>
  <div class="sep"></div>
  <button onclick="autoResize()" title="Fit each column to its widest content">⇔ Auto Resize</button>
  <div class="sep"></div>
  <input id="search-input" type="text" placeholder="🔍 Search…" oninput="applySearch(this.value)">
  <button onclick="clearSearch()">✕</button>
  <span id="filepath-display">Untitled</span>
  <div class="sep"></div>
  <button onclick="toggleGHPanel()" title="GitHub versioning">⎇ GitHub</button>
</div>

<div id="diff-banner">
  <span>Comparing with <strong id="diff-banner-label"></strong></span>
  <span class="diff-stats" id="diff-banner-stats"></span>
  <button onclick="clearInlineDiff()">✕ Clear diff</button>
</div>

<div id="table-wrap">
  <table id="csv-table">
    <thead><tr id="header-row"></tr></thead>
    <tbody id="table-body"></tbody>
  </table>
</div>

<div id="sheet-tabs">
  <button class="sheet-tab-add" onclick="addSheet()" title="Add sheet">+</button>
</div>

<div id="statusbar">
  <span id="status-left">Ready</span>
  <span id="status-right"></span>
</div>

<!-- ── GitHub panel ── -->
<div id="gh-panel">
  <h2>
    <span>⎇ GitHub Versioning</span>
    <button onclick="toggleGHPanel()" title="Close">✕</button>
  </h2>
  <div class="gh-tabs">
    <div class="gh-tab active" id="tab-config"  onclick="ghTab('config')">Config</div>
    <div class="gh-tab"        id="tab-history" onclick="ghTab('history')">History</div>
    <div class="gh-tab"        id="tab-commit"  onclick="ghTab('commit')">Commit</div>
  </div>

  <!-- Config -->
  <div class="gh-body" id="gh-config">
    <div class="gh-field">
      <label>Personal Access Token (repo scope)</label>
      <input type="password" id="gh-token" placeholder="ghp_…">
    </div>
    <div class="gh-field">
      <label>Repository (owner/repo)</label>
      <input type="text" id="gh-repo" placeholder="username/my-data">
    </div>
    <div class="gh-field">
      <label>Branch</label>
      <input type="text" id="gh-branch" placeholder="main">
    </div>
    <div class="gh-field">
      <label>File path in repo</label>
      <input type="text" id="gh-path" placeholder="data/my_file.csv">
    </div>
    <button class="gh-btn" onclick="saveGHConfig()">Save Config</button>
    <div id="gh-status"></div>
  </div>

  <!-- History -->
  <div class="gh-body" id="gh-history" style="display:none">
    <button class="gh-btn" onclick="loadHistory()" id="gh-load-btn">Load History</button>
    <div id="gh-commits" style="margin-top:10px"></div>
  </div>

  <!-- Commit -->
  <div class="gh-body" id="gh-commit" style="display:none">
    <div class="gh-field">
      <label>Commit message</label>
      <input type="text" id="gh-commit-msg" placeholder="Update data">
    </div>
    <button class="gh-btn" onclick="doCommit()">Push Commit</button>
    <div id="gh-commit-status" style="margin-top:8px;font-size:11px"></div>
  </div>
</div>

<!-- ── Diff overlay ── -->
<div id="diff-overlay">
  <div id="diff-toolbar">
    <h3 id="diff-title">Visual Diff</h3>
    <button onclick="loadVersionIntoEditor()">Load this version</button>
    <button onclick="closeDiff()">✕ Close</button>
  </div>
  <div id="diff-legend">
    <span><span class="legend-dot" style="background:#b8ddc2"></span>Added</span>
    <span><span class="legend-dot" style="background:#f5b8b8"></span>Removed</span>
    <span><span class="legend-dot" style="background:#fff3a0"></span>Changed cell</span>
    <span><span class="legend-dot" style="background:#fff"></span>Unchanged</span>
  </div>
  <div id="diff-wrap"><table id="diff-table"><thead></thead><tbody></tbody></table></div>
</div>

<!-- Generic text-input modal -->
<div id="modal-overlay">
  <div id="modal-box">
    <h3 id="modal-title">Input</h3>
    <input type="text" id="modal-input">
    <div class="modal-btns">
      <button onclick="modalCancel()">Cancel</button>
      <button class="primary" onclick="modalOk()">OK</button>
    </div>
  </div>
</div>

<script>
// ── State ──────────────────────────────────────────────────────────────────
let S = { headers: [], rows: [], filepath: null, modified: false };
let _xlsxWb = null;   // original SheetJS workbook kept in memory to preserve cell styles
let selectedRows = new Set();   // indices of selected rows
let selectedCols = new Set();   // indices of selected columns
let anchorRow = -1;             // shift-click anchor
let anchorCol = -1;
let _pendingMouseEvent = null;  // mousedown captured before focus fires
let sortCol = -1;
let sortAsc = true;
let searchTerm = '';
let highlightCells = new Set(); // "r,c"

// ── Boot ───────────────────────────────────────────────────────────────────
(async function init() {
  const res = await fetch('/api/data');
  S = await res.json();
  setBaseline(S.headers, S.rows);
  render();
  document.addEventListener('keydown', onKeyDown);
})();

// ── Rendering ──────────────────────────────────────────────────────────────
function render() {
  renderHeaders();
  renderRows();
  updateStatus();
  autoResize();
  renderSheetTabs();
}

// ── Auto-resize columns to fit content ─────────────────────────────────────
const _measureCtx = document.createElement('canvas').getContext('2d'); // reused across calls
function autoResize() {
  const PAD = 28, MIN = 60, MAX = 500;
  const ths = document.querySelectorAll('#header-row th');

  S.headers.forEach((header, c) => {
    _measureCtx.font = 'bold 13px Arial';
    let maxPx = _measureCtx.measureText(header).width;

    _measureCtx.font = '13px Arial';
    for (const row of S.rows) {
      const w = _measureCtx.measureText(row[c] ?? '').width;
      if (w > maxPx) maxPx = w;
    }

    const width = Math.max(MIN, Math.min(MAX, Math.ceil(maxPx) + PAD));
    const th = ths[c + 1]; // +1 to skip the row-number corner
    if (th) {
      th.style.width    = width + 'px';
      th.style.minWidth = width + 'px';
    }
  });
}

function renderHeaders() {
  const tr = document.getElementById('header-row');
  tr.innerHTML = '';
  const corner = th('#', 'row-num');
  tr.appendChild(corner);

  S.headers.forEach((h, i) => {
    const el = document.createElement('th');
    el.title = 'Click to sort/select • Shift/Cmd+click to multi-select • Double-click to rename';
    el.style.minWidth = '120px';
    if (selectedCols.has(i)) el.classList.add('selected-col');
    const arrow = sortCol === i ? (sortAsc ? ' ▲' : ' ▼') : '';
    el.textContent = h + arrow;
    el.onclick = (e) => {
      selectCol(i, e);
      if (!e.shiftKey && !e.metaKey && !e.ctrlKey) sortBy(i);
    };
    el.ondblclick = (e) => { e.stopPropagation(); renameColumn(i); };
    tr.appendChild(el);
  });
}

function renderRows() {
  const tbody = document.getElementById('table-body');
  tbody.innerHTML = '';

  // Index incoming (new-in-commit) rows by the current-row index they appear before
  const incomingBefore = {};
  if (_inlineDiff) {
    _inlineDiff.incomingRows.forEach(g => {
      (incomingBefore[g.beforeCurIdx] = incomingBefore[g.beforeCurIdx] || []).push(g);
    });
  }

  function insertIncoming(beforeIdx) {
    (incomingBefore[beforeIdx] || []).forEach(g => {
      const tr = document.createElement('tr');
      tr.className = 'row-diff-added';
      const rn = document.createElement('td');
      rn.className = 'row-num';
      rn.textContent = '+';
      tr.appendChild(rn);
      S.headers.forEach(h => {
        const oi = g.commitH.indexOf(h);
        const td = document.createElement('td');
        td.textContent = oi >= 0 ? (g.commitRow[oi] ?? '') : '';
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
  }

  insertIncoming(0);

  S.rows.forEach((row, r) => {
    const entry = _inlineDiff ? (_inlineDiff.mapping[r] ?? { type: 'added' }) : null;
    const tr = document.createElement('tr');
    tr.className = r % 2 === 0 ? 'even' : 'odd';
    if (selectedRows.has(r)) tr.classList.add('selected-row');
    if (entry?.type === 'removed') tr.classList.add('row-diff-removed');

    const rn = document.createElement('td');
    rn.className = 'row-num';
    rn.textContent = r + 1;
    rn.title = 'Click • Shift+click to range • Cmd+click to toggle';
    rn.addEventListener('mousedown', (e) => { e.preventDefault(); selectRow(r, e); });
    tr.appendChild(rn);

    S.headers.forEach((h, c) => {
      const td = document.createElement('td');
      const key = `${r},${c}`;
      if (highlightCells.has(key)) td.classList.add('highlight');
      if (selectedCols.has(c)) td.classList.add('selected-col');

      const inp = document.createElement('input');
      inp.type = 'text';
      inp.value = row[c] != null ? row[c] : '';
      inp.dataset.r = r;
      inp.dataset.c = c;

      // Baseline (unsaved-edit) highlight
      const bv0 = baselineVal(r, h);
      if (bv0 === undefined || inp.value !== bv0) td.classList.add('cell-modified');

      // Inline diff: show commit value in yellow, strike through current value
      let oldValLabel = null;
      if (entry?.type === 'changed') {
        const oi = _inlineDiff.commitH.indexOf(h);
        const commitVal = oi >= 0 ? (entry.commitRow[oi] ?? '') : '';
        const curVal = row[c] ?? '';
        if (curVal !== commitVal) {
          td.classList.add('cell-diff-changed');
          inp.value = commitVal;
          oldValLabel = document.createElement('span');
          oldValLabel.className = 'old-val';
          oldValLabel.textContent = curVal;
        }
      }

      td.addEventListener('mousedown', (e) => { _pendingMouseEvent = e; });
      inp.addEventListener('focus', () => {
        selectCell(r, c, _pendingMouseEvent);
        _pendingMouseEvent = null;
      });
      inp.addEventListener('input', () => {
        S.rows[r][c] = inp.value;
        markModified();
        const bv = baselineVal(r, h);
        td.classList.toggle('cell-modified', bv === undefined || inp.value !== bv);
        if (entry?.type === 'changed') {
          const oi = _inlineDiff.commitH.indexOf(h);
          const commitVal = oi >= 0 ? (entry.commitRow[oi] ?? '') : '';
          const changed = inp.value !== commitVal;
          td.classList.toggle('cell-diff-changed', changed);
          if (oldValLabel) oldValLabel.style.display = changed ? '' : 'none';
        }
      });
      inp.addEventListener('keydown', onCellKeyDown);
      if (oldValLabel) td.appendChild(oldValLabel);
      td.appendChild(inp);
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
    insertIncoming(r + 1);
  });
}

function th(text, cls) {
  const el = document.createElement('th');
  el.textContent = text;
  if (cls) el.className = cls;
  return el;
}

// ── Selection ──────────────────────────────────────────────────────────────
function resetSelection() {
  selectedRows = new Set(); selectedCols = new Set(); anchorRow = -1; anchorCol = -1;
}

function _applyRange(set, anchor, target, max) {
  set.clear();
  const lo = Math.min(anchor, target), hi = Math.max(anchor, target);
  for (let i = lo; i <= hi; i++) set.add(i);
}

function selectRow(r, e) {
  if (e && e.shiftKey && anchorRow >= 0) {
    _applyRange(selectedRows, anchorRow, r, S.rows.length - 1);
  } else if (e && (e.metaKey || e.ctrlKey)) {
    if (selectedRows.has(r)) selectedRows.delete(r); else { selectedRows.add(r); anchorRow = r; }
  } else {
    selectedRows = new Set([r]);
    anchorRow = r;
  }
  updateSelectionUI();
}

function selectCol(c, e) {
  if (e && e.shiftKey && anchorCol >= 0) {
    _applyRange(selectedCols, anchorCol, c, S.headers.length - 1);
  } else if (e && (e.metaKey || e.ctrlKey)) {
    if (selectedCols.has(c)) selectedCols.delete(c); else { selectedCols.add(c); anchorCol = c; }
  } else {
    selectedCols = new Set([c]);
    anchorCol = c;
  }
  updateSelectionUI();
}

function selectCell(r, c, e) {
  if (e && e.shiftKey) {
    if (anchorRow >= 0) _applyRange(selectedRows, anchorRow, r, S.rows.length - 1);
    else { selectedRows = new Set([r]); anchorRow = r; }
    if (anchorCol >= 0) _applyRange(selectedCols, anchorCol, c, S.headers.length - 1);
    else { selectedCols = new Set([c]); anchorCol = c; }
  } else if (e && (e.metaKey || e.ctrlKey)) {
    if (selectedRows.has(r)) selectedRows.delete(r); else { selectedRows.add(r); anchorRow = r; }
    if (selectedCols.has(c)) selectedCols.delete(c); else { selectedCols.add(c); anchorCol = c; }
  } else {
    selectedRows = new Set([r]); anchorRow = r;
    selectedCols = new Set([c]); anchorCol = c;
  }
  updateSelectionUI();
}

function updateSelectionUI() {
  // Rows
  document.querySelectorAll('#table-body tr').forEach((tr, i) =>
    tr.classList.toggle('selected-row', selectedRows.has(i)));
  // Columns — clear then re-apply
  document.querySelectorAll('.selected-col').forEach(el => el.classList.remove('selected-col'));
  selectedCols.forEach(c => {
    const nth = c + 2; // +1 for row-num col, +1 for 1-based nth-child
    document.querySelectorAll(`#header-row th:nth-child(${nth}), #table-body td:nth-child(${nth})`)
      .forEach(el => el.classList.add('selected-col'));
  });
  updateStatus();
}

// ── Keyboard navigation ────────────────────────────────────────────────────
function onCellKeyDown(e) {
  const r = +e.target.dataset.r;
  const c = +e.target.dataset.c;
  const cols = S.headers.length;
  const rows = S.rows.length;
  let nr = r, nc = c;

  if (e.key === 'Tab') {
    e.preventDefault();
    nc = e.shiftKey ? c - 1 : c + 1;
    if (nc < 0)  { nc = cols - 1; nr = r - 1; }
    if (nc >= cols) { nc = 0; nr = r + 1; }
  } else if (e.key === 'ArrowDown'  && !e.altKey) { nr = r + 1; }
    else if (e.key === 'ArrowUp'    && !e.altKey) { nr = r - 1; }
    else if (e.key === 'Enter') {
      e.preventDefault();
      nr = e.shiftKey ? r - 1 : r + 1;
    } else { return; }

  nr = Math.max(0, Math.min(nr, rows - 1));
  nc = Math.max(0, Math.min(nc, cols - 1));
  focusCell(nr, nc);
}

function onKeyDown(e) {
  if ((e.metaKey || e.ctrlKey) && e.key === 's') {
    e.preventDefault();
    e.shiftKey ? saveAs() : saveFile();
  }
  if ((e.metaKey || e.ctrlKey) && e.key === 'o') { e.preventDefault(); openFile(); }
}

function focusCell(r, c) {
  // Set selection directly (keyboard nav = single-select, no mouse event)
  selectedRows = new Set([r]); anchorRow = r;
  selectedCols = new Set([c]); anchorCol = c;
  const inp = document.querySelector(`input[data-r="${r}"][data-c="${c}"]`);
  if (inp) { inp.focus(); inp.select(); }
  // focus will fire selectCell(r,c,null) which collapses to single-select — that's fine
}

// ── Data operations ────────────────────────────────────────────────────────
function addRow() {
  const insertAt = selectedRows.size > 0 ? Math.max(...selectedRows) + 1 : S.rows.length;
  S.rows.splice(insertAt, 0, Array(S.headers.length).fill(''));
  selectedRows = new Set([insertAt]); anchorRow = insertAt;
  markModified();
  renderRows();
  focusCell(insertAt, selectedCols.size > 0 ? Math.min(...selectedCols) : 0);
}

function deleteSelectedRow() {
  if (S.rows.length === 0 || selectedRows.size === 0) return;
  const toDelete = [...selectedRows].sort((a, b) => b - a); // reverse so splices don't shift
  const label = toDelete.length === 1 ? `row ${toDelete[0] + 1}` : `${toDelete.length} rows`;
  if (!confirm(`Delete ${label}?`)) return;
  toDelete.forEach(r => S.rows.splice(r, 1));
  const next = S.rows.length > 0 ? Math.min(Math.min(...toDelete), S.rows.length - 1) : -1;
  selectedRows = next >= 0 ? new Set([next]) : new Set(); anchorRow = next;
  markModified();
  renderRows();
  if (next >= 0) focusCell(next, selectedCols.size > 0 ? Math.min(...selectedCols) : 0);
  updateStatus();
}

function addColumn() {
  const insertAt = selectedCols.size > 0 ? Math.max(...selectedCols) + 1 : S.headers.length;
  prompt_('New column name', '', name => {
    if (!name) return;
    S.headers.splice(insertAt, 0, name);
    S.rows.forEach(row => row.splice(insertAt, 0, ''));
    selectedCols = new Set([insertAt]); anchorCol = insertAt;
    markModified();
    render();
    if (selectedRows.size > 0) focusCell(Math.min(...selectedRows), insertAt);
  });
}

function deleteColumn() {
  if (S.headers.length === 0 || selectedCols.size === 0) return;
  const toDelete = [...selectedCols].sort((a, b) => b - a);
  const label = toDelete.length === 1
    ? `column "${S.headers[toDelete[0]]}"`
    : `${toDelete.length} columns (${toDelete.map(c => S.headers[c]).join(', ')})`;
  if (!confirm(`Delete ${label}?`)) return;
  toDelete.forEach(c => { S.headers.splice(c, 1); S.rows.forEach(row => row.splice(c, 1)); });
  const next = S.headers.length > 0 ? Math.min(Math.min(...toDelete), S.headers.length - 1) : -1;
  selectedCols = next >= 0 ? new Set([next]) : new Set(); anchorCol = next;
  markModified();
  render();
}

function renameColumn(idx) {
  idx = idx != null ? idx : (selectedCols.size > 0 ? Math.min(...selectedCols) : 0);
  const old = S.headers[idx];
  prompt_(`Rename "${old}" to:`, old, name => {
    if (!name) return;
    S.headers[idx] = name;
    markModified();
    renderHeaders();
    autoResize();
  });
}

function sortBy(col) {
  if (sortCol === col) sortAsc = !sortAsc;
  else { sortCol = col; sortAsc = true; }
  S.rows.sort((a, b) => {
    const av = a[col] || '', bv = b[col] || '';
    const an = parseFloat(av), bn = parseFloat(bv);
    const cmp = !isNaN(an) && !isNaN(bn) ? an - bn : av.localeCompare(bv);
    return sortAsc ? cmp : -cmp;
  });
  markModified();
  render();
}

// ── Search ─────────────────────────────────────────────────────────────────
function applySearch(term) {
  searchTerm = term.toLowerCase();
  highlightCells.clear();
  if (searchTerm) {
    S.rows.forEach((row, r) => {
      row.forEach((val, c) => {
        if ((val || '').toLowerCase().includes(searchTerm)) highlightCells.add(`${r},${c}`);
      });
    });
  }
  renderRows();
  const count = highlightCells.size;
  document.getElementById('status-right').textContent =
    searchTerm ? `${count} match${count !== 1 ? 'es' : ''}` : '';
}

function clearSearch() {
  document.getElementById('search-input').value = '';
  applySearch('');
}

// ── File operations ────────────────────────────────────────────────────────
function openFile() {
  const inp = document.createElement('input');
  inp.type = 'file';
  inp.accept = '.csv,.tsv,.txt,.xlsx';
  inp.onchange = async e => {
    const file = e.target.files[0];
    if (!file) return;
    const isXlsx = file.name.toLowerCase().endsWith('.xlsx');
    let body;
    if (isXlsx) {
      const buf = await file.arrayBuffer();
      const wb  = new ExcelJS.Workbook();
      await wb.xlsx.load(buf);
      _xlsxWb = wb;  // keep for style-preserving save
      const fmt = v => {
        if (v === null || v === undefined) return '';
        if (v instanceof Date) return v.toISOString().slice(0, 10);
        if (typeof v === 'object' && v.richText) return v.richText.map(r => r.text).join('');
        if (typeof v === 'number') return Number.isInteger(v) ? String(v) : String(v);
        return String(v);
      };
      const sheets = wb.worksheets.map(ws => {
        const rows2d = [];
        ws.eachRow({ includeEmpty: true }, row => {
          rows2d.push(row.values.slice(1).map(fmt));
        });
        // Trim trailing blank rows
        while (rows2d.length && rows2d[rows2d.length - 1].every(c => c === '')) rows2d.pop();
        const headers = rows2d[0] || [];
        const rows    = rows2d.slice(1);
        return { name: ws.name, headers, rows };
      });
      body = JSON.stringify({ sheets, filename: file.name, format: 'xlsx' });
    } else {
      _xlsxWb = null;
      const content = await file.text();
      body = JSON.stringify({ content, filename: file.name, format: 'csv' });
    }
    const res  = await fetch('/api/load-content', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body
    });
    const json = await res.json();
    if (json.ok) {
      S = await fetch('/api/data').then(r => r.json());
      setBaseline(S.headers, S.rows);
      sortCol = -1; sortAsc = true;
      resetSelection();
      clearSearch();
      render();
    } else {
      alert('Failed to load file: ' + (json.error || 'unknown error'));
    }
  };
  inp.click();
}

async function _buildXlsxBytes() {
  // Flush active sheet edits to server first
  await syncNow();
  const all = await fetch('/api/all-sheets').then(r => r.json());

  if (_xlsxWb) {
    // Update values in-place in the original ExcelJS workbook — styles are preserved
    all.sheets.forEach(sh => {
      let ws = _xlsxWb.getWorksheet(sh.name);
      if (!ws) ws = _xlsxWb.addWorksheet(sh.name);
      const data = [sh.headers || [], ...(sh.rows || [])];
      data.forEach((row, ri) => {
        row.forEach((val, ci) => {
          const cell = ws.getCell(ri + 1, ci + 1);
          // Update value only — cell.style is untouched so fill/font are preserved
          const num = val !== '' && val !== null && !isNaN(val) && val.trim() !== '';
          cell.value = num ? Number(val) : (val === '' ? null : val);
        });
        // Clear any cells beyond the new column count (deleted columns)
        const oldColCount = ws.getRow(ri + 1).cellCount;
        for (let ci = row.length + 1; ci <= oldColCount; ci++)
          ws.getCell(ri + 1, ci).value = null;
      });
      // Clear any rows beyond the new row count (deleted rows)
      const oldRowCount = ws.rowCount;
      for (let ri = data.length + 1; ri <= oldRowCount; ri++)
        ws.getRow(ri).eachCell(cell => { cell.value = null; });
    });
    return await _xlsxWb.xlsx.writeBuffer();
  }

  // No original workbook — create fresh (no styles)
  const wb = new ExcelJS.Workbook();
  all.sheets.forEach(sh => {
    const ws = wb.addWorksheet(sh.name);
    ws.addRow(sh.headers || []);
    (sh.rows || []).forEach(r => ws.addRow(r));
  });
  return await wb.xlsx.writeBuffer();
}

async function _saveXlsx(filepath) {
  // Post xlsx bytes directly as binary (avoids JSON base64 encoding issues)
  let u8;
  try { u8 = await _buildXlsxBytes(); }
  catch (err) { alert('Could not build xlsx: ' + err.message); return false; }
  const url = '/api/save-xlsx?filepath=' + encodeURIComponent(filepath || '');
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/octet-stream' },
    body: new Uint8Array(u8)
  });
  const data = await res.json();
  if (data.ok) {
    S.filepath = data.filepath; S.filetype = 'xlsx'; S.modified = false; updateStatus();
    return true;
  }
  alert('Save failed: ' + data.error);
  return false;
}

async function saveFile() {
  const isXlsx = (S.filepath || '').toLowerCase().endsWith('.xlsx') || S.filetype === 'xlsx';
  if (isXlsx) { await _saveXlsx(S.filepath); return; }
  await syncNow();
  const res = await fetch('/api/save', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({ filepath: S.filepath })
  });
  const data = await res.json();
  if (data.ok) {
    S.filepath = data.filepath; S.modified = false; updateStatus();
  } else if (data.error === 'no_path') {
    saveAs();
  } else {
    alert('Save failed: ' + data.error);
  }
}

async function saveAs() {
  prompt_('Save to path (leave blank to download as ' + (S.filetype || 'csv').toUpperCase() + '):', S.filepath || '', async path => {
    if (!path) { downloadFile(); return; }
    const isXlsx = path.toLowerCase().endsWith('.xlsx');
    if (isXlsx) { await _saveXlsx(path); return; }
    await syncNow();
    const res = await fetch('/api/save', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ filepath: path })
    });
    const data = await res.json();
    if (data.ok) { S.filepath = data.filepath; S.filetype = data.filetype || S.filetype; S.modified = false; updateStatus(); }
    else alert('Save failed: ' + data.error);
  });
}

async function downloadFile(fmt) {
  await syncNow();
  fmt = fmt || S.filetype || 'csv';
  const base = (S.filepath ? S.filepath.split(/[\\/]/).pop() : 'data').replace(/\.(csv|xlsx)$/i, '');
  if (fmt === 'xlsx') {
    try {
      const buf  = await _buildXlsxBytes();
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a'); a.href = url; a.download = base + '.xlsx'; a.click();
      URL.revokeObjectURL(url);
    } catch (err) { alert('Export failed: ' + err.message); return; }
  } else {
    const res = await fetch('/api/export', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ format: 'csv' })
    });
    if (!res.ok) { alert('Export failed'); return; }
    const blob = await res.blob();
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href = url;
    a.download = base + '.csv';
    a.click();
    URL.revokeObjectURL(url);
  }
  S.modified = false;
  updateStatus();
}

// Legacy alias
async function downloadCSV() { return downloadFile('csv'); }

// ── Server sync ────────────────────────────────────────────────────────────
let syncTimer = null;
function markModified() {
  S.modified = true;
  updateStatus();
  clearTimeout(syncTimer);
  syncTimer = setTimeout(syncNow, 800);
}

async function syncNow() {
  clearTimeout(syncTimer);
  await fetch('/api/update', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({ headers: S.headers, rows: S.rows })
  });
}

// ── Status bar ─────────────────────────────────────────────────────────────
function updateStatus() {
  const name = S.filepath ? S.filepath.split('/').pop().split('\\').pop() : 'Untitled';
  const mod  = S.modified ? ' ●' : '';
  document.title = `CSV Editor — ${name}${mod}`;
  document.getElementById('filepath-display').textContent = S.filepath || 'Untitled';
  const selParts = [];
  if (selectedRows.size === 1) selParts.push(`Row ${[...selectedRows][0] + 1}`);
  else if (selectedRows.size > 1) selParts.push(`${selectedRows.size} rows`);
  if (selectedCols.size === 1) selParts.push(`Col "${S.headers[[...selectedCols][0]]}"`);
  else if (selectedCols.size > 1) selParts.push(`${selectedCols.size} cols`);
  document.getElementById('status-left').textContent =
    `${S.rows.length} rows × ${S.headers.length} cols` +
    (selParts.length ? `  •  ${selParts.join(', ')} selected` : '') +
    (S.modified ? '  •  Unsaved changes' : '');
}

// ── GitHub panel ───────────────────────────────────────────────────────────
let _ghOpen = false;
let _diffVersion = null;  // { sha, headers, rows } of the version shown in diff

async function toggleGHPanel() {
  _ghOpen = !_ghOpen;
  document.getElementById('gh-panel').classList.toggle('open', _ghOpen);
  if (_ghOpen) await loadGHConfig();
}

async function loadGHConfig() {
  const cfg = await fetch('/api/github/config').then(r => r.json());
  document.getElementById('gh-repo').value   = cfg.repo   || '';
  document.getElementById('gh-branch').value = cfg.branch || 'main';
  document.getElementById('gh-path').value   = cfg.path   || '';
  // token: don't pre-fill for security
  setGHStatus(cfg.configured ? '✓ Configured' : 'Not configured', cfg.configured ? 'ok' : '');
}

async function saveGHConfig() {
  const token  = document.getElementById('gh-token').value.trim();
  const repo   = document.getElementById('gh-repo').value.trim();
  const branch = document.getElementById('gh-branch').value.trim() || 'main';
  const path   = document.getElementById('gh-path').value.trim();
  if (!repo || !path) { setGHStatus('Repo and file path are required', 'err'); return; }
  const body = { repo, branch, path };
  if (token) body.token = token;
  await fetch('/api/github/config', {
    method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(body)
  });
  setGHStatus('✓ Config saved', 'ok');
  document.getElementById('gh-token').value = '';
}

function ghTab(tab) {
  ['config', 'history', 'commit'].forEach(t => {
    document.getElementById(`gh-${t}`).style.display    = t === tab ? '' : 'none';
    document.getElementById(`tab-${t}`).classList.toggle('active', t === tab);
  });
}

function setStatus(id, msg, cls) {
  const el = document.getElementById(id);
  el.textContent = msg; el.className = cls || '';
}
function setGHStatus(msg, cls) { setStatus('gh-status', msg, cls); }

// ── Baseline (last committed / loaded state for instant diff) ──────────────
let baseline = { headers: [], rows: [] };

function setBaseline(headers, rows) {
  baseline = {
    headers: [...headers],
    rows: rows.map(r => [...r]),
  };
  // Re-apply highlights after baseline changes (e.g. after a commit)
  applyBaselineDiff();
}

// Return the baseline value for a cell, or undefined if the row/column is new
function baselineVal(r, colName) {
  const bc = baseline.headers.indexOf(colName);
  if (bc === -1) return undefined;          // column didn't exist in baseline
  const br = baseline.rows[r];
  if (!br) return undefined;                // row didn't exist in baseline
  return br[bc] ?? '';
}

// Scan every visible input and set/clear the cell-modified class
function applyBaselineDiff() {
  document.querySelectorAll('#table-body input[data-r]').forEach(inp => {
    const r = +inp.dataset.r, c = +inp.dataset.c;
    const bv = baselineVal(r, S.headers[c]);
    inp.closest('td').classList.toggle('cell-modified',
      bv === undefined || inp.value !== bv);
  });
}

let _commits = [];  // kept so button handlers can look up data by index without onclick string injection

async function loadHistory() {
  const btn = document.getElementById('gh-load-btn');
  btn.disabled = true; btn.textContent = 'Loading…';
  const data = await fetch('/api/github/history').then(r => r.json());
  btn.disabled = false; btn.textContent = 'Refresh';
  const box = document.getElementById('gh-commits');
  if (data.error) { box.innerHTML = `<div style="color:#e07070">${data.error}</div>`; return; }
  if (!data.commits.length) { box.innerHTML = '<div style="color:#778899">No commits found.</div>'; return; }
  _commits = data.commits;
  box.innerHTML = _commits.map((c, idx) => `
    <div class="commit-item">
      <div class="commit-sha">${c.short}</div>
      <div class="commit-msg">${escHtml(c.message)}</div>
      <div class="commit-meta">${escHtml(c.author)} · ${fmtDate(c.date)}</div>
      <div class="commit-actions">
        <button data-idx="${idx}" class="gh-diff-btn">⊕ Diff vs. current</button>
        <button data-idx="${idx}" class="gh-load-version-btn">↓ Load</button>
      </div>
    </div>`).join('');
  // Use event delegation — avoids any JS-in-HTML-attribute quoting issues
  box.onclick = e => {
    const diffBtn = e.target.closest('.gh-diff-btn');
    const loadBtn = e.target.closest('.gh-load-version-btn');
    if (diffBtn) { const c = _commits[+diffBtn.dataset.idx]; showDiff(c.sha, c.message); }
    if (loadBtn) { const c = _commits[+loadBtn.dataset.idx]; loadVersion(c.sha); }
  };
}

async function doCommit() {
  const msg = document.getElementById('gh-commit-msg').value.trim();
  if (!msg) { alert('Enter a commit message.'); return; }
  await syncNow();
  setStatus('gh-commit-status', 'Pushing…', '');
  const data = await fetch('/api/github/commit', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({ message: msg })
  }).then(r => r.json());
  if (data.error) {
    setStatus('gh-commit-status', '✗ ' + data.error, 'err');
  } else {
    setStatus('gh-commit-status', `✓ Committed ${data.sha}`, 'ok');
    document.getElementById('gh-commit-msg').value = '';
    setBaseline(S.headers, S.rows);  // committed state is the new clean baseline
    S.modified = false; updateStatus();
  }
}

// ── Inline diff engine ──────────────────────────────────────────────────────

// _inlineDiff: null | { sha, label, commitH, commitR, mapping, incomingRows }
// mapping[curRowIdx]: { type: 'same'|'changed'|'removed', commitRow? }
//   same    – row unchanged; no highlight
//   changed – row modified; commitRow has the incoming values
//   removed – row exists in current but not in commit; shown as ghost
// incomingRows: [{ beforeCurIdx, commitH, commitRow }]
//   rows that exist in commit but not in current; shown in green between current rows
let _inlineDiff = null;

function buildInlineDiffMapping(curH, curR, commitH, commitR) {
  // Diff current→commit: + means new in commit, - means removed from commit (ghost)
  const ops = computeDiff(curR, commitR);
  const mapping = [], incomingRows = [];
  let curIdx = 0;
  ops.forEach(op => {
    if      (op.t === '=') { mapping.push({ type: 'same',    commitRow: op.n }); curIdx++; }
    else if (op.t === '-') { mapping.push({ type: 'removed'                  }); curIdx++; }
    else if (op.t === '+') { incomingRows.push({ beforeCurIdx: curIdx, commitH, commitRow: op.n }); }
    else if (op.t === '~') { mapping.push({ type: 'changed',  commitRow: op.n }); curIdx++; }
  });
  return { mapping, incomingRows };
}

async function showDiff(sha, label) {
  const data = await fetch(`/api/github/version?sha=${sha}`).then(r => r.json());
  if (data.error) { alert('Could not fetch version: ' + data.error); return; }

  const { headers: commitH, rows: commitR } = data.format === 'xlsx'
    ? { headers: data.headers, rows: data.rows }
    : parseCSV(data.content);
  _diffVersion = { sha, headers: commitH, rows: commitR };

  const { mapping, incomingRows } = buildInlineDiffMapping(S.headers, S.rows, commitH, commitR);
  _inlineDiff = { sha, label, commitH, commitR, mapping, incomingRows };

  const added   = incomingRows.length;
  const removed = mapping.filter(m => m.type === 'removed').length;
  const changed = mapping.filter(m => m.type === 'changed').length;
  document.getElementById('diff-banner-label').textContent = sha.slice(0, 7) + (label ? ` "${label}"` : '');
  document.getElementById('diff-banner-stats').textContent =
    [added && `${added} added`, removed && `${removed} removed`, changed && `${changed} modified`]
      .filter(Boolean).join(' · ') || 'no changes';
  document.getElementById('diff-banner').classList.add('open');

  render();
}

function clearInlineDiff() {
  _inlineDiff = null;
  document.getElementById('diff-banner').classList.remove('open');
  render();
}

function closeDiff() {
  document.getElementById('diff-overlay').classList.remove('open');
}

function loadVersionIntoEditor() {
  if (!_diffVersion) return;
  S.headers = _diffVersion.headers;
  S.rows    = _diffVersion.rows;
  setBaseline(S.headers, S.rows);
  selectedRows = new Set(); selectedCols = new Set(); anchorRow = -1; anchorCol = -1;
  markModified(); render(); closeDiff();
}

async function loadVersion(sha) {
  if (!confirm('Replace current data with this version? Unsaved changes will be lost.')) return;
  const data = await fetch(`/api/github/version?sha=${sha}`).then(r => r.json());
  if (data.error) { alert('Error: ' + data.error); return; }
  const { headers, rows } = data.format === 'xlsx'
    ? { headers: data.headers, rows: data.rows }
    : parseCSV(data.content);
  S.headers = headers; S.rows = rows;
  setBaseline(headers, rows);
  selectedRows = new Set(); selectedCols = new Set(); anchorRow = -1; anchorCol = -1;
  markModified(); render();
}

function parseCSV(text) {
  const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  const parse = line => {
    const cells = []; let cur = '', inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') { if (inQ && line[i+1] === '"') { cur += '"'; i++; } else inQ = !inQ; }
      else if (ch === ',' && !inQ) { cells.push(cur); cur = ''; }
      else cur += ch;
    }
    cells.push(cur); return cells;
  };
  const all = lines.filter(l => l.trim()).map(parse);
  if (!all.length) return { headers: [], rows: [] };
  const headers = all[0];
  const rows = all.slice(1).map(r => {
    while (r.length < headers.length) r.push('');
    return r;
  });
  return { headers, rows };
}

// LCS-based row diff
function computeDiff(oldR, newR) {
  const n = oldR.length, m = newR.length;
  // Pre-compute keys once — avoids O(n×m) repeated JSON.stringify calls in the DP loop
  const oldK = oldR.map(r => JSON.stringify(r));
  const newK = newR.map(r => JSON.stringify(r));
  const dp = Array.from({length: n+1}, () => new Int32Array(m+1));
  for (let i = 1; i <= n; i++)
    for (let j = 1; j <= m; j++)
      dp[i][j] = oldK[i-1] === newK[j-1]
        ? dp[i-1][j-1] + 1
        : Math.max(dp[i-1][j], dp[i][j-1]);

  const ops = []; let i = n, j = m;
  while (i > 0 || j > 0) {
    if (i > 0 && j > 0 && oldK[i-1] === newK[j-1])
      { ops.unshift({t:'=', o:oldR[i-1], n:newR[j-1]}); i--; j--; }
    else if (j > 0 && (i === 0 || dp[i][j-1] >= dp[i-1][j]))
      { ops.unshift({t:'+', n:newR[j-1]}); j--; }
    else
      { ops.unshift({t:'-', o:oldR[i-1]}); i--; }
  }

  // Collect consecutive runs of + and - (in any order) and pair by similarity.
  // Unpaired + = truly new incoming row. Unpaired - = truly deleted ghost row.
  function cellSim(r1, r2) {
    let m = 0;
    for (let i = 0; i < Math.min(r1.length, r2.length); i++)
      if ((r1[i] ?? '') === (r2[i] ?? '')) m++;
    return m;
  }
  const result = []; let k = 0;
  while (k < ops.length) {
    if (ops[k].t === '+' || ops[k].t === '-') {
      const adds = [], dels = [];
      while (k < ops.length && (ops[k].t === '+' || ops[k].t === '-')) {
        if (ops[k].t === '+') adds.push(ops[k++]);
        else dels.push(ops[k++]);
      }
      const usedAdd = new Set();
      const pairedAdd = new Array(dels.length).fill(-1); // pairedAdd[d] = add index
      dels.forEach((del, d) => {
        let bestA = -1, bestSim = 0;          // bestSim=0 means "must beat 0 to pair"
        adds.forEach((add, a) => {
          if (usedAdd.has(a)) return;
          const s = cellSim(del.o, add.n);
          if (s > bestSim) { bestSim = s; bestA = a; }
        });
        if (bestA >= 0) { pairedAdd[d] = bestA; usedAdd.add(bestA); }
      });

      // Build reverse map: add index → del index (for paired adds)
      const delForAdd = new Map();
      pairedAdd.forEach((a, d) => { if (a >= 0) delForAdd.set(a, d); });
      // Emit in commit order (adds order) so incoming rows land at the right position
      let nextDel = 0;
      adds.forEach((add, a) => {
        if (delForAdd.has(a)) {
          const d = delForAdd.get(a);
          while (nextDel < d) result.push(dels[nextDel++]); // unpaired dels before this pair
          result.push({ t:'~', o:dels[d].o, n:add.n });
          nextDel = d + 1;
        } else {
          result.push(add); // truly new row — emitted in commit order
        }
      });
      while (nextDel < dels.length) result.push(dels[nextDel++]); // remaining unpaired dels
    } else {
      result.push(ops[k++]);
    }
  }
  return result;
}

function renderDiff(oldH, oldR, newH, newR) {
  const allH = [...new Set([...oldH, ...newH])];
  const diff = computeDiff(oldR, newR);

  const thead = document.querySelector('#diff-table thead');
  const tbody = document.querySelector('#diff-table tbody');
  thead.innerHTML = `<tr><th>±</th><th>#</th>${allH.map(h => `<th>${escHtml(h)}</th>`).join('')}</tr>`;
  tbody.innerHTML = '';

  let oldIdx = 0, newIdx = 0;
  diff.forEach(op => {
    if (op.t === '=') {
      const tr = makeRow('', ++oldIdx + '/' + ++newIdx, allH, oldH, op.o, null, null);
      tbody.appendChild(tr); return;
    }
    if (op.t === '+') {
      const tr = makeRow('+', '—/' + ++newIdx, allH, newH, op.n, null, null);
      tr.className = 'diff-added'; tbody.appendChild(tr); return;
    }
    if (op.t === '-') {
      const tr = makeRow('-', ++oldIdx + '/—', allH, oldH, op.o, null, null);
      tr.className = 'diff-removed'; tbody.appendChild(tr); return;
    }
    if (op.t === '~') {
      // old row — compare against new so changed cells are highlighted
      const trOld = makeRow('-', ++oldIdx + '/…', allH, oldH, op.o, newH, op.n);
      trOld.className = 'diff-modified-old'; tbody.appendChild(trOld);
      // new row — compare against old
      const trNew = makeRow('+', '…/' + ++newIdx, allH, newH, op.n, oldH, op.o);
      trNew.className = 'diff-modified-new'; tbody.appendChild(trNew);
    }
  });

  // Summary row
  const added   = diff.filter(o => o.t === '+').length;
  const removed = diff.filter(o => o.t === '-').length;
  const changed = diff.filter(o => o.t === '~').length;
  const summary = document.createElement('tr');
  summary.innerHTML = `<td colspan="${allH.length + 2}" style="padding:8px 12px;color:#555;font-size:11px;background:#f8f8f8">
    ${added} added · ${removed} removed · ${changed} modified · ${diff.filter(o=>o.t==='=').length} unchanged</td>`;
  tbody.appendChild(summary);
}

// compareH is the header set that indexes compareData (may differ from rowH)
function makeRow(marker, idx, allH, rowH, rowData, compareH, compareData) {
  const tr = document.createElement('tr');
  const mTd = document.createElement('td');
  mTd.className = 'diff-marker'; mTd.textContent = marker;
  tr.appendChild(mTd);
  const iTd = document.createElement('td');
  iTd.style.cssText = 'font-size:10px;color:#888;white-space:nowrap;font-family:monospace';
  iTd.textContent = idx; tr.appendChild(iTd);

  allH.forEach(h => {
    const colIdx = rowH.indexOf(h);
    const val = colIdx >= 0 ? (rowData[colIdx] ?? '') : '';
    const td  = document.createElement('td');
    td.textContent = val;
    if (compareData !== null) {
      // Use compareH (the other side's headers) to index into compareData correctly
      const cmpIdx = compareH.indexOf(h);
      const cmpVal = cmpIdx >= 0 ? (compareData[cmpIdx] ?? '') : '';
      if (val !== cmpVal) td.classList.add('cell-changed');
    }
    tr.appendChild(td);
  });
  return tr;
}

// ── Utilities ───────────────────────────────────────────────────────────────
function escHtml(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function fmtDate(iso) {
  try { return new Date(iso).toLocaleString(undefined, {dateStyle:'medium', timeStyle:'short'}); }
  catch { return iso; }
}

// ── Modal prompt ───────────────────────────────────────────────────────────
let _modalResolve = null;
function prompt_(title, defaultVal, cb) {
  _modalResolve = cb;
  document.getElementById('modal-title').textContent = title;
  const inp = document.getElementById('modal-input');
  inp.value = defaultVal || '';
  document.getElementById('modal-overlay').classList.add('open');
  setTimeout(() => { inp.focus(); inp.select(); }, 50);
  inp.onkeydown = e => { if (e.key === 'Enter') modalOk(); else if (e.key === 'Escape') modalCancel(); };
}
function modalOk() {
  const val = document.getElementById('modal-input').value;
  document.getElementById('modal-overlay').classList.remove('open');
  if (_modalResolve) { _modalResolve(val); _modalResolve = null; }
}
function modalCancel() {
  document.getElementById('modal-overlay').classList.remove('open');
  _modalResolve = null;
}

// ── Sheet tabs ──────────────────────────────────────────────────────────────
function renderSheetTabs() {
  const bar = document.getElementById('sheet-tabs');
  // Remove existing tabs (keep the + button)
  [...bar.querySelectorAll('.sheet-tab')].forEach(t => t.remove());
  const addBtn = bar.querySelector('.sheet-tab-add');
  (S.sheets || []).forEach((sh, i) => {
    const tab = document.createElement('div');
    tab.className = 'sheet-tab' + (i === S.activeSheet ? ' active' : '');
    tab.textContent = sh.name;
    tab.title = 'Click to switch \u2022 Double-click to rename';
    tab.onclick = () => switchSheet(i);
    tab.ondblclick = (e) => { e.stopPropagation(); renameSheet(i, sh.name); };
    // Right-click context for delete
    tab.oncontextmenu = (e) => {
      e.preventDefault();
      if ((S.sheets || []).length > 1 && confirm(`Delete sheet "${sh.name}"?`)) deleteSheet(i);
    };
    bar.insertBefore(tab, addBtn);
  });
}

async function switchSheet(idx) {
  if (idx === S.activeSheet) return;
  await syncNow();
  await fetch('/api/switch-sheet', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({ index: idx })
  });
  S = await fetch('/api/data').then(r => r.json());
  setBaseline(S.headers, S.rows);
  resetSelection(); clearSearch();
  render();
}

async function addSheet() {
  await syncNow();
  const res  = await fetch('/api/add-sheet', { method: 'POST', headers: {'Content-Type': 'application/json'}, body: '{}' });
  const data = await res.json();
  S = await fetch('/api/data').then(r => r.json());
  setBaseline(S.headers, S.rows);
  resetSelection(); clearSearch();
  render();
}

async function renameSheet(idx, current) {
  prompt_('Rename sheet:', current, async name => {
    if (!name || name === current) return;
    await fetch('/api/rename-sheet', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ index: idx, name })
    });
    S.sheets[idx].name = name;
    renderSheetTabs();
  });
}

async function deleteSheet(idx) {
  await syncNow();
  await fetch('/api/delete-sheet', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({ index: idx })
  });
  S = await fetch('/api/data').then(r => r.json());
  setBaseline(S.headers, S.rows);
  resetSelection(); clearSearch();
  render();
}
</script>
</body>
</html>
"""


# ── HTTP handler ──────────────────────────────────────────────────────────────

class Handler(BaseHTTPRequestHandler):

    def do_GET(self):
        if self.path == '/':
            self._serve_html()
        elif self.path == '/api/data':
            self._json({
                'headers':     state.headers,
                'rows':        state.rows,
                'filepath':    state.filepath,
                'filetype':    state.filetype,
                'sheets':      [{'name': s['name']} for s in state.sheets],
                'activeSheet': state.active_sheet,
            })
        elif self.path == '/api/all-sheets':
            self._json({'sheets': state.sheets, 'activeSheet': state.active_sheet})
        elif self.path == '/api/github/config':
            self._json({
                'token':       ('*' * 4 + gh.token[-4:]) if len(gh.token) > 4 else ('*' * len(gh.token)),
                'repo':        gh.repo,
                'branch':      gh.branch,
                'path':        gh.path,
                'configured':  bool(gh.token and gh.repo and gh.path),
            })
        elif self.path == '/api/github/history':
            if not gh.token or not gh.repo:
                self._json({'error': 'Not configured'}); return
            qs = f'?path={urllib.parse.quote(gh.path, safe="/")}&sha={urllib.parse.quote(gh.branch, safe="/")}&per_page=40'
            result, err = _gh_api('GET', f'/repos/{gh.repo}/commits{qs}')
            if err:
                self._json({'error': err}); return
            self._json({'commits': [{
                'sha':     c['sha'],
                'short':   c['sha'][:7],
                'message': c['commit']['message'].split('\n')[0][:80],
                'author':  c['commit']['author']['name'],
                'date':    c['commit']['author']['date'],
            } for c in result]})
        elif self.path.startswith('/api/github/version'):
            qs   = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            sha  = qs.get('sha', [''])[0]
            path = qs.get('path', [gh.path])[0]
            if not sha or not gh.token or not gh.repo:
                self._json({'error': 'Missing params'}); return
            result, err = _gh_api('GET', f'/repos/{gh.repo}/contents/{path.lstrip("/")}?ref={sha}')
            if err:
                self._json({'error': err}); return
            raw_bytes = base64.b64decode(result['content'].replace('\n', ''))
            if path.lower().endswith('.xlsx'):
                if not _XLSX_OK:
                    self._json({'error': 'openpyxl not installed — run: pip install openpyxl'}); return
                try:
                    headers, rows = _parse_xlsx(raw_bytes)
                    self._json({'headers': headers, 'rows': rows, 'sha': sha, 'format': 'xlsx'})
                except Exception as e:
                    self._json({'error': str(e)})
            else:
                content = raw_bytes.decode('utf-8-sig')
                self._json({'content': content, 'sha': sha, 'format': 'csv'})
        else:
            self._send(404, b'Not found')

    def do_POST(self):
        length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(length) if length else b'{}'
        # Binary endpoints must be handled before JSON parsing
        if self.path.startswith('/api/save-xlsx'):
            qs       = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            filepath = qs.get('filepath', [''])[0] or state.filepath or ''
            print(f'[save-xlsx] filepath={filepath!r} body_len={len(body)}')
            if not filepath:
                print('[save-xlsx] ERROR: no_path')
                self._json({'ok': False, 'error': 'no_path'}); return
            if '/' not in filepath and '\\' not in filepath:
                filepath = str(Path.cwd() / filepath)
                print(f'[save-xlsx] resolved to {filepath!r}')
            try:
                if not body:
                    print('[save-xlsx] ERROR: no data received')
                    self._json({'ok': False, 'error': 'no data received'}); return
                p = Path(filepath)
                p.parent.mkdir(parents=True, exist_ok=True)
                p.write_bytes(body)
                print(f'[save-xlsx] wrote {len(body)} bytes to {filepath!r}')
                state.filepath = filepath
                state.filetype = 'xlsx'
                self._json({'ok': True, 'filepath': filepath})
            except Exception as e:
                print(f'[save-xlsx] ERROR: {e}')
                self._json({'ok': False, 'error': str(e)})
            return
        try:
            data = json.loads(body)
        except json.JSONDecodeError:
            data = {}

        if self.path == '/api/load-content':
            filename = data.get('filename', 'data.csv')
            fmt      = data.get('format', 'csv')
            if fmt == 'xlsx':
                # xlsx is parsed client-side by SheetJS; server receives {sheets: [{name, headers, rows}]}
                raw_sheets = data.get('sheets', [])
                if not raw_sheets:
                    raw_sheets = [{'name': 'Sheet1', 'headers': [], 'rows': []}]
                state.sheets       = raw_sheets
                state.active_sheet = 0
                state.filetype     = 'xlsx'
            else:
                content  = data.get('content', '')
                reader   = csv.reader(io.StringIO(content))
                all_rows = list(reader)
                if all_rows:
                    state_headers = all_rows[0]
                    state_rows    = [list(r) for r in all_rows[1:]]
                    for row in state_rows:
                        while len(row) < len(state_headers):
                            row.append('')
                else:
                    state_headers = []
                    state_rows    = []
                state.sheets       = [{'name': 'Sheet1', 'headers': state_headers, 'rows': state_rows}]
                state.active_sheet = 0
                state.filetype     = 'csv'
            state.filepath = filename
            self._json({'ok': True})

        elif self.path == '/api/update':
            state.headers = data.get('headers', state.headers)
            state.rows    = data.get('rows',    state.rows)
            self._json({'ok': True})

        elif self.path == '/api/save':
            filepath = data.get('filepath') or state.filepath
            if not filepath:
                self._json({'ok': False, 'error': 'no_path'})
                return
            # If only a bare filename (no directory), resolve against cwd
            p_check = Path(filepath)
            if not p_check.is_absolute() and '/' not in filepath and '\\' not in filepath:
                filepath = str(Path.cwd() / filepath)
            try:
                p = Path(filepath)
                p.parent.mkdir(parents=True, exist_ok=True)
                if p.suffix.lower() == '.xlsx':
                    # xlsx bytes are generated client-side by SheetJS and written
                    # via /api/save-bytes — this branch handles CSV-only saves
                    # (fallback: if raw_bytes provided here, write them directly)
                    raw_b64 = data.get('raw_bytes')
                    if raw_b64:
                        p.write_bytes(base64.b64decode(raw_b64))
                    else:
                        self._json({'ok': False, 'error': 'xlsx_needs_bytes'}); return
                    state.filetype = 'xlsx'
                else:
                    with open(filepath, 'w', newline='', encoding='utf-8') as f:
                        csv.writer(f).writerow(state.headers)
                        csv.writer(f).writerows(state.rows)
                    state.filetype = 'csv'
                state.filepath = filepath
                self._json({'ok': True, 'filepath': filepath, 'filetype': state.filetype})
            except Exception as e:
                self._json({'ok': False, 'error': str(e)})

        elif self.path == '/api/export':
            fmt = data.get('format', state.filetype)
            if fmt == 'xlsx':
                if not _XLSX_OK:
                    self._json({'error': 'openpyxl not installed — run: pip install openpyxl'}); return
                try:
                    raw = _write_xlsx(state.headers, state.rows)
                except Exception as e:
                    self._json({'error': str(e)}); return
                self.send_response(200)
                self.send_header('Content-Type',
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Length', str(len(raw)))
                self.end_headers()
                self.wfile.write(raw)
            else:
                buf = io.StringIO()
                w   = csv.writer(buf)
                w.writerow(state.headers)
                w.writerows(state.rows)
                raw = buf.getvalue().encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'text/csv; charset=utf-8')
                self.send_header('Content-Length', str(len(raw)))
                self.end_headers()
                self.wfile.write(raw)

        elif self.path == '/api/switch-sheet':
            idx = int(data.get('index', 0))
            if 0 <= idx < len(state.sheets):
                state.active_sheet = idx
            self._json({'ok': True})

        elif self.path == '/api/add-sheet':
            name = data.get('name', '').strip()
            if not name:
                existing = {s['name'] for s in state.sheets}
                i = len(state.sheets) + 1
                while f'Sheet{i}' in existing:
                    i += 1
                name = f'Sheet{i}'
            state.sheets.append({'name': name, 'headers': list(state.headers) if state.headers else [], 'rows': []})
            state.active_sheet = len(state.sheets) - 1
            self._json({'ok': True, 'index': state.active_sheet, 'name': name})

        elif self.path == '/api/rename-sheet':
            idx  = int(data.get('index', state.active_sheet))
            name = data.get('name', '').strip()
            if name and 0 <= idx < len(state.sheets):
                state.sheets[idx]['name'] = name
            self._json({'ok': True})

        elif self.path == '/api/delete-sheet':
            idx = int(data.get('index', state.active_sheet))
            if len(state.sheets) > 1 and 0 <= idx < len(state.sheets):
                state.sheets.pop(idx)
                state.active_sheet = min(state.active_sheet, len(state.sheets) - 1)
            self._json({'ok': True})

        elif self.path == '/api/github/config':
            gh.token  = data.get('token',  gh.token)
            gh.repo   = data.get('repo',   gh.repo).strip()
            gh.branch = data.get('branch', gh.branch).strip() or 'main'
            gh.path   = data.get('path',   gh.path).strip().lstrip('/')
            save_gh_config()
            self._json({'ok': True})

        elif self.path == '/api/github/commit':
            if not gh.token or not gh.repo or not gh.path:
                self._json({'error': 'GitHub not configured'}); return
            message    = data.get('message', 'Update file').strip() or 'Update file'
            clean_path = gh.path.lstrip('/')
            existing, _ = _gh_api('GET',
                f'/repos/{gh.repo}/contents/{clean_path}?ref={urllib.parse.quote(gh.branch)}')
            file_sha = existing.get('sha') if isinstance(existing, dict) else None
            # Encode content based on current filetype
            if state.filetype == 'xlsx':
                if not _XLSX_OK:
                    self._json({'error': 'openpyxl not installed — run: pip install openpyxl'}); return
                try:
                    encoded = base64.b64encode(_write_xlsx(state.headers, state.rows)).decode()
                except Exception as e:
                    self._json({'error': str(e)}); return
            else:
                buf = io.StringIO()
                csv.writer(buf).writerow(state.headers)
                csv.writer(buf).writerows(state.rows)
                encoded = base64.b64encode(buf.getvalue().encode('utf-8')).decode()
            payload = {'message': message, 'content': encoded, 'branch': gh.branch}
            if file_sha:
                payload['sha'] = file_sha
            result, err = _gh_api('PUT', f'/repos/{gh.repo}/contents/{clean_path}', payload)
            if err:
                self._json({'error': err})
            else:
                self._json({'ok': True, 'sha': result['commit']['sha'][:7],
                            'url': result['commit']['html_url']})
        else:
            self._send(404, b'Not found')

    # ── helpers ──

    def _serve_html(self):
        raw = HTML.encode('utf-8')
        self.send_response(200)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.send_header('Content-Length', str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def _json(self, obj):
        raw = json.dumps(obj).encode('utf-8')
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def _send(self, code, body):
        self.send_response(code)
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *_):
        pass   # silence server logs


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    load_gh_config()

    if not _XLSX_OK:
        print('Note: openpyxl not installed — Excel (.xlsx) support disabled.')
        print('      pip install openpyxl\n')

    # Optionally load a file passed as a CLI argument
    if len(sys.argv) > 1:
        path = Path(sys.argv[1]).expanduser().resolve()
        if path.exists():
            if path.suffix.lower() == '.xlsx':
                if not _XLSX_OK:
                    print('Error: openpyxl is required to open .xlsx files.')
                    print('       pip install openpyxl')
                    sys.exit(1)
                headers, rows = _parse_xlsx(path.read_bytes())
                state.sheets       = [{'name': 'Sheet1', 'headers': headers, 'rows': rows}]
                state.active_sheet = 0
                state.filetype     = 'xlsx'
            else:
                with open(path, newline='', encoding='utf-8-sig') as f:
                    rows = list(csv.reader(f))
                if rows:
                    headers = rows[0]
                    parsed_rows = [list(r) for r in rows[1:]]
                    for row in parsed_rows:
                        while len(row) < len(headers):
                            row.append('')
                else:
                    headers = []
                    parsed_rows = []
                state.sheets       = [{'name': 'Sheet1', 'headers': headers, 'rows': parsed_rows}]
                state.active_sheet = 0
                state.filetype     = 'csv'
            state.filepath = str(path)
        else:
            print(f'File not found: {path}')

    server = HTTPServer(('localhost', 0), Handler)
    port   = server.server_address[1]
    url    = f'http://localhost:{port}'

    print(f'CSV Editor  →  {url}')
    print('Press Ctrl+C to quit.\n')

    threading.Timer(0.6, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('Shutting down.')


if __name__ == '__main__':
    main()
