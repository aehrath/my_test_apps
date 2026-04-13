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


# ── HTML/CSS/JS (loaded from index.html next to this script) ─────────────────

_HTML_FILE = Path(__file__).parent / 'index.html'

def _load_html():
    return _HTML_FILE.read_text(encoding='utf-8')


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
            # Detect xlsx by extension OR by zip magic bytes (PK\x03\x04) — gh.path may lack extension
            is_xlsx = path.lower().endswith('.xlsx') or raw_bytes[:4] == b'PK\x03\x04'
            if is_xlsx:
                # Return raw bytes as base64 — browser parses with SheetJS (no openpyxl needed)
                self._json({'rawB64': base64.b64encode(raw_bytes).decode('ascii'), 'sha': sha, 'format': 'xlsx'})
            else:
                try:
                    content = raw_bytes.decode('utf-8-sig')
                except UnicodeDecodeError:
                    content = raw_bytes.decode('latin-1')
                self._json({'content': content, 'sha': sha, 'format': 'csv'})
        else:
            self._send(404, b'Not found')

    def do_POST(self):
        length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(length) if length else b'{}'
        # Binary endpoints must be handled before JSON parsing
        if self.path.startswith('/api/github/commit-xlsx'):
            qs      = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            message = qs.get('message', ['Update file'])[0] or 'Update file'
            if not gh.token or not gh.repo or not gh.path:
                self._json({'error': 'GitHub not configured'}); return
            clean_path = gh.path.lstrip('/')
            existing, _ = _gh_api('GET',
                f'/repos/{gh.repo}/contents/{clean_path}?ref={urllib.parse.quote(gh.branch)}')
            file_sha = existing.get('sha') if isinstance(existing, dict) else None
            encoded  = base64.b64encode(body).decode('ascii')
            payload  = {'message': message, 'content': encoded, 'branch': gh.branch}
            if file_sha:
                payload['sha'] = file_sha
            result, err = _gh_api('PUT', f'/repos/{gh.repo}/contents/{clean_path}', payload)
            if err:
                self._json({'error': err})
            else:
                self._json({'ok': True, 'sha': result['commit']['sha'][:7],
                            'url': result['commit']['html_url']})
            return
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
            # Encode content (CSV only — xlsx is handled by /api/github/commit-xlsx)
            if state.filetype == 'xlsx':
                self._json({'error': 'use /api/github/commit-xlsx for xlsx files'}); return
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
        raw = _load_html().encode('utf-8')
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
