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
import colorsys as _colorsys
import xml.etree.ElementTree as _ET
import zipfile as _zipfile
import posixpath as _ppath

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


_THEME_ORDER = ['lt1','dk1','lt2','dk2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink']
_INDEXED_COLORS = {
    0: '000000', 1: 'FFFFFF', 2: 'FF0000', 3: '00FF00', 4: '0000FF', 5: 'FFFF00', 6: 'FF00FF', 7: '00FFFF',
    8: '000000', 9: 'FFFFFF', 10: 'FF0000', 11: '00FF00', 12: '0000FF', 13: 'FFFF00', 14: 'FF00FF', 15: '00FFFF',
    16: '800000', 17: '008000', 18: '000080', 19: '808000', 20: '800080', 21: '008080', 22: 'C0C0C0', 23: '808080',
    24: '9999FF', 25: '993366', 26: 'FFFFCC', 27: 'CCFFFF', 28: '660066', 29: 'FF8080', 30: '0066CC', 31: 'CCCCFF',
    32: '000080', 33: 'FF00FF', 34: 'FFFF00', 35: '00FFFF', 36: '800080', 37: '800000', 38: '008080', 39: '0000FF',
    40: '00CCFF', 41: 'CCFFFF', 42: 'CCFFCC', 43: 'FFFF99', 44: '99CCFF', 45: 'FF99CC', 46: 'CC99FF', 47: 'FFCC99',
    48: '3366FF', 49: '33CCCC', 50: '99CC00', 51: 'FFCC00', 52: 'FF9900', 53: 'FF6600', 54: '666699', 55: '969696',
    56: '003366', 57: '339966', 58: '003300', 59: '333300', 60: '993300', 61: '993366', 62: '333399', 63: '333333',
}


def _parse_theme_colors(wb):
    xml = getattr(wb, 'loaded_theme', None)
    if not xml:
        return None
    try:
        root = _ET.fromstring(xml)
        scheme = root.find('.//{*}clrScheme')
        if scheme is None:
            return None
        colors = []
        for tag in _THEME_ORDER:
            el = scheme.find(f'{{*}}{tag}')
            if el is None:
                colors.append('000000')
                continue
            srgb = el.find('{*}srgbClr')
            sysc = el.find('{*}sysClr')
            colors.append(
                (srgb.get('val') if srgb is not None else None)
                or (sysc.get('lastClr') if sysc is not None else None)
                or '000000'
            )
        return colors
    except Exception:
        return None


def _apply_tint(hex_color, tint):
    if not tint:
        return hex_color.upper()
    r = int(hex_color[0:2], 16) / 255.0
    g = int(hex_color[2:4], 16) / 255.0
    b = int(hex_color[4:6], 16) / 255.0
    h, l, s = _colorsys.rgb_to_hls(r, g, b)
    l = l * (1 - tint) + tint if tint >= 0 else l * (1 + tint)
    l = max(0.0, min(1.0, l))
    nr, ng, nb = _colorsys.hls_to_rgb(h, l, s)
    return ''.join(f'{round(v * 255):02X}' for v in (nr, ng, nb))


def _resolve_openpyxl_color(color, theme_colors):
    if color is None:
        return None
    typ = getattr(color, 'type', None)
    if typ == 'rgb' and getattr(color, 'rgb', None):
        return '#' + color.rgb[-6:]
    if typ == 'theme' and getattr(color, 'theme', None) is not None and theme_colors:
        base = theme_colors[color.theme] if color.theme < len(theme_colors) else '000000'
        return '#' + _apply_tint(base, getattr(color, 'tint', 0) or 0)
    if typ == 'indexed' and getattr(color, 'indexed', None) in _INDEXED_COLORS:
        return '#' + _INDEXED_COLORS[color.indexed]
    return None


def _excel_col_name(idx):
    name = ''
    n = idx + 1
    while n:
        n, rem = divmod(n - 1, 26)
        name = chr(65 + rem) + name
    return name


def _row_to_visible_area(row_idx, header_idx):
    if header_idx < 0:
        return ('body', row_idx)
    if row_idx == header_idx:
        return ('header', 0)
    body_idx = row_idx
    if row_idx > header_idx:
        body_idx -= 1
    return ('body', body_idx)


def _parse_sheet_merges(ws, header_idx):
    merges = []
    for merged in ws.merged_cells.ranges:
        min_row = merged.min_row - 1
        max_row = merged.max_row - 1
        start_area, start_visible_row = _row_to_visible_area(min_row, header_idx)
        end_area, end_visible_row = _row_to_visible_area(max_row, header_idx)
        if not start_area or start_area != end_area:
            continue
        if start_area == 'header':
            if merged.min_row != merged.max_row:
                continue
            merges.append({
                'area': 'header',
                'row': 0,
                'col': merged.min_col - 1,
                'rowspan': 1,
                'colspan': merged.max_col - merged.min_col + 1,
            })
            continue
        merges.append({
            'area': 'body',
            'row': start_visible_row,
            'col': merged.min_col - 1,
            'rowspan': end_visible_row - start_visible_row + 1,
            'colspan': merged.max_col - merged.min_col + 1,
        })
    return merges


def _merge_ref_for_sheet(sheet, merge):
    area = merge.get('area', 'body')
    row = int(merge.get('row', 0) or 0)
    col = int(merge.get('col', 0) or 0)
    rowspan = max(1, int(merge.get('rowspan', 1) or 1))
    colspan = max(1, int(merge.get('colspan', 1) or 1))
    header_idx_raw = sheet.get('headerRowIndex', 0)
    header_idx = int(header_idx_raw) if header_idx_raw not in (None, '') else 0
    promoted_header = header_idx >= 0
    if area == 'header' and promoted_header:
        start_row = header_idx + 1
    else:
        if not promoted_header:
            start_row = row + 1
        else:
            start_row = row + 1 if row < header_idx else row + 2
    end_row = start_row + rowspan - 1
    start_col = col + 1
    end_col = start_col + colspan - 1
    return f'{_excel_col_name(start_col - 1)}{start_row}:{_excel_col_name(end_col - 1)}{end_row}'


def _parse_xlsx_sheets_with_styles(raw_bytes):
    if not _XLSX_OK:
        raise RuntimeError('openpyxl not installed — run: pip install openpyxl')
    wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
    theme_colors = _parse_theme_colors(wb)
    image_map = _parse_xlsx_images(raw_bytes)
    sheets = []
    for ws in wb.worksheets:
        max_row = ws.max_row or 0
        max_col = ws.max_column or 0
        rows2d = []
        styles2d = []
        for r in range(1, max_row + 1):
            row_vals = []
            row_styles = []
            for c in range(1, max_col + 1):
                cell = ws.cell(r, c)
                row_vals.append(_cell_to_str(cell.value))
                st = {}
                fill = cell.fill
                font = cell.font
                align = cell.alignment
                if getattr(fill, 'fill_type', None) and fill.fill_type != 'none':
                    bg = _resolve_openpyxl_color(fill.fgColor, theme_colors) or _resolve_openpyxl_color(fill.bgColor, theme_colors)
                    if bg:
                        st['bg'] = bg
                fc = _resolve_openpyxl_color(getattr(font, 'color', None), theme_colors)
                if fc:
                    st['color'] = fc
                if getattr(font, 'bold', False):
                    st['bold'] = True
                if getattr(font, 'italic', False):
                    st['italic'] = True
                if getattr(font, 'sz', None):
                    st['fontSize'] = font.sz
                if getattr(font, 'name', None):
                    st['fontFamily'] = font.name
                if getattr(align, 'horizontal', None):
                    st['align'] = align.horizontal
                row_styles.append(st or None)
            rows2d.append(row_vals)
            styles2d.append(row_styles)
        while rows2d and all(v == '' for v in rows2d[-1]) and all(s is None for s in styles2d[-1]):
            rows2d.pop()
            styles2d.pop()
        if rows2d:
            nonempty_counts = [sum(1 for v in row if v != '') for row in rows2d]
            header_idx = max(range(len(rows2d)), key=lambda i: (nonempty_counts[i], -i))
            if nonempty_counts[header_idx] == 0:
                header_idx = 0
        else:
            header_idx = 0
        has_preamble = header_idx > 0
        if has_preamble:
            headers = [_excel_col_name(i) for i in range(max_col)]
            header_styles = []
            leading_rows = rows2d
            trailing_rows = []
            leading_styles = styles2d
            trailing_styles = []
            header_idx_out = -1
        else:
            headers = rows2d[header_idx] if rows2d else []
            header_styles = styles2d[header_idx] if styles2d else []
            leading_rows = rows2d[:header_idx]
            trailing_rows = rows2d[header_idx + 1:] if len(rows2d) > header_idx + 1 else []
            leading_styles = styles2d[:header_idx]
            trailing_styles = styles2d[header_idx + 1:] if len(styles2d) > header_idx + 1 else []
            header_idx_out = header_idx
        sheets.append({
            'name': ws.title,
            'headers': headers,
            'rows': leading_rows + trailing_rows,
            'headerStyles': header_styles,
            'rowStyles': leading_styles + trailing_styles,
            'images': image_map.get(ws.title, []),
            'headerRowIndex': header_idx_out,
            'merges': _parse_sheet_merges(ws, header_idx_out),
        })
    return sheets


def _ordered_sheet_rows(sheet):
    header_idx = int(sheet.get('headerRowIndex', 0) or 0)
    if header_idx < 0:
        return [list(r) for r in sheet.get('rows', [])]
    rows = list(sheet.get('rows', []))
    header = list(sheet.get('headers', []))
    leading = rows[:header_idx]
    trailing = rows[header_idx:]
    return leading + [header] + trailing


def _set_xml_cell_value(cell_el, value):
    original_t = cell_el.attrib.get('t')
    for child in list(cell_el):
        if child.tag.endswith(('v', 'is', 'f')):
            cell_el.remove(child)
    if value == '' or value is None:
        cell_el.attrib.pop('t', None)
        return
    py_val = _str_to_cell(value) if isinstance(value, str) else value
    force_string = original_t in ('s', 'str', 'inlineStr')
    if isinstance(value, str):
        stripped = value.strip()
        has_ambiguous_leading_zero = (
            stripped.startswith('0')
            and len(stripped) > 1
            and stripped[1:].isdigit()
        ) or (
            stripped.startswith('-0')
            and len(stripped) > 2
            and stripped[2:].isdigit()
        )
        if has_ambiguous_leading_zero:
            force_string = True
    if isinstance(py_val, bool):
        cell_el.set('t', 'b')
        v = _ET.SubElement(cell_el, cell_el.tag[:-1] + 'v')
        v.text = '1' if py_val else '0'
        return
    if isinstance(py_val, (int, float)) and not force_string:
        cell_el.attrib.pop('t', None)
        v = _ET.SubElement(cell_el, cell_el.tag[:-1] + 'v')
        v.text = str(py_val)
        return
    cell_el.set('t', 'inlineStr')
    is_el = _ET.SubElement(cell_el, cell_el.tag[:-1] + 'is')
    t_el = _ET.SubElement(is_el, cell_el.tag[:-1] + 't')
    if isinstance(value, str) and (value.startswith(' ') or value.endswith(' ') or '\n' in value):
        t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t_el.text = '' if value is None else str(value)


def _write_xlsx_from_template(raw_bytes, sheets):
    ns = {
        'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    }
    by_name = {sh.get('name', ''): sh for sh in sheets}
    zin = io.BytesIO(raw_bytes)
    zout = io.BytesIO()
    with _zipfile.ZipFile(zin, 'r') as src, _zipfile.ZipFile(zout, 'w', _zipfile.ZIP_DEFLATED) as dst:
        wb = _ET.fromstring(src.read('xl/workbook.xml'))
        wb_rels = _ET.fromstring(src.read('xl/_rels/workbook.xml.rels'))
        rel_targets = {rel.attrib['Id']: rel.attrib['Target'] for rel in wb_rels.findall('{*}Relationship')}
        sheet_paths = {}
        for sheet in wb.findall('a:sheets/a:sheet', ns):
            rid = sheet.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            target = rel_targets.get(rid)
            if target:
                sheet_paths[sheet.attrib.get('name', '')] = _ppath.normpath('xl/' + target)

        for name in src.namelist():
            data = src.read(name)
            if name in sheet_paths.values():
                sheet_name = next((k for k, v in sheet_paths.items() if v == name), None)
                sheet = by_name.get(sheet_name or '')
                if sheet:
                    root = _ET.fromstring(data)
                    sheet_data = root.find('a:sheetData', ns)
                    if sheet_data is not None:
                        rows = _ordered_sheet_rows(sheet)
                        max_row = max(len(rows), len(sheet_data.findall('a:row', ns)))
                        max_col = max([len(r) for r in rows] + [0])
                        row_map = {
                            int(row.attrib.get('r', idx + 1)): row
                            for idx, row in enumerate(sheet_data.findall('a:row', ns))
                        }
                        cell_maps = {}
                        for rnum, row_el in row_map.items():
                            cmap = {}
                            for cell in row_el.findall('a:c', ns):
                                ref = cell.attrib.get('r', '')
                                col = ''.join(ch for ch in ref if ch.isalpha())
                                cmap[col] = cell
                            cell_maps[rnum] = cmap
                        for child in list(sheet_data):
                            sheet_data.remove(child)
                        for r in range(1, max_row + 1):
                            old_row = row_map.get(r)
                            row_el = _ET.Element(old_row.tag if old_row is not None else f'{{{ns["a"]}}}row')
                            if old_row is not None:
                                row_el.attrib.update(old_row.attrib)
                            row_el.set('r', str(r))
                            values = rows[r - 1] if r - 1 < len(rows) else []
                            old_cells = cell_maps.get(r, {})
                            for c in range(max(max_col, len(values))):
                                col_name = _excel_col_name(c)
                                ref = f'{col_name}{r}'
                                old_cell = old_cells.get(col_name)
                                val = values[c] if c < len(values) else ''
                                if old_cell is None and (val == '' or val is None):
                                    continue
                                cell_el = _ET.fromstring(_ET.tostring(old_cell)) if old_cell is not None else _ET.Element(f'{{{ns["a"]}}}c')
                                cell_el.set('r', ref)
                                _set_xml_cell_value(cell_el, val)
                                row_el.append(cell_el)
                            sheet_data.append(row_el)
                        dim = root.find('a:dimension', ns)
                        if dim is not None and max_row and max_col:
                            dim.set('ref', f'A1:{_excel_col_name(max_col - 1)}{max_row}')
                        merge_cells = root.find('a:mergeCells', ns)
                        if merge_cells is not None:
                            root.remove(merge_cells)
                        merge_refs = []
                        for merge in sheet.get('merges', []) or []:
                            try:
                                merge_refs.append(_merge_ref_for_sheet(sheet, merge))
                            except Exception:
                                continue
                        if merge_refs:
                            merge_cells = _ET.Element(f'{{{ns["a"]}}}mergeCells')
                            merge_cells.set('count', str(len(merge_refs)))
                            for ref in merge_refs:
                                merge_el = _ET.SubElement(merge_cells, f'{{{ns["a"]}}}mergeCell')
                                merge_el.set('ref', ref)
                            insert_idx = list(root).index(sheet_data) + 1
                            root.insert(insert_idx, merge_cells)
                        data = _ET.tostring(root, encoding='utf-8', xml_declaration=True)
            dst.writestr(name, data)
    return zout.getvalue()


def _parse_xlsx_images(raw_bytes):
    ns = {
        'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
        'd': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'pr': 'http://schemas.openxmlformats.org/package/2006/relationships',
    }
    out = {}
    try:
        with _zipfile.ZipFile(io.BytesIO(raw_bytes)) as zf:
            wb = _ET.fromstring(zf.read('xl/workbook.xml'))
            wb_rels = _ET.fromstring(zf.read('xl/_rels/workbook.xml.rels'))
            rel_targets = {rel.attrib['Id']: rel.attrib['Target'] for rel in wb_rels.findall('{*}Relationship')}
            for sheet in wb.findall('a:sheets/a:sheet', ns):
                name = sheet.attrib.get('name', '')
                rid = sheet.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                target = rel_targets.get(rid)
                if not target:
                    continue
                sheet_path = _ppath.normpath('xl/' + target)
                rels_path = _ppath.normpath(_ppath.join(_ppath.dirname(sheet_path), '_rels', _ppath.basename(sheet_path) + '.rels'))
                if rels_path not in zf.namelist():
                    continue
                sheet_rels = _ET.fromstring(zf.read(rels_path))
                sheet_rel_targets = {rel.attrib['Id']: rel.attrib['Target'] for rel in sheet_rels.findall('{*}Relationship')}
                sheet_xml = _ET.fromstring(zf.read(sheet_path))
                drawing = sheet_xml.find('a:drawing', ns)
                if drawing is None:
                    continue
                drawing_rid = drawing.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                drawing_target = sheet_rel_targets.get(drawing_rid)
                if not drawing_target:
                    continue
                drawing_path = _ppath.normpath(_ppath.join(_ppath.dirname(sheet_path), drawing_target))
                drawing_rels_path = _ppath.normpath(_ppath.join(_ppath.dirname(drawing_path), '_rels', _ppath.basename(drawing_path) + '.rels'))
                if drawing_path not in zf.namelist() or drawing_rels_path not in zf.namelist():
                    continue
                drawing_xml = _ET.fromstring(zf.read(drawing_path))
                drawing_rels = _ET.fromstring(zf.read(drawing_rels_path))
                drawing_rel_targets = {rel.attrib['Id']: rel.attrib['Target'] for rel in drawing_rels.findall('{*}Relationship')}
                images = []
                for anchor in drawing_xml.findall('xdr:twoCellAnchor', ns) + drawing_xml.findall('xdr:oneCellAnchor', ns):
                    blip = anchor.find('.//d:blip', ns)
                    if blip is None:
                        continue
                    embed = blip.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                    media_target = drawing_rel_targets.get(embed)
                    if not media_target:
                        continue
                    media_path = _ppath.normpath(_ppath.join(_ppath.dirname(drawing_path), media_target))
                    if media_path not in zf.namelist():
                        continue
                    data = zf.read(media_path)
                    ext = _ppath.splitext(media_path)[1].lower()
                    mime = {
                        '.jpg': 'image/jpeg',
                        '.jpeg': 'image/jpeg',
                        '.png': 'image/png',
                        '.gif': 'image/gif',
                    }.get(ext, 'application/octet-stream')
                    fr = anchor.find('xdr:from', ns)
                    if fr is not None:
                        row = int(fr.findtext('xdr:row', default='0', namespaces=ns))
                        col = int(fr.findtext('xdr:col', default='0', namespaces=ns))
                    else:
                        row = col = 0
                    images.append({
                        'row': row,
                        'col': col,
                        'src': f'data:{mime};base64,' + base64.b64encode(data).decode('ascii'),
                    })
                if images:
                    out[name] = images
    except Exception:
        return {}
    return out


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
                try:
                    sheets = _parse_xlsx_sheets_with_styles(raw_bytes)
                    self._json({
                        'sheets': sheets,
                        'rawB64': base64.b64encode(raw_bytes).decode('ascii'),
                        'sha': sha,
                        'format': 'xlsx',
                    })
                except Exception:
                    # Fallback to raw bytes if server-side style parsing fails
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
        if self.path == '/api/parse-xlsx':
            if not _XLSX_OK:
                self._json({'error': 'openpyxl not installed — run: pip install openpyxl'}); return
            try:
                sheets = _parse_xlsx_sheets_with_styles(body)
            except Exception as e:
                self._json({'error': str(e)}); return
            self._json({'ok': True, 'sheets': sheets})
            return
        if self.path == '/api/build-xlsx-from-template':
            try:
                raw = _write_xlsx_from_template(body, state.sheets)
            except Exception as e:
                self._json({'error': str(e)}); return
            self.send_response(200)
            self.send_header('Content-Type',
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Length', str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
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
            active = state.sheets[state.active_sheet]
            active['merges'] = data.get('merges', active.get('merges', []))
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
