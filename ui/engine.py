# -*- coding: utf-8 -*-
"""
A small, dependency-free spreadsheet engine.

It mimics exactly the Excel surface the converted legacy workbooks talk
through (ActiveWorkbook / ActiveSheet / ActiveCell / Cells / Range / named
ranges / Offset / WorksheetFunction), so the tkinter forms below can read and
write records the same way the VBA did -- no Excel, no pandas needed.

The convention used by the Smart Coast tools:
    row 1 : column captions
    row 2 : internal variable names
    row 3 : data types (char / list / date / flag / url / database / ...)
    row 4+: data records
"""
import csv
import os
import re
import time


def _to_int(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return 1


class Cell(object):
    """One spreadsheet cell (1-based row/col like VBA)."""
    __slots__ = ('sheet', 'row', 'col')

    def __init__(self, sheet, row, col):
        self.sheet = sheet
        self.row = row
        self.col = col

    def _key(self):
        return (self.row, self.col)

    @property
    def Value(self):
        return self.sheet._data.get(self._key(), '')

    @Value.setter
    def Value(self, v):
        self.sheet._data[self._key()] = '' if v is None else v

    @property
    def Text(self):
        return '' if self.Value is None else str(self.Value)

    @property
    def Column(self):
        return self.col

    @property
    def Row(self):
        return self.row

    def Offset(self, dr=0, dc=0):
        return Cell(self.sheet, self.row + _to_int(dr), self.col + _to_int(dc))

    @property
    def EntireRow(self):
        return RowSpan(self.sheet, self.row)


class RowSpan(object):
    """A full spreadsheet row (covers .EntireRow.Range('A1') idioms)."""
    def __init__(self, sheet, row):
        self.sheet = sheet
        self.row = row

    def Range(self, _a1):
        return Cell(self.sheet, self.row, 1)


class Range(object):
    """A rectangular region; .Value is scalar for 1x1 else list-of-lists."""
    def __init__(self, sheet, r1, c1, r2=None, c2=None):
        self.sheet = sheet
        self.r1, self.c1 = r1, c1
        self.r2 = r1 if r2 is None else r2
        self.c2 = c1 if c2 is None else c2

    @property
    def Value(self):
        if self.r1 == self.r2 and self.c1 == self.c2:
            return self.sheet._data.get((self.r1, self.c1), '')
        rows = []
        for r in range(self.r1, self.r2 + 1):
            rows.append([self.sheet._data.get((r, c), '')
                         for c in range(self.c1, self.c2 + 1)])
        return rows

    @Value.setter
    def Value(self, v):
        if self.r1 == self.r2 and self.c1 == self.c2:
            self.sheet._data[(self.r1, self.c1)] = v
            return
        for i, r in enumerate(range(self.r1, self.r2 + 1)):
            rowvals = v[i] if isinstance(v, (list, tuple)) else v
            if not isinstance(rowvals, (list, tuple)):
                rowvals = [rowvals] * (self.c2 - self.c1 + 1)
            for j, c in enumerate(range(self.c1, self.c2 + 1)):
                self.sheet._data[(r, c)] = rowvals[j] if j < len(rowvals) else ''

    @property
    def Text(self):
        v = self.Value
        return '' if v is None else str(v)

    def __getitem__(self, idx):
        r, c = idx
        return Cell(self.sheet, self.r1 + r, self.c1 + c)


class Sheet(object):
    """Named worksheet with named-range lookups."""
    def __init__(self, name, rows=20, cols=30):
        self.name = name
        self._data = {}
        self.named = {}
        self._rows = rows
        self._cols = cols

    def Cells(self, r, c=1):
        return Cell(self, _to_int(r), _to_int(c))

    def Range(self, a, b=None):
        if isinstance(a, str):
            if a in self.named:
                r1, c1, r2, c2 = self.named[a]
                return Range(self, r1, c1, r2, c2)
            m = re.fullmatch(r'([A-Z]+)(\d+)', a)
            if m:
                col = self._colnum(m.group(1))
                return Range(self, _to_int(m.group(2)), col)
            return Range(self, 1, 1)
        if isinstance(a, Cell) and isinstance(b, Cell):
            return Range(self, a.row, a.col, b.row, b.col)
        return Range(self, 1, 1, 1, 1)

    def define_name(self, name, r1, c1, r2=None, c2=None):
        self.named[name] = (r1, c1, r2, c2)

    @property
    def UsedRange(self):
        rs = [k[0] for k in self._data]
        cs = [k[1] for k in self._data]
        r2 = max(rs) if rs else self._rows
        c2 = max(cs) if cs else self._cols
        return Range(self, 1, 1, r2, c2)

    @property
    def Name(self):
        return self.name

    def set_cell(self, r, c, v):
        self._data[(r, c)] = v

    def get_cell(self, r, c, default=''):
        return self._data.get((r, c), default)

    def nrows(self):
        rs = [k[0] for k in self._data]
        return (max(rs) if rs else self._rows)

    def ncols(self):
        cs = [k[1] for k in self._data]
        return (max(cs) if cs else self._cols)

    @staticmethod
    def _colnum(letters):
        n = 0
        for ch in letters.upper():
            n = n * 26 + (ord(ch) - 64)
        return n


class Workbook(object):
    def __init__(self, name=None):
        self.name = name or 'Wb'
        self.sheets = []

    def add_sheet(self, sheet):
        self.sheets.append(sheet)
        return sheet

    def Sheets(self, name=None):
        if name is None:
            return self.sheets
        for s in self.sheets:
            if s.name == name:
                return s
        sh = Sheet(name)
        self.sheets.append(sh)
        return sh

    def ActiveSheet(self):
        return self.sheets[0] if self.sheets else Sheet('Sheet1')


class Application(object):
    def __init__(self):
        self._EnableCancelKey = 0
        self._WaitUntil = None

    def EnableCancelKey(self, *a):           # property-ish access
        if a:
            self._EnableCancelKey = a[0]
        return self._EnableCancelKey

    def Wait(self, Time=None, **kw):
        return None

    def OnKey(self, *a, **k):
        return None


class WF(object):
    """The handful of WorksheetFunction calls the converted code uses."""
    def Max(self, *a, **k):
        vals = self._flatten(a)
        return max(vals) if vals else 0

    def CountIf(self, rng, crit):
        vals = self._flatten([rng])
        if isinstance(crit, str) and crit:
            op, tail = crit[0], crit[1:]
            if op == '>':
                return sum(1 for v in vals if str(v) > tail)
            if op == '<':
                return sum(1 for v in vals if str(v) < tail)
        return sum(1 for v in vals if str(v) == str(crit))

    def Transpose(self, v):
        if not isinstance(v, list):
            return v
        return [v[i][0] for i in range(len(v))] if v and isinstance(v[0], list) else v

    def _flatten(self, a):
        out = []
        for x in a:
            if isinstance(x, Range):
                v = x.Value
                out.extend(self._flatten([v]))
            elif isinstance(x, (list, tuple)):
                for y in x:
                    out.extend(self._flatten([y]))
            else:
                try:
                    out.append(float(x))
                except (TypeError, ValueError):
                    pass
        return out


class AppState(object):
    """Composes the workbook + active-cell state the legacy code assumes."""

    def __init__(self):
        self.app = Application()
        self.wf = WF()
        self.wb = None
        self.awb = None
        self.asheet = None
        self.acell = None
        self.sel = None
        self._row = 4          # current data record row
        self.init_workbooks()

    def init_workbooks(self):
        self.wb = Workbook('Smc')
        self.awb = Workbook('Smc')
        self.asheet = self.wb.add_sheet(Sheet('SMCCopy'))
        self.acell = self.asheet.Cells(4, 1)
        self.sel = self.asheet.Cells(4, 1)

    def active_row(self):
        return self.acell.row if self.acell else self._row

    def set_active_row(self, r):
        r = max(1, int(r))
        self._row = r
        if self.acell is not None:
            self.acell = self.asheet.Cells(r, 1)
            self.sel = self.asheet.Cells(r, 1)


# ---------------------------------------------------------------------------
# helpers used by the converted business code (RecastLat, CustomFileName, ...)
# ---------------------------------------------------------------------------

def strip_spaces(s):
    if s is None:
        return ''
    return ''.join(str(s).split())


def number_char(s):
    m = re.match(r'^(\d+)', str(s or '').strip())
    return int(m.group(1)) if m else 0


def is_empty(v):
    return v is None or str(v).strip() == ''


def latlon2dec(s):
    """Parse 'DD MM SS' / 'DD.MMSS' or plain decimal to decimal degrees."""
    if s is None:
        return 0.0
    s = str(s).strip()
    if s == '':
        return 0.0
    m = re.match(r'^([+-]?\d+)\s+(\d+)\s+([\d.]+)$', s)
    if m:
        deg, mn, sec = map(float, m.groups())
        sign = -1 if deg < 0 else 1
        return sign * (abs(deg) + mn / 60.0 + sec / 3600.0)
    try:
        return float(s)
    except ValueError:
        return 0.0


def dec2latlon(v):
    v = float(v or 0)
    sign = -1 if v < 0 else 1
    v = abs(v)
    deg = int(v)
    mn = int((v - deg) * 60)
    sec = (v - deg) * 3600 - mn * 60
    return '%d %02d %06.3f' % (sign * deg, mn, sec)


def recast_dec_lat(v):
    return dec2latlon(latlon2dec(v))


def recast_dec_lon(v):
    return dec2latlon(latlon2dec(v))


def custom_file_name(old):
    """Best-effort port of the make-file-name helper (uppercase, safe chars)."""
    if old is None:
        return ''
    s = re.sub(r'[^\w.\- ]', '', str(old)).strip().replace(' ', '_')
    return s.upper()


# ---------------------------------------------------------------------------
# attaching to the runtime module so the generated .py shares the same world
# ---------------------------------------------------------------------------

def attach(vba_module):
    """Point a vba_runtime module's Excel globals at a fresh AppState."""
    if 'smc_engine' not in vba_module.__dict__:
        vba_module.smc_engine = AppState()
    st = vba_module.smc_engine
    vba_module.wb = st.wb
    vba_module.awb = st.awb
    vba_module.asheet = st.asheet
    vba_module.acell = st.acell
    vba_module.sel = st.sel
    vba_module.app = st.app
    vba_module.wf = st.wf
    return st


# ---------------------------------------------------------------------------
# sample data so the rebuilt forms open with real-looking columns
# ---------------------------------------------------------------------------

SC_COLUMNS = [
    ('ID', 'log_id', 'integer'),
    ('Status', 'status', 'status'),
    ('Citation Key', 'short_authors', 'list'),
    ('Full Citation Key', 'long_authors', 'list'),
    ('Book Editors', 'book_editors', 'list'),
    ('Database Record', 'database', 'database'),
    ('Publication Date', 'pub_date', 'date'),
    ('Data URL', 'data_url', 'url'),
    ('Insert Author', 'insert', 'char'),
    ('Last Update Author', 'update', 'char'),
    ('Reference Type', 'ref_type', 'char'),
    ('Latitude (deg N)', 'lat1', 'lat'),
    ('Longitude (deg E)', 'lon1', 'lon'),
    ('Sample ID', 'sample_ids', 'flag'),
    ('Region Name', 'region_name', 'char'),
    ('Search Key', 'search_key', 'skip'),
    ('DMO Flag', 'dmo_flag', 'flag'),
    ('Convention', 'convention', 'list'),
    ('GMT Scale', 'gmt_scale', 'char'),
    ('Remarks', 'remarks', 'char'),
    ('Created Stamp', 'created_stamp', 'date'),
    ('Compute Flag', 'compute_flag', 'flag'),
    ('Grid Size (m)', 'grid_size_meter', 'integer'),
    ('Pacer log', 'pacer_log_ids', 'list'),
]

SC_RECORDS = [
    ['2026-0001', 'New ...', 'Smith et al.', 'Smith AG, Jones PL', '',
     '', '2025-03-11', 'https://example.org/ds/2026-0001', '', '', 'Dataset',
     '31.5', '-118.2', 'S0001, S0002', 'California Borderlands', '', 'N',
     'AGU', '1:', 'Coordinates from cruise log.', '2026-01-02', 'Y', '250', 'l1'],
    ['2026-0002', 'Frozen ...', 'Jones and Lee', 'Jones RL, Lee TM', '',
     '', '2024-11-02', '', '', '', 'Cruise', '-42.1', '170.5',
     'S0010', 'Tasman Sea', '', 'Y', 'N/A', '1:', '', '2025-12-01', 'N',
     '250', 'l2'],
]

ERDPP_COLUMNS = [
    ('ID', 'erdpp_id', 'integer'),
    ('Status', 'status', 'status'),
    ('ER Mail ID', 'er_mail_id', 'char'),
    ('ER Level ID', 'er_level_id', 'char'),
    ('Inserted', 'inserted', 'date'),
    ('Primary Authors', 'long_authors', 'list'),
    ('Book Editors', 'book_editors', 'list'),
    ('Insert Author', 'insert', 'char'),
    ('Update Author', 'update', 'char'),
    ('Database URL', 'database_url', 'url'),
    ('Reference Type', 'ref_type', 'char'),
    ('GMT Scale', 'gmt_scale', 'char'),
    ('Copyright', 'copyright', 'flag'),
    ('Insertion Author', 'insert_flag', 'flag'),
]

ERDPP_RECORDS = [
    ["E-001", "New ...", "er@mydomain.org", "L1", "2025-01-15", "O'Hara B",
     '', '', '', '', 'Preprint', '1:', 'N', ''],
]

# Map options used by the map setup form (UserForm15) Y/N matrix.
MAP_OPTIONS = [
    'Amplitude Anomaly', 'Vertical Gravity Gradient', 'Predicted Bathymetry',
    'Multibeam Bathymetry', 'Sediment Thickness', 'Free Air Gravity',
    'Mercator', 'Pseudomercator', 'Polar Stereographic', 'Contour Labels',
    'Gridlines', 'Frame Line', 'Scalebar', 'N-S Arrow', 'Logo', 'Caption',
]


def build_sample_state():
    """Create the SC style workbook with columns + records + named ranges."""
    st = AppState()
    sheet = st.asheet
    # header block (rows 1..3)
    for j, (caption, var, dtyp) in enumerate(SC_COLUMNS, start=1):
        sheet.set_cell(1, j, caption)
        sheet.set_cell(2, j, var)
        sheet.set_cell(3, j, dtyp)
    for i, rec in enumerate(SC_RECORDS, start=4):
        for j, v in enumerate(rec, start=1):
            sheet.set_cell(i, j, v)
    sheet.define_name('tagnames', 2, 1, 2, len(SC_COLUMNS))
    # map sheet for UserForm15
    maps = st.awb.add_sheet(Sheet('SC_maps'))
    headers = ['FILE_NAME', 'FOLDER', 'FILE_CAPTION', 'CREATION_STATUS',
               'CREATION_TYPE', 'GRID_LAT1', 'GRID_LON1', 'GRID_LAT2',
               'GRID_LON2', 'GRID_SIZE_METER', 'GRID_SIZE_MIN', 'MAP_LAT1',
               'MAP_LON1', 'MAP_LAT2', 'MAP_LON2', 'GMT_SCALE', 'GMT_CI1',
               'GMT_CI2', 'GMT_SUN', 'SEAMOUNT_INDEX', 'SEAMOUNT_NAME',
               'REGION_NAME', 'GRIDLINE_INTERVAL', 'LABEL_INTERVAL',
               'FRAME_INTERVAL', 'CPT_MAP', 'CPT_SCALEBAR', 'MAP_OPTIONS',
               'GMT_NORMALIZATION', 'OCEAN_NAME', 'UPDATED', 'PACER_LOG_IDS']
    for j, h in enumerate(headers, start=1):
        maps.set_cell(1, j, h)
        maps.set_cell(2, j, h.lower().replace('_', ' '))
    row = {h.lower(): '' for h in headers}
    row.update({'file_name': 'SAMPLE_GRAVITY', 'file_caption': 'Sample gravity',
                'creation_type': 'SSGRID', 'map_lat1': '30 00 00',
                'map_lon1': '-120 00 00', 'map_lat2': '33 00 00',
                'map_lon2': '-116 00 00', 'map_scale': '1:', 'gmt_scale': '1:',
                'grid_size_meter': '250', 'seamount_name': 'None',
                'region_name': 'California Borderlands',
                'ocean_name': 'Pacific', 'map_options': 'YYYNNNNNNNNNNNNN'})
    for j, h in enumerate(headers, start=1):
        maps.set_cell(4, j, row.get(h.lower(), ''))
    st.set_active_row(4)
    return st


def load_csv_sheet(sheet, path):
    with open(path, newline='', encoding='utf-8-sig') as fh:
        for r, row in enumerate(csv.reader(fh), start=1):
            for c, v in enumerate(row, start=1):
                sheet.set_cell(r, c, v)


def save_csv_sheet(sheet, path):
    nrows = sheet.nrows()
    ncols = sheet.ncols()
    with open(path, 'w', newline='', encoding='utf-8') as fh:
        w = csv.writer(fh)
        for r in range(1, nrows + 1):
            w.writerow([sheet.get_cell(r, c) for c in range(1, ncols + 1)])