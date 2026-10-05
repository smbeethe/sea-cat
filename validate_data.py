# -*- coding: utf-8 -*-
"""
Data validation for the Sea-Cat legacy Excel workbooks.

The tools (SC.console, ERDPP.console) store every table with a standard
three-row header block:

    row 1 : column captions  (e.g. "Seamount\nID")
    row 2 : field names      (e.g. "sc_seamount_id")
    row 3 : data types, which double as validation rules:
               id | long | nmb | date | year | skip
               varchar2(n) | char(n) | list(n)
    row 4+ : the data records.

This script re-implements those rules in plain Python and additionally
cross-checks the *hierarchies* between the tables:

    * SC_maps  -> SC_seamounts / SC_regions        (index-name references)
    * MGG_samples -> MGG_expeditions               (expedition names)
    * ER_expertlevels.pacer_log_ids -> PACER_loggers
    * PACER_logsheet.pacer_log_ids  -> PACER_loggers

Usage:
    python validate_data.py [path.xls ...] [--report out.txt] [--limit N]

Default input: scdpp.xls and log.xls from the SMCCopy source folder.
"""
import argparse
import re
import sys
import time
from collections import defaultdict

import xlrd

SCHEMA_ROWS = 3          # captions / field names / datatypes
DATA_START = 4           # first data row (1-based Excel row)

_DEFAULT_SRC = r'C:\Users\beethes.ONID\sea-cat\SMC\SMCCopy'
_DEFAULT_FILES = ['scdpp.xls', 'log.xls']

# ---------------------------------------------------------------------------
# schema parsing
# ---------------------------------------------------------------------------

_TYPE_RE = re.compile(
    r'^(?P<kind>id|long|nmb|date|year|skip|varchar2|char|list)'
    r'(?:\((?P<arg>\d+)\))?$')


def parse_type(tok):
    """Return (kind, arg) for a row-3 datatype token, or None if unknown."""
    if tok is None:
        return None
    tok = str(tok).strip().lower()
    if not tok:
        return None
    m = _TYPE_RE.match(tok)
    if not m:
        return None
    arg = int(m.group('arg')) if m.group('arg') else None
    return (m.group('kind'), arg)


def field_name(raw):
    """Normalise a row-2 label into an attribute-ish key."""
    if raw is None:
        return ''
    return re.sub(r'[^a-z0-9_]+', '_', str(raw).strip().lower()).strip('_')


class Column(object):
    __slots__ = ('var', 'kind', 'arg', 'caption')

    def __init__(self, var, kind, arg, caption):
        self.var = var
        self.kind = kind
        self.arg = arg
        self.caption = caption


class Table(object):
    """One worksheet reconstructed as (schema, rows)."""

    def __init__(self, name):
        self.name = name
        self.columns = []
        self.colmap = {}       # normalized var name -> Column
        self.rows = []         # list of dicts var -> raw cell value
        self.data_start = None  # recorded Excel row number of first data row

    def n_data(self):
        return len(self.rows)


# ---------------------------------------------------------------------------
# loading
# ---------------------------------------------------------------------------

def load_table(sheet):
    """Reconstruct a Table from an xlrd sheet using the 3-row convention."""
    tbl = Table(sheet.name)
    if sheet.nrows < DATA_START:
        return None
    caps = [sheet.cell_value(0, c) for c in range(sheet.ncols)]
    vars_ = [field_name(sheet.cell_value(1, c)) for c in range(sheet.ncols)]
    types = [sheet.cell_value(2, c) for c in range(sheet.ncols)]

    for c in range(sheet.ncols):
        pt = parse_type(types[c])
        if pt is None and c >= sheet.ncols:
            continue
        col = Column(vars_[c], pt[0] if pt else 'skip',
                     pt[1] if pt else None, caps[c])
        tbl.columns.append(col)
        if vars_[c]:
            tbl.colmap[vars_[c]] = col

    # only keep sheets that actually carry an id-like key or a datatype row
    if not tbl.colmap:
        return None
    if not any(k and (k.endswith('_id') or k == 'id')
               for k in tbl.colmap):
        # still keep: some tables may key differently; require a datatype row
        if not any(parse_type(t) for t in types):
            return None

    tbl.data_start = DATA_START
    # read data rows into dicts (skip fully-empty rows)
    for r in range(DATA_START - 1, sheet.nrows):
        row = {}
        any_val = False
        for c, col in enumerate(tbl.columns):
            v = sheet.cell_value(r, c)
            if v != '':
                any_val = True
            row[col.var or ('c%d' % (c + 1))] = v
        if any_val:
            rownum = r + 1
            row['_row'] = rownum
            row['_super'] = None
            tbl.rows.append(row)
    return tbl


def load_workbook(path, limit=None):
    """Return {sheetname: Table} for a legacy workbook."""
    wb = xlrd.open_workbook(path, on_demand=True, formatting_info=False)
    tables = {}
    for sh in wb.sheets():
        if limit and sh.nrows > DATA_START and sh.nrows - DATA_START + 1 > limit:
            # shallow copy with truncated rows
            tbl = load_table(sh)
            if tbl:
                tbl.rows = tbl.rows[:min(limit, 0) or limit]
                tables[sh.name] = tbl
            continue
        tbl = load_table(sh)
        if tbl:
            tables[sh.name] = tbl
    wb.release_resources()
    return tables


# ---------------------------------------------------------------------------
# validation
# ---------------------------------------------------------------------------

class Report(object):
    def __init__(self):
        self.issues = []          # (table, excel_row, var, category, message)

    def add(self, table, row, var, category, message):
        self.issues.append((table, row, var, category, message))

    def counts(self):
        c = defaultdict(int)
        for _t, _r, _v, cat, _m in self.issues:
            c[cat] += 1
        return c

    def table_counts(self):
        c = defaultdict(int)
        for t, _r, _v, _c, _m in self.issues:
            c[t] += 1
        return c


def _is_empty(v):
    return v is None or (isinstance(v, float) and v != v) or v == ''


def _celltext(v, cap=80):
    s = str(v)
    return s if len(s) <= cap else s[:cap - 1] + '~'


def _excel_str(v):
    """Display value the way Excel would (1.0 -> '1', 0.5 -> '0.5')."""
    if isinstance(v, float):
        if v == int(v):
            return str(int(v))
        return format(v, '.15g')
    if isinstance(v, bool):
        return 'TRUE' if v else 'FALSE'
    return str(v)


def validate_table(tbl, rep):
    """Row-level type / range checks driven by the row-3 datatype column."""
    seen_ids = {}
    n = len(tbl.rows)
    for idx, row in enumerate(tbl.rows):
        exrow = row.get('_row', '?')
        for col in tbl.columns:
            var = col.var
            if not var:
                continue
            v = row.get(var)
            if _is_empty(v):
                continue
            kind, arg = col.kind, col.arg

            # -- primary keys: only the 'id'-typed column is unique --
            if col.kind == 'id':
                if _is_empty(v):
                    rep.add(tbl.name, exrow, var, 'key',
                            'missing primary key value')
                    continue
                if not _int_coerce(v):
                    rep.add(tbl.name, exrow, var, 'key',
                            'non-numeric key value %r' % _celltext(v))
                key = _key_of(v)
                if key is not None and key in seen_ids:
                    rep.add(tbl.name, exrow, var, 'duplicate',
                            'duplicate key %s (first at excel row %s)'
                            % (key, seen_ids[key]))
                elif key is not None:
                    seen_ids[key] = exrow

            # -- datatype rules --
            err = _type_error(v, kind, arg)
            if err:
                cat = 'type'
                if kind in ('id', 'long', 'nmb'):
                    # numbers embedded in text labels (e.g. '23 (2017-...') are
                    # tolerated by the tool (it reads the leading digits).
                    if re.match(r'^\s*-?\d', _excel_str(v)):
                        cat = 'warn'
                rep.add(tbl.name, exrow, var, cat,
                        '%s expected [row-3=%s] got %r' % (
                            err, _celltext(col.caption or kind),
                            _celltext(v)))
        # progress hint for the long multibeam sheets
        if (idx + 1) % 50000 == 0:
            print('   ... %s row %d/%d' % (tbl.name, idx + 1, n))

    # coordinate boxes: any lat/lon group xxx1..xxx4
    validate_boxes(tbl, rep)


def _int_coerce(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return False
    return f == int(f)


def _key_of(v):
    try:
        f = float(v)
        if f == int(f):
            return int(f)
    except (TypeError, ValueError):
        pass
    if isinstance(v, str):
        return v.strip()
    return str(v)


def _type_error(v, kind, arg):
    """Return a short description if v violates the datatype, else None."""
    if kind in ('id', 'long', 'nmb'):
        try:
            float(v)
        except (TypeError, ValueError):
            return 'numeric'
        # 'id' must be an integral excel number; 'long'/'nmb' accept any number
        # (grid_dx etc. legitimately store decimal intervals)
        if kind == 'id' and not _int_coerce(v):
            return 'integer key (%s)' % _celltext(v)
        return None
    if kind == 'date':
        try:
            serial = float(v)
        except (TypeError, ValueError):
            return 'Excel serial date'
        if serial == 0:
            return None                     # empty-ish placeholder
        if serial < 15000 or serial > 90000:
            return 'plausible Excel serial date'
        return None
    if kind == 'year':
        try:
            y = int(float(v))
        except (TypeError, ValueError):
            return '4-digit year'
        if not (1900 <= y <= 2200):
            return 'year in 1900-2200'
        return None
    if kind in ('varchar2', 'char', 'list'):
        s = _excel_str(v)
        if arg is None:
            return None
        if kind == 'list':
            # expected syntax :tok1:tok2:tok3:  (colon-delimited tokens)
            if len(s) > arg and arg <= 512:
                return 'max length %d' % arg
            if not (s.startswith(':') and s.endswith(':')):
                return 'list syntax :a:b: (missing leading/trailing colon)'
            inner = s[1:-1]
            if inner and (inner.startswith(':') or inner.endswith(':')
                          or '::' in inner):
                return 'list syntax :a:b: (empty token)'
            return None
        if len(s) > arg:
            return 'max length %d' % arg
        return None
    return None


def validate_boxes(tbl, rep):
    """Faithful port of My_Calculate_Lat_Lon_Areas (SC.console).

    Computed only for the tables the tool run on: SC_seamounts / SC_regions
    (lat1,lon1,lat2,lon2) and MGG_multibeam* (lat3,lon3,lat4,lon4), with

        area = abs((LAT2-LAT1) * (LON2-LON1))

    Rows whose creation_status records "Error: Zero area ..." are ignored
    (they are known); we flag only the *gaps*:
      * area == 0 but no "Error: Zero area" marker recorded, or
      * marker recorded but the computed area is non-zero (drift).
    lat/lon range sanity is applied to any box family in any table.
    """
    name = tbl.name
    if name == 'SC_seamounts' or name == 'SC_regions':
        family = ('lat1', 'lon1', 'lat2', 'lon2')
        compute = True
    elif name.startswith('MGG_multibeam'):
        family = ('lat3', 'lon3', 'lat4', 'lon4')
        compute = True
    else:
        family = ()
        compute = False

    # range sanity on all lat/lon corners
    for v in list(tbl.colmap):
        if re.match(r'^lat[0-4]$', v):
            for row in tbl.rows:
                raw = row.get(v)
                if _is_empty(raw):
                    continue
                try:
                    x = float(raw)
                except (TypeError, ValueError):
                    continue
                if abs(x) > 90:
                    rep.add(name, row.get('_row', '?'), v, 'lat',
                            '|lat|>90: %g' % x)
        elif re.match(r'^lon[0-4]$', v):
            for row in tbl.rows:
                raw = row.get(v)
                if _is_empty(raw):
                    continue
                try:
                    x = float(raw)
                except (TypeError, ValueError):
                    continue
                if not 0 <= x <= 360:
                    rep.add(name, row.get('_row', '?'), v, 'lon',
                            'lon outside 0..360: %g' % x)

    if not compute or not all(x in tbl.colmap for x in family):
        return
    la1v, lo1v, la2v, lo2v = family
    for row in tbl.rows:
        exrow = row.get('_row', '?')
        if any(_is_empty(row.get(k)) for k in family):
            continue
        try:
            la1 = float(row.get(la1v)); lo1 = float(row.get(lo1v))
            la2 = float(row.get(la2v)); lo2 = float(row.get(lo2v))
        except (TypeError, ValueError):
            continue
        area = abs((la2 - la1) * (lo2 - lo1))
        status = str(row.get('creation_status') or '')
        is_rec_zero = status.startswith('Error: Zero area')
        if area == 0 and not is_rec_zero:
            rep.add(name, exrow, la1v, 'area',
                    'zero-area box (%g,%g -> %g,%g); no error marker'
                    % (la1, lo1, la2, lo2))
        elif area != 0 and is_rec_zero:
            rep.add(name, exrow, 'creation_status', 'status',
                    'recorded "Error: Zero area" but computed area = %g' % area)


# ---------------------------------------------------------------------------
# cross-table integrity (the hierarchy)
# ---------------------------------------------------------------------------

def _col_set(table, var):
    s = set()
    for row in table.rows:
        v = row.get(var)
        if not _is_empty(v):
            s.add(_key_of(v))
    return s


def cross_table_check(tables, rep):
    """Referential integrity between the database tables."""
    def ref(table, var):
        return _col_set(table, var)

    def exists(s):
        return s is not None

    checks = [
        ('SC_maps', 'seamount_index', 'SC_seamounts', 'seamount_index', 'seamount'),
        ('SC_maps', 'region_index', 'SC_regions', 'region_index', 'region'),
        ('MGG_samples', 'er_expedition_name', 'MGG_expeditions',
         'er_expedition_name', 'expedition'),
    ]
    refset = {}
    for src, sv, dst, dv, label in checks:
        if src not in tables or dst not in tables:
            continue
        ss = refset.setdefault(dst + '.' + dv, ref(tables[dst], dv))
        for row in tables[src].rows:
            v = row.get(sv)
            if _is_empty(v):
                continue
            key = _key_of(v)
            if key not in ss:
                rep.add(src, row.get('_row', '?'), sv, 'ref',
                        '%s "%s" not found in %s.%s' % (label, key, dst, dv))


def cross_file_loggers(tables, rep):
    """:n: tokens in every pacer_log_ids field must name a PACER_logger."""
    if 'PACER_loggers' not in tables:
        return
    lg_ids = _col_set(tables['PACER_loggers'], 'pacer_log_id')
    for tname, tbl in tables.items():
        for col in tbl.columns:
            if col.var != 'pacer_log_ids':
                continue
            for row in tbl.rows:
                v = row.get('pacer_log_ids')
                if _is_empty(v):
                    continue
                toks = [t for t in str(v).split(':') if t != '']
                for t in toks:
                    try:
                        key = int(float(t))
                    except (TypeError, ValueError):
                        continue
                    if key not in lg_ids:
                        rep.add(tname, row.get('_row', '?'), 'pacer_log_ids',
                                'ref', 'logger id :%d: not in PACER_loggers' % key)


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main(argv=None):
    ap = argparse.ArgumentParser(description='Sea-Cat data validation')
    ap.add_argument('paths', nargs='*', help='.xls files (default: scdpp, log)')
    ap.add_argument('--report', '-o', help='write report to file')
    ap.add_argument('--limit', type=int, default=0,
                    help='cap rows/table for a quick pass')
    ap.add_argument('--quiet', '-q', action='store_true')
    args = ap.parse_args(argv)

    paths = args.paths or [_DEFAULT_SRC + '\\' + f for f in _DEFAULT_FILES]
    t0 = time.time()
    tables = {}
    for p in paths:
        if not args.quiet:
            print('Reading %s ...' % p, flush=True)
        try:
            t = load_workbook(p, limit=args.limit)
        except Exception as e:
            print('!! could not read %s: %s' % (p, e))
            continue
        tables.update(t)

    # inventory
    print('\n=== sheet inventory ===')
    for name in sorted(tables):
        tbl = tables[name]
        print('  %-18s rows=%6d cols=%3d  (fields: %s)'
              % (name, tbl.n_data(), len(tbl.columns),
                 ', '.join(c.var for c in tbl.columns[:6])
                 + (' ...' if len(tbl.columns) > 6 else '')))

    rep = Report()

    print('\n=== validating rows ===', flush=True)
    for name in sorted(tables):
        validate_table(tables[name], rep)

    print('\n=== box / area checks ===', flush=True)
    for name in sorted(tables):
        validate_boxes(tables[name], rep)

    print('\n=== cross-table references ===', flush=True)
    cross_table_check(tables, rep)
    cross_file_loggers(tables, rep)

    # report
    counts = rep.counts()
    tcounts = rep.table_counts()
    lines = []
    lines.append('Data validation report')
    lines.append('=' * 40)
    for p in paths:
        lines.append('input : %s' % p)
    lines.append('')
    for name in sorted(tables):
        lines.append('%-16s %6d records, %3d fields'
                     % (name, tables[name].n_data(), len(tables[name].columns)))
    lines.append('')
    lines.append('Issue summary (by category):')
    for cat in sorted(counts):
        lines.append('  %-12s %d' % (cat, counts[cat]))
    lines.append('')
    lines.append('Issue summary (by sheet):')
    for t in sorted(tcounts):
        lines.append('  %-18s %d' % (t, tcounts[t]))
    lines.append('')
    from collections import Counter
    groups = Counter((cat, tbl, var) for tbl, _r, var, cat, _m in rep.issues)
    lines.append('Top issue groups (category / sheet / field):')
    for (cat, tbl, var), n in groups.most_common(20):
        lines.append('  %-10s %-16s %-16s %d' % (cat, tbl, var, n))
    lines.append('')

    lines.append('First %d issues (of %d):' % (min(60, len(rep.issues)),
                                               len(rep.issues)))
    for tbl, row, var, cat, msg in rep.issues[:60]:
        lines.append('  %-16s excel-row %-6s %-16s %-9s %s'
                     % (tbl, row, var, cat, msg))
    if len(rep.issues) > 60:
        lines.append('  ... and %d more' % (len(rep.issues) - 60))

    lines.append('')
    lines.append('Elapsed: %.1f s' % (time.time() - t0))
    body = '\n'.join(lines)
    if args.report:
        with open(args.report, 'w', encoding='utf-8') as fh:
            fh.write(body)
    if not args.quiet:
        print(body)

    sys.exit(1 if rep.issues else 0)


if __name__ == '__main__':
    main()