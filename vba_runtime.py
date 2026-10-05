# -*- coding: utf-8 -*-
"""vba2py runtime -- a small, best-effort stand-in for VBA built-ins and the
Excel object model used by code emitted by vba2py.

Nothing here reproduces all of VBA; it exists so generated ports can be
imported and exercised.  spreadsheets are kept in memory; if openpyxl is
installed, vba.open_workbook()/save() bridge to real .xlsx files.
"""

import datetime
import glob
import math
import os
import re
import random
import shutil
import subprocess
import sys
import time
import warnings

try:
    import openpyxl
except ImportError:
    openpyxl = None


# --------------------------------------------------------------------------
# constants (also exposed via module __getattr__ fallback)
# --------------------------------------------------------------------------

CONST = {
    'xlDisabled': -4146, 'xlWait': -4135, 'xlDefault': -4143,
    'xlVeryHidden': 2, 'xlNormal': -4143, 'xlContinuous': 1,
    'xlEdgeTop': 8, 'xlEdgeBottom': 9, 'xlEdgeLeft': 7, 'xlEdgeRight': 10,
    'xlInsideHorizontal': 12, 'xlInsideVertical': 11,
    'xlLeft': -4131, 'xlCenter': -4108, 'xlRight': -4152, 'xlJustify': -4130,
    'xlDown': -4121, 'xlUp': -4162, 'xlToLeft': -4159, 'xlToRight': -4161,
    'xlNone': -4142, 'xlCellTypeLastCell': 11, 'xlScreen': 1, 'xlPrinter': 2,
    'xlSortAscending': 1, 'xlSortDescending': 2, 'xlNo': 2, 'xlYes': 1,
    'xlAuto': -4105, 'xlEnglish': 3, 'xlSum': -4157, 'xlCount': -4112,
    'xlAverage': -4106, 'xlMax': -4136, 'xlMin': -4139, 'xlProduct': -4149,
    'xlCountNums': -4112, 'xlStDev': -4155, 'xlVar': -4164, 'xlNumbers': 2,
    'xlText': -4168, 'xlCalculationAutomatic': -4105,
    'xlCalculationManual': -4135,
    'xlDouble': -4119, 'xlThin': 2, 'xlHairline': 1, 'xlDiagonalDown': 5,
    'xlDiagonalUp': 6, 'xlFill': 1, 'xlSemiGray75': 10,
    'vbBlack': 0x0, 'vbRed': 0xFF, 'vbGreen': 0xFF00, 'vbYellow': 0xFFFF,
    'vbBlue': 0xFF0000, 'vbMagenta': 0xFF00FF, 'vbCyan': 0xFFFF00,
    'vbWhite': 0xFFFFFF, 'vbOKOnly': 0, 'vbOKCancel': 1, 'vbAbortRetryIgnore': 2,
    'vbYesNoCancel': 3, 'vbYesNo': 4, 'vbRetryCancel': 5,
    'vbCritical': 16, 'vbQuestion': 32, 'vbExclamation': 48,
    'vbInformation': 64, 'vbDefaultButton1': 0, 'vbDefaultButton2': 256,
    'vbDefaultButton3': 512, 'vbYes': 6, 'vbNo': 7, 'vbCancel': 2,
    'vbOK': 1, 'vbVariant': 12, 'vbString': 8, 'vbLong': 3, 'vbInteger': 2,
    'vbSingle': 4, 'vbDouble': 5, 'vbBoolean': 11, 'vbObject': 9,
    'vbApplication': 1, 'vbIgnore': 5, 'vbNoFocus': 3,
    'vbBack': '\b', 'vbCr': '\r', 'vbCrLf': '\r\n', 'vbDoubleQuote': '"',
    'vbFormFeed': '\x0c', 'vbLf': '\n', 'vbNewLine': '\n',
    'vbNullChar': '\x00', 'vbNullString': '', 'vbPipe': '|', 'vbSpace': ' ',
    'vbTab': '\t', 'vbVerticalTab': '\x0b', 'vbColon': ':',
}


def __getattr__(name):
    if name in CONST:
        return CONST[name]
    raise AttributeError(name)


# --------------------------------------------------------------------------
# miscellaneous VBA built-ins
# --------------------------------------------------------------------------

def array(*items):
    return list(items)


def iif(cond, a, b):
    return a if cond else b


def between(x, lo, hi):
    return lo <= x <= hi


def choose(idx, *items):
    try:
        return items[int(idx) - 1]
    except (IndexError, TypeError, ValueError):
        return None


def switch_pairs(*pairs):
    for i in range(0, len(pairs) - 1, 2):
        if pairs[i]:
            return pairs[i + 1]
    return None


def switch(*pairs):
    return switch_pairs(*pairs)


def UCase(s):
    return str(s).upper()


def LCase(s):
    return str(s).lower()


def trimv(s):
    if s is None:
        return ''
    return str(s).strip()


def ltrim(s):
    return str(s).lstrip()


def rtrim(s):
    return str(s).rstrip()


def left(s, n):
    return str(s)[:int(n)]


def leftb(s, n):
    return str(s)[:int(n)]


def right(s, n):
    t = str(s)
    return t[-int(n):]


def rightb(s, n):
    return right(s, n)


def mid(s, start, length=None):
    t = str(s)
    i = int(start) - 1
    if length is None:
        return t[i:]
    return t[i:i + int(length)]


def midb(s, start, length=None):
    return mid(s, start, length)


def length(s):
    return len(str(s))


def lenb(s):
    return len(str(s).encode('utf-8', 'ignore'))


def asci(s):
    return ord(str(s)[0])


def ascb(s):
    return ord(str(s)[0])


def chrb(n):
    return chr(int(n))[0]


def instr(start_s, s2=None, *rest):
    """VBA InStr([start,] s1, s2)."""
    if s2 is None:
        return -1
    if isinstance(start_s, int):
        start, s1 = start_s, s2
    else:
        start, s1 = 1, start_s
    idx = str(s1).find(str(s2), int(start) - 1)
    return idx + 1 if idx >= 0 else 0


def instr_rev(s1, s2):
    return str(s1).rfind(str(s2)) + 1


def replace(s, find, repl, start=1, count=-1, compare=-1):
    t = str(s)
    if start > 1:
        t = t[start - 1:]
    t = t.replace(str(find), str(repl))
    return t


def str_comp(s1, s2, compare=0):
    a, b = str(s1), str(s2)
    if compare == 1:
        a, b = a.lower(), b.lower()
    return (a > b) - (a < b)


def str_conv(s, conversion, localeid=0):
    t = str(s)
    if conversion == 1:      # vbUpperCase
        return t.upper()
    if conversion == 2:      # vbLowerCase
        return t.lower()
    if conversion == 3:      # vbProperCase
        return t.title()
    return t


def str_reverse(s):
    return str(s)[::-1]


def strx(n, ch):
    ch = str(ch)
    return ch * int(n) if ch else ''


def split(s, delim=' ', limit=-1, compare=-1):
    t = str(s)
    if delim == '' or delim is None:
        return list(t)
    parts = t.split(delim)
    while parts and parts[-1] == '':
        parts.pop()
    return parts


def join(items, delim=' '):
    return delim.join(str(x) for x in items)


def space(n):
    return ' ' * int(n)


def spc(n):
    return space(n)


def _tab(n):
    return space(n)


def tab(n):
    return _tab(n)


def ubound(arr, dim=1):
    try:
        return len(arr) - 1
    except TypeError:
        return 0


def lbound(arr, dim=1):
    return 0


def fix(x):
    try:
        return int(x)
    except (TypeError, ValueError):
        try:
            return math.trunc(float(x))
        except Exception:
            return x


def roundv(x, ndigits=0):
    return round(float(x), int(ndigits))


def sgn(x):
    if x > 0:
        return 1
    if x < 0:
        return -1
    return 0


def sqr(x):
    return math.sqrt(x)


def rnd():
    return random.random()


def randomize(seed=None):
    if seed is not None:
        random.seed(seed)


def val(s):
    m = re.match(r'\s*[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?', str(s or ''))
    if not m:
        return 0
    return float(m.group(0))


def is_numeric(v):
    try:
        float(v)
        return True
    except (TypeError, ValueError):
        return False


def is_empty(v):
    return v is None


def is_missing(v):
    return v is None


def is_null(v):
    return v is None


def is_date(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return True
    return is_numeric(v)


def is_array(v):
    return isinstance(v, (list, tuple))


def is_object(v):
    return hasattr(v, '__dict__')


def is_error(v):
    return isinstance(v, (SyntaxError, Exception)) and not isinstance(v, (TypeError, ValueError))


def type_name(v):
    return type(v).__name__


def cstr(v):
    if v is None:
        return ''
    if isinstance(v, bool):
        return 'True' if v else 'False'
    return str(v)


def cvar(v):
    return v


def cbyte(v):
    return int(v) & 0xFF


def ccur(v):
    return round(float(v), 4)


def cdec(v):
    return round(float(v), 28)


def var_type(v):
    types = {bool: 11, int: 3, float: 5, str: 8, list: 0x2000 + 3, tuple: 0x2000 + 3}
    return types.get(type(v), 0)


def hexv(n, length=None):
    s = format(int(n), 'X')
    return s if length is None else s.rjust(int(length), '0')


def octv(n, length=None):
    s = format(int(n), 'o')
    return s if length is None else s.rjust(int(length), '0')


def rgb(r, g, b):
    return (int(r) & 0xFF) * 0x10000 + (int(g) & 0xFF) * 0x100 + (int(b) & 0xFF)


def qbcolor(n):
    cols = [0, 0x800000, 0x8000, 0x808000, 0x80, 0x800080, 0x8080, 0xC0C0C0,
            0x808080, 0xFF0000, 0xFF00, 0xFFFF00, 0xFF, 0xFF00FF, 0xFFFF, 0xFFFFFF]
    return cols[int(n) % len(cols)]


def varray(v):
    return list(v)


# --------------------------------------------------------------------------
# date / time
# --------------------------------------------------------------------------

def Now():
    return datetime.datetime.now()


def Date():
    return datetime.date.today()


def Time():
    return datetime.datetime.now().time()


def Timer():
    return time.time() - _TIMER0


_TIMER0 = time.time()


def Day(d):
    return d.day


def Month(d):
    return d.month


def Year(d):
    return d.year


def Hour(t):
    return t.hour


def Minute(t):
    return t.minute


def Second(t):
    return t.second


def Weekday(d, firstdayofweek=0):
    iso = d.isoweekday()
    return iso + 1 if iso != 7 else 1


def weekday_name(n, abbreviate=False, firstdayofweek=0):
    names = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
    return names[int(n) - 1][:3] if abbreviate else names[int(n) - 1]


def month_name(n, abbreviate=False):
    names = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
             'August', 'September', 'October', 'November', 'December']
    return names[int(n) - 1][:3] if abbreviate else names[int(n) - 1]


def date_value(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    return datetime.datetime.fromisoformat(str(v)).date()


def time_value(v):
    if isinstance(v, datetime.time):
        return v
    if isinstance(v, datetime.datetime):
        return v.time()
    return datetime.time.fromisoformat(str(v))


def date_serial(y, m, d):
    return datetime.date(int(y), int(m), int(d))


def time_serial(h, m, s):
    return datetime.time(int(h), int(m), int(s))


def date_part(interval, d):
    if interval in ('yyyy', 'yyyy'):
        return d.year
    if interval in ('m', 'month'):
        return d.month
    if interval in ('d', 'day'):
        return d.day
    if interval in ('ww', 'week'):
        return d.isocalendar()[1]
    if interval in ('q', 'quarter'):
        return (d.month - 1) // 3 + 1
    return 0


def date_add(interval, n, d):
    n = int(n)
    if interval == 'yyyy':
        return d.replace(year=d.year + n)
    if interval == 'm':
        y = (d.month - 1 + n) // 12
        m = (d.month - 1 + n) % 12 + 1
        return d.replace(year=d.year + y, month=m)
    if interval == 'd':
        return d + datetime.timedelta(days=n)
    if interval == 'ww':
        return d + datetime.timedelta(weeks=n)
    if interval == 'h':
        return d + datetime.timedelta(hours=n)
    return d


def date_diff(interval, d1, d2):
    pass  # provided for completeness


# --------------------------------------------------------------------------
# Format strings (the '0'/'#'/'.'/',' percent subset)
# --------------------------------------------------------------------------

def Format(x, fmt=''):
    if fmt is None:
        return cstr(x)
    fmt = str(fmt)
    if x is None or x == '':
        return ''
    if fmt.startswith('0') or fmt.startswith('#'):
        return _format_number(x, fmt)
    if fmt.startswith('@') or fmt.startswith('&'):
        return cstr(x)
    if fmt.startswith('0%') or fmt.startswith('#%'):
        return '%d%%' % round(float(x) * 100)
    return _format_number(x, fmt)


def _format_number(x, fmt):
    if isinstance(x, str):
        try:
            x = float(x)
        except ValueError:
            return x
    neg = x < 0
    s = re.sub(r'[^0#.,%-]', '', fmt)
    decimal_digits = len(s.split('.')[1]) if '.' in s else 0
    thousands = ',' in s
    f = abs(float(x))
    if '.%' or '%' in fmt:
        f = f * 100
    if decimal_digits:
        txt = ('%%.%df' % decimal_digits) % f
    else:
        txt = '%.0f' % f
    if thousands:
        ipart, _, fpart = txt.partition('.')
        ipart = _group(ipart)
        txt = ipart + ('.' + fpart if fpart else '')
    return ('-' if neg else '') + txt


def _group(digits):
    out = []
    for i, ch in enumerate(digits[::-1]):
        if i and i % 3 == 0:
            out.append(',')
        out.append(ch)
    return ''.join(reversed(out))


def format_number(x, numdigits=-1, leadingdigits=-1, useParens=False, groupdigits=-1):
    return Format(x, '#,##0' + (('.%d' % numdigits) if numdigits >= 0 else ''))


def format_percent(x, numdigits=-1, leadingdigits=-1, useParens=False, groupdigits=-1):
    return ('%.%df%%' % (numdigits if numdigits >= 0 else 0)) % (float(x) * 100)


def format_currency(x, numdigits=-1, leadingdigits=-1):
    return '$' + Format(x, '#,##0' + (('.%d' % numdigits) if numdigits >= 0 else ''))


def format_datetime(x, namedformat=0):
    return x.strftime('%m/%d/%Y %H:%M:%S')


def cdate(v):
    if isinstance(v, datetime.datetime):
        return v
    if isinstance(v, datetime.date):
        return datetime.datetime(v.year, v.month, v.day)
    if isinstance(v, (int, float)):
        return datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(v))
    try:
        return datetime.datetime.fromisoformat(str(v))
    except ValueError:
        return str(v)


def command():
    return ''


def load_picture(path):
    return None


def partition(n, lo, hi, interval):
    return '%d:%d' % (lo, hi)


# --------------------------------------------------------------------------
# environment / system
# --------------------------------------------------------------------------

def Environ(name):
    key = str(name)
    if key.isdigit():
        vals = [k for k in os.environ if k]
        return vals[int(key) - 1] if 0 < int(key) <= len(vals) else ''
    return os.environ.get(key, '')


def Dir(path=None):
    if path is None:
        return ''
    if any(ch in path for ch in '*?'):
        matches = glob.glob(path)
        return os.path.basename(matches[0]) if matches else ''
    return os.path.basename(path) if os.path.exists(path) else ''


def curdir():
    return os.getcwd()


def chdir(path):
    os.chdir(path)
    return True


def mkdir(path):
    os.makedirs(str(path), exist_ok=True)
    return True


def rmdir(path):
    try:
        os.rmdir(str(path))
    except OSError:
        shutil.rmtree(str(path), ignore_errors=True)
    return True


def Kill(path):
    for p in glob.glob(str(path)):
        try:
            os.remove(p)
        except OSError:
            pass
    return True


def FileCopy(src, dst):
    shutil.copy(str(src), str(dst))
    return True


def Shell(cmd, windowstyle=1, wait=True):
    if wait:
        return subprocess.call(str(cmd), shell=True)
    return subprocess.Popen(str(cmd), shell=True).pid


def DoEvents():
    return 0


def file_len(path):
    return os.path.getsize(str(path))


def file_datetime(path):
    return datetime.datetime.fromtimestamp(os.path.getmtime(str(path)))


def get_attr(path):
    return 0


def beep():
    try:
        import ctypes
        ctypes.windll.kernel32.Beep(800, 200)
    except Exception:
        pass


def send_keys(*args):
    pass


def CreateObject(cls):
    warnings.warn('CreateObject(%r) stub' % (cls,))
    return Dummy(cls)


def GetObject(*args):
    warnings.warn('GetObject stub')
    return Dummy('object')


def call_by_name(obj, name, calltype, *args):
    fn = getattr(obj, name)
    return fn(*args)


def like(s, pattern):
    rx = []
    i = 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == '*':
            rx.append('.*')
        elif ch == '?':
            rx.append('.')
        elif ch == '[':
            j = pattern.find(']', i + 1)
            if j < 0:
                rx.append(re.escape(ch))
            else:
                rx.append(pattern[i:j + 1])
                i = j
        else:
            rx.append(re.escape(ch))
        i += 1
    return re.match('^' + ''.join(rx) + '$', str(s), re.IGNORECASE) is not None


# --------------------------------------------------------------------------
# message boxes / input
# --------------------------------------------------------------------------

def MsgBox(prompt='', buttons=0, title='', helpfile=None, context=None):
    sys.stdout.write(str(prompt) + '\n')
    return 1  # vbOK


def InputBox(prompt='', title='', default='', xpos=None, ypos=None):
    sys.stdout.write(str(prompt) + ' ')
    sys.stdout.flush()
    line = sys.stdin.readline().rstrip('\n')
    return line if line != '' else str(default)


# --------------------------------------------------------------------------
# file numbers / old-style file I/O
# --------------------------------------------------------------------------

_FILES = {}
_NEXT_FNUM = [100]


def free_file(rangestart=1, rangeend=255):
    n = _NEXT_FNUM[0]
    _NEXT_FNUM[0] += 1
    return n


def file_open(num, path, mode='r'):
    num = int(num)
    _FILES[num] = open(path, mode, encoding='utf-8', errors='ignore')
    return num


def file_print(num, *items):
    fh = _FILES[int(num)]
    fh.write(''.join(cstr(x) for x in items) + '\n')


def file_write(num, *items):
    fh = _FILES[int(num)]
    fh.write(','.join(cstr(x) for x in items) + '\n')


def file_line_input(num):
    fh = _FILES[int(num)]
    line = fh.readline()
    if line == '':
        raise EOFError('Line Input past end of file')
    return line.rstrip('\r\n')


def file_input(num):
    fh = _FILES[int(num)]
    line = fh.readline()
    if line == '':
        raise EOFError('Input past end of file')
    fields = [f.strip() for f in line.split(',')]
    out = []
    for f in fields:
        try:
            out.append(int(f))
        except ValueError:
            try:
                out.append(float(f))
            except ValueError:
                out.append(f.strip('"'))
    return out


def file_close(num):
    fh = _FILES.pop(int(num), None)
    if fh is not None:
        try:
            fh.close()
        except Exception:
            pass


def file_close_all():
    while _FILES:
        k, v = _FILES.popitem()
        try:
            v.close()
        except Exception:
            pass


# --------------------------------------------------------------------------
# Excel-like object model (in-memory)
# --------------------------------------------------------------------------

class Dummy:
    """Object that accepts any attribute get/set and any call."""

    def __init__(self, name='obj'):
        self.__dict__['name'] = name

    def __getattr__(self, item):
        return Dummy('%s.%s' % (self.name, item))

    def __setattr__(self, key, value):
        self.__dict__[key] = value

    def __call__(self, *a, **k):
        return self

    def __repr__(self):
        return '<VBA %s>' % self.name


def new_object(name):
    """VBA 'New X' fallback: return a known class, else a Dummy."""
    obj = globals().get(name)
    if obj is None:
        return Dummy()
    try:
        return obj()
    except TypeError:
        return Dummy()


def _col_to_letters(n):
    s = ''
    n = int(n)
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _letters_to_col(s):
    col = 0
    for ch in s:
        col = col * 26 + (ord(ch.upper()) - 64)
    return col


class _RangeColl:
    def __init__(self, ws, index, count):
        self._ws = ws
        self._index = index
        self._count = count

    @property
    def Count(self):
        return self._count

    @property
    def Rows(self):
        return self

    @property
    def Columns(self):
        return self

    def Item(self, i):
        return self(i)

    def __call__(self, i):
        return Range(self._ws, int(i), int(i))


class Range:
    def __init__(self, ws, r1, c1, r2=None, c2=None):
        self._ws = ws
        self.r1, self.c1 = int(r1), int(c1)
        self.r2 = int(r2) if r2 is not None else int(r1)
        self.c2 = int(c2) if c2 is not None else int(c1)

    @property
    def Value(self):
        if self.r1 == self.r2 and self.c1 == self.c2:
            return self._ws.get(self.r1, self.c1)
        rows = []
        for r in range(self.r1, self.r2 + 1):
            rows.append([self._ws.get(r, c) for c in range(self.c1, self.c2 + 1)])
        return rows

    @Value.setter
    def Value(self, v):
        if self.r1 == self.r2 and self.c1 == self.c2:
            self._ws.set(self.r1, self.c1, v)
            return
        if isinstance(v, (list, tuple)):
            r = self.r1
            for row in v:
                c = self.c1
                if isinstance(row, (list, tuple)):
                    for item in row:
                        self._ws.set(r, c, item)
                        c += 1
                else:
                    self._ws.set(r, c, row)
                r += 1
        else:
            for r in range(self.r1, self.r2 + 1):
                for c in range(self.c1, self.c2 + 1):
                    self._ws.set(r, c, v)

    @property
    def Formula(self):
        return self.Value

    @Formula.setter
    def Formula(self, v):
        self.Value = v

    @property
    def Row(self):
        return self.r1

    @property
    def Column(self):
        return self.c1

    @property
    def Address(self):
        return '$%s$%d' % (_col_to_letters(self.c1), self.r1)

    @property
    def EntireRow(self):
        return Range(self._ws, self.r1, 1, self.r1, 16384)

    @property
    def EntireColumn(self):
        return Range(self._ws, 1, self.c1, 1048576, self.c1)

    @property
    def Rows(self):
        return _RangeColl(self._ws, 1, self.r2 - self.r1 + 1)

    @property
    def Columns(self):
        return _RangeColl(self._ws, 2, self.c2 - self.c1 + 1)

    @property
    def Width(self):
        return 8.43

    def Offset(self, dr, dc=0):
        return Range(self._ws, self.r1 + int(dr), self.c1 + int(dc))

    def Resize(self, r, c=None):
        c = c if c is not None else (self.c2 - self.c1 + 1)
        return Range(self._ws, self.r1, self.c1, self.r1 + int(r) - 1, self.c1 + int(c) - 1)

    def End(self, direction):
        if direction == -4162:  # xlUp
            r = self.r1
            while r > 1 and self._ws.get(r - 1, self.c1) is None:
                r -= 1
            r -= 1
            return Range(self._ws, r, self.c1)
        return self

    def Select(self):
        vba_sel = self
        return None

    def Activate(self):
        return None

    def Delete(self, shift=None):
        return None

    def Insert(self, shift=None, copyorigin=None):
        return None

    def Copy(self, destination=None):
        return None

    def Cut(self, destination=None):
        return None

    def Clear(self):
        return None

    def ClearContents(self):
        return None

    def __getitem__(self, key):
        if isinstance(key, tuple):
            r, c = key
            return Range(self._ws, self.r1 + int(r) - 1, self.c1 + int(c) - 1)
        return Range(self._ws, self.r1 + int(key) - 1, self.c1)

    def __repr__(self):
        return '<Range %s!%s>' % (self._ws.Name, self.Address)


class _CellsProxy:
    def __init__(self, ws):
        self._ws = ws

    @property
    def Count(self):
        return 1048576 * 16384

    @property
    def Rows(self):
        return _RangeColl(self._ws, 1, 1048576)

    @property
    def Columns(self):
        return _RangeColl(self._ws, 2, 16384)

    def Item(self, r, c=1):
        return Range(self._ws, r, c)

    def __call__(self, r, c=1):
        return Range(self._ws, r, c)

    def Select(self):
        return None

    def __repr__(self):
        return '<Cells of %s>' % self._ws.Name


class Worksheet:
    def __init__(self, wb, name, index, cells=None):
        self._wb = wb
        self.Name = name
        self.Index = index
        self._cells = cells if cells is not None else {}
        self.Visible = '{0}'
        self.ProtectContents = False
        self._cells_proxy = _CellsProxy(self)

    def get(self, r, c):
        return self._cells.get((int(r), int(c)))

    def set(self, r, c, v):
        self._cells[(int(r), int(c))] = v

    @property
    def Cells(self):
        return self._cells_proxy

    def Cells_(self, r, c=1):
        return Range(self, r, c)

    def Range(self, ref):
        m = re.match(r'^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$', str(ref).strip())
        if not m:
            return Range(self, 1, 1)
        c1, r1 = _letters_to_col(m.group(1)), int(m.group(2))
        if m.group(3):
            return Range(self, r1, c1, int(m.group(4)), _letters_to_col(m.group(3)))
        return Range(self, r1, c1)

    def Names(self, name=None):
        return Dummy('Names')

    @property
    def Rows(self):
        return _RangeColl(self, 1, 1048576)

    @property
    def Columns(self):
        return _RangeColl(self, 2, 16384)

    @property
    def UsedRange(self):
        return Range(self, 1, 1)

    def Activate(self):
        global asheet
        asheet = self
        return None

    def Select(self, replace=None):
        global asheet
        asheet = self
        return None

    def Protect(self, *a, **k):
        return None

    def Unprotect(self, *a, **k):
        return None

    def __repr__(self):
        return '<Worksheet %s>' % self.Name


class Workbook:
    def __init__(self, name='BOOK1'):
        self.Name = name
        self._sheets = {}
        self.SaveAs = None
        self.FullName = name

    def _get_or_create_sheet(self, key):
        if isinstance(key, int):
            idx = key
            if idx not in self._sheets:
                self._sheets[idx] = Worksheet(self, 'Sheet%d' % idx, idx)
            return self._sheets[idx]
        key = str(key)
        for k, v in self._sheets.items():
            if isinstance(k, str) and k.lower() == key.lower():
                return v
        idx = len(self._sheets) + 1
        ws = Worksheet(self, key, idx)
        self._sheets[key] = ws
        return ws

    def Sheets(self, key=None):
        if key is None:
            return [v for k, v in sorted(self._sheets.items(), key=lambda kv: str(kv[0]))]
        return self._get_or_create_sheet(key)

    def Worksheets(self, key=None):
        return self.Sheets(key)

    def Charts(self, key=None):
        return self.Sheets(key)

    def Activate(self):
        return None

    def Save(self):
        return None

    def Close(self, savechanges=None):
        return None

    def __repr__(self):
        return '<Workbook %s>' % self.Name


def workbooks(name=None):
    name = str(name)
    if name in _WBS:
        return _WBS[name]
    wb = Workbook(name)
    _WBS[name] = wb
    return wb


_WBS = {'': Workbook('')}


class Application:
    def __init__(self):
        self._d = {}

    def __getattr__(self, item):
        return self._d.get(item, Dummy('app.' + item))

    def __setattr__(self, key, value):
        if key.startswith('_'):
            super().__setattr__(key, value)
        else:
            self._d[key] = value

    def Quit(self):
        return None


class WF:
    def __init__(self):
        pass

    def CountA(self, rng):
        if not isinstance(rng, Range):
            return 0
        total = 0
        for r in range(rng.r1, rng.r2 + 1):
            for c in range(rng.c1, rng.c2 + 1):
                v = rng._ws.get(r, c)
                if v not in (None, ''):
                    total += 1
        return total

    def Count(self, rng):
        if not isinstance(rng, Range):
            return 0
        vals = rng.Value
        flat = vals if isinstance(vals, list) else [vals]
        return sum(1 for row in (flat if isinstance(flat[0], list) else [flat]) for v in row if v is not None)

    def CountIf(self, rng, criteria):
        if not isinstance(rng, Range):
            return 0
        crit = str(criteria)
        total = 0
        vals = rng.Value
        if not isinstance(vals, list):
            vals = [[vals]]
        for row in vals:
            for v in row:
                if crit.startswith('>'):
                    try:
                        if float(v) > float(crit[1:]):
                            total += 1
                    except (TypeError, ValueError):
                        continue
                elif str(v) == crit:
                    total += 1
        return total

    def Sum(self, rng):
        return self._sum(rng)

    def SumIf(self, rng, criteria, sumrange=None):
        return self.Sum(rng)

    def Average(self, rng):
        vals = self._values(rng)
        vals = [v for v in vals if v is not None]
        return sum(vals) / len(vals) if vals else 0

    def Max(self, rng):
        vals = self._values(rng)
        return max(vals) if vals else 0

    def Min(self, rng):
        vals = self._values(rng)
        return min(vals) if vals else 0

    def Median(self, rng):
        vals = sorted(self._values(rng))
        if not vals:
            return 0
        n = len(vals)
        mid = n // 2
        return vals[mid] if n % 2 else (vals[mid - 1] + vals[mid]) / 2

    def StDev(self, rng):
        vals = self._values(rng)
        if len(vals) < 2:
            return 0
        mean = sum(vals) / len(vals)
        return math.sqrt(sum((v - mean) ** 2 for v in vals) / (len(vals) - 1))

    def Var(self, rng):
        vals = self._values(rng)
        if len(vals) < 2:
            return 0
        mean = sum(vals) / len(vals)
        return sum((v - mean) ** 2 for v in vals) / (len(vals) - 1)

    def RoundUp(self, x, digits=0):
        f = 10 ** int(digits)
        return math.ceil(float(x) * f) / f

    def RoundDown(self, x, digits=0):
        f = 10 ** int(digits)
        return math.floor(float(x) * f) / f

    def Round(self, x, digits=0):
        return round(float(x), int(digits))

    def Abs(self, x):
        return abs(x)

    def Mod(self, a, b):
        return a % b

    def Trim(self, s):
        return re.sub(r'\s+', ' ', str(s)).strip()

    def Transpose(self, data):
        if isinstance(data, (list, tuple)):
            if data and isinstance(data[0], (list, tuple)):
                return [[row[i] for row in data] for i in range(len(data[0]))]
            return [[v] for v in data]
        return data

    def Match(self, value, rng, matchtype=0):
        vals = rng.Value if isinstance(rng, Range) else rng
        for i, v in enumerate(vals, 1):
            if v == value:
                return i
        return -1

    def Index(self, arr, rownum, colnum=None):
        if isinstance(arr, Range):
            arr = arr.Value
        if colnum is None:
            return arr[int(rownum) - 1]
        return arr[int(rownum) - 1][int(colnum) - 1]

    def VLookup(self, value, table, colindex, rangelookup=True):
        rows = table.Value if isinstance(table, Range) else table
        if not isinstance(rows, list):
            rows = [rows]
        for row in rows:
            if not isinstance(row, (list, tuple)):
                continue
            if row and row[0] == value:
                try:
                    return row[int(colindex) - 1]
                except IndexError:
                    return None
        return None

    def IfError(self, v, viferr):
        return viferr if is_error(v) or v is None else v

    def SumProduct(self, *args):
        rows = args[0].Value
        s = 0
        for row in rows:
            s += sum(0 if v is None else v for v in row)
        return s

    def Rand(self):
        return random.random()

    def _values(self, rng):
        vals = rng.Value if isinstance(rng, Range) else rng
        if not isinstance(vals, list):
            return [vals] if vals is not None else []
        flat = []
        for row in vals:
            if isinstance(row, (list, tuple)):
                flat.extend(v for v in row if v is not None)
            elif row is not None:
                flat.append(row)
        return flat

    def _sum(self, rng):
        return sum(self._values(rng))

    def __getattr__(self, item):
        def _missing(*a, **k):
            raise NotImplementedError('WorksheetFunction.%s not implemented in the '
                                      'vba2py runtime shim' % item)
        return _missing


class Form(Dummy):
    def Show(self, *a, **k):
        return None

    def Hide(self):
        return None

    def Unload(self):
        return None

    def Load(self):
        return None


_FORMS = {}


def form(n):
    name = 'UserForm%d' % int(n) if str(n).isdigit() else str(n)
    if name not in _FORMS:
        _FORMS[name] = Form(name)
    return _FORMS[name]


def form_unload(n):
    name = getattr(n, 'name', None)
    if name is None:
        name = str(n)
    _FORMS.pop(name, None)
    return None


def dummy(name):
    return Dummy(name)


# --------------------------------------------------------------------------
# runtime state for generated code
# --------------------------------------------------------------------------

wb = _WBS['']
awb = wb
default_sheet = wb._get_or_create_sheet(1)
asheet = default_sheet
acell = Range(asheet, 1, 1)
sel = acell
awin = Dummy('ActiveWindow')
app = Application()
wf = WF()


def sheet(name_or_idx):
    return wb._get_or_create_sheet(name_or_idx)


def active_sheet():
    return asheet


def open_workbook(path):
    wb = Workbook(os.path.basename(path))
    if openpyxl is not None:
        xl = openpyxl.load_workbook(path, data_only=True)
        for idx, ws in enumerate(xl.worksheets, 1):
            sheet = Worksheet(wb, ws.title, idx)
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        sheet.set(cell.row, cell.column, cell.value)
            wb._sheets[ws.title] = sheet
    _WBS[wb.Name] = wb
    return wb