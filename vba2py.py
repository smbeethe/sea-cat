#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Convert Excel VBA macros (from legacy .xls/.xlsm workbooks) to Python.

Reads the VBA source of every module in the given 97-2003 Excel files,
translates procedures and statements into Python, and writes:

  <workbook>.py               standard .bas modules merged into one flat namespace
                              (mirrors VBA's shared module namespace)
  <workbook>__<class>.py      each .cls / .frm module, one file
  <workbook>_report.txt       log of every heuristic / unsupported construct
  vba_runtime.py              helper runtime (VBA built-ins + Excel-like API)

Usage:
    python vba2py.py [PATH...] [-o OUTDIR] [--no-runtime] [--quiet]

PATH may be a .xls/.xlsm file or a directory (scanned recursively).
Defaults to the current directory.

Translation is syntactic, best-effort porting assistance.  Review the
generated code and the report before use.
"""

import argparse
import os
import re
import sys
import warnings

from oletools.olevba import VBA_Parser

warnings.filterwarnings("ignore")

PY_KEYWORDS = frozenset("""
False None True and as assert async await break class continue def del elif else
except finally for from global if import in is lambda nonlocal not or pass raise
return try while with yield
""".split())

# ---------------------------------------------------------------------------
# VBA -> Python dictionaries
# ---------------------------------------------------------------------------

EXCEL_TOP = {
    'ActiveCell', 'ActiveChart', 'ActiveSheet', 'ActiveWindow', 'ActiveWorkbook',
    'Application', 'Cells', 'Charts', 'CommandBars', 'Dialogs', 'Forms',
    'Names', 'Range', 'Rows', 'Columns', 'Sheets', 'Selection', 'ThisWorkbook',
    'Workbooks', 'Worksheets',
}

EXCEL_CALL_NAMES = {
    'Areas', 'Borders', 'Cells', 'Charts', 'Columns', 'CommandBars', 'Controls',
    'Dialogs', 'End', 'EntireColumn', 'EntireRow', 'Find', 'Font', 'Fonts',
    'Forms', 'Items', 'Names', 'Offset', 'OpenText', 'PrintOut', 'Range',
    'Replace', 'Resize', 'Rows', 'Shapes', 'Sheets', 'Workbooks', 'Worksheets',
    'WorksheetFunction',
}

BUILTIN_CALLS = {
    'Abs': 'abs', 'Asc': 'ord', 'AscB': 'vba.ascb', 'Array': 'vba.array',
    'Beep': 'vba.beep', 'CBool': 'bool', 'CByte': 'vba.cbyte',
    'CCur': 'vba.ccur', 'CDate': 'vba.cdate', 'CDbl': 'float',
    'CDec': 'vba.cdec', 'CInt': 'int', 'CLng': 'int',
    'CSng': 'float', 'CStr': 'vba.cstr', 'CVar': 'vba.cvar',
    'Choose': 'vba.choose', 'Chr': 'chr', 'ChrB': 'vba.chrb', 'ChrW': 'chr',
    'Command': 'vba.command', 'CreateObject': 'vba.CreateObject',
    'CurDir': 'vba.curdir', 'Date': 'vba.Date', 'DateAdd': 'vba.date_add',
    'DateDiff': 'vba.date_diff', 'DatePart': 'vba.date_part',
    'DateSerial': 'vba.date_serial', 'DateValue': 'vba.date_value',
    'Day': 'vba.Day', 'Dir': 'vba.Dir', 'DoEvents': 'vba.DoEvents',
    'Environ': 'vba.Environ', 'FileDateTime': 'vba.file_datetime',
    'FileLen': 'vba.file_len', 'Fix': 'vba.fix', 'Format': 'vba.Format',
    'FormatCurrency': 'vba.format_currency',
    'FormatDateTime': 'vba.format_datetime', 'FormatNumber': 'vba.format_number',
    'FormatPercent': 'vba.format_percent', 'FreeFile': 'vba.free_file',
    'GetAttr': 'vba.get_attr', 'GetObject': 'vba.GetObject',
    'Hex': 'vba.hexv', 'Hour': 'vba.Hour', 'IIf': 'vba.iif',
    'InputBox': 'vba.InputBox', 'InStr': 'vba.instr', 'InStrRev': 'vba.instr_rev', 'LoadPicture': 'vba.load_picture',
    'Int': 'int', 'IsArray': 'vba.is_array', 'IsDate': 'vba.is_date',
    'IsEmpty': 'vba.is_empty', 'IsError': 'vba.is_error',
    'IsMissing': 'vba.is_missing', 'IsNull': 'vba.is_null',
    'IsNumeric': 'vba.is_numeric', 'IsObject': 'vba.is_object',
    'Join': 'vba.join', 'LBound': 'vba.lbound', 'LCase': 'vba.LCase',
    'LTrim': 'vba.ltrim', 'Left': 'vba.left', 'LeftB': 'vba.leftb',
    'Len': 'len', 'LenB': 'vba.lenb', 'Max': 'max', 'Min': 'min',
    'Mid': 'vba.mid', 'MidB': 'vba.midb', 'Minute': 'vba.Minute',
    'Month': 'vba.Month', 'MonthName': 'vba.month_name',
    'MsgBox': 'vba.MsgBox', 'Now': 'vba.Now', 'Oct': 'vba.octv',
    'Partition': 'vba.partition', 'QBColor': 'vba.qbcolor',
    'Replace': 'vba.replace', 'Right': 'vba.right', 'RightB': 'vba.rightb',
    'Rnd': 'vba.Rnd', 'Randomize': 'vba.randomize', 'RTrim': 'vba.rtrim', 'Round': 'vba.roundv',
    'Second': 'vba.Second', 'Sgn': 'vba.sgn', 'Shell': 'vba.Shell',
    'Space': 'vba.Space', 'Spc': 'vba.spc', 'Split': 'vba.split',
    'Sqr': 'vba.sqr', 'Str': 'vba.cstr', 'StrComp': 'vba.str_comp',
    'StrConv': 'vba.str_conv', 'StrReverse': 'vba.str_reverse',
    'String': 'vba.strx', 'Switch': 'vba.switch', 'Tab': 'vba.tab',
    'Time': 'vba.Time', 'TimeSerial': 'vba.time_serial',
    'TimeValue': 'vba.time_value', 'Timer': 'vba.Timer', 'Trim': 'vba.trimv',
    'TypeName': 'vba.type_name', 'UBound': 'vba.ubound', 'UCase': 'vba.UCase',
    'Val': 'vba.val', 'VarType': 'vba.var_type', 'Weekday': 'vba.Weekday',
    'WeekdayName': 'vba.weekday_name', 'Year': 'vba.Year',
    'CallByName': 'vba.call_by_name', 'RGB': 'vba.rgb',
}

VB_STRING_CONST = {
    'vbBack': "'\\b'", 'vbCr': "'\\r'", 'vbCrLf': "'\\r\\n'",
    'vbDoubleQuote': "'\\\"'", 'vbFormFeed': "'\\x0c'", 'vbLf': "'\\n'",
    'vbNewLine': "'\\n'", 'vbNullChar': "'\\x00'", 'vbNullString': "''",
    'vbPipe': "'|'", 'vbSpace': "' '", 'vbTab': "'\\t'",
    'vbVerticalTab': "'\\x0b'",
}

VB_CONST2 = ['vbBlack', 'vbBlue', 'vbCyan', 'vbDarkBlue', 'vbDefaultButton1',
             'vbDefaultButton2', 'vbDefaultButton3', 'vbExclamation', 'vbGreen',
             'vbInformation', 'vbMagenta', 'vbOKCancel', 'vbOKOnly',
             'vbQuestion', 'vbRed', 'vbRetryCancel', 'vbWhite', 'vbYellow',
             'vbYesNo', 'vbYesNoCancel', 'xlCalculationAutomatic',
             'xlCalculationManual', 'xlCellTypeLastCell', 'xlContinuous',
             'xlDefault', 'xlDisabled', 'xlEdgeBottom', 'xlEdgeLeft',
             'xlEdgeRight', 'xlEdgeTop', 'xlInsideHorizontal',
             'xlInsideVertical', 'xlNone', 'xlScreen', 'xlSortAscending',
             'xlSortDescending', 'xlToLeft', 'xlToRight', 'xlUp', 'xlVeryHidden',
             'xlWait', 'xlDown', 'xlCenter', 'xlLeft', 'xlRight', 'xlJustify',
             'xlFill', 'xlAuto', 'xlEnglish', 'xlNo', 'xlYes', 'xlSum',
             'xlCount', 'xlAverage', 'xlMax', 'xlMin', 'xlProduct',
             'xlCountNums', 'xlStDev', 'xlVar', 'xlNumbers', 'xlText']


def _fix_ident(name):
    return name + '_' if name in PY_KEYWORDS else name


class Notes(object):
    def __init__(self):
        self.items = []

    def add(self, kind, msg):
        self.items.append((kind, msg))


# ---------------------------------------------------------------------------
# lexical helpers
# ---------------------------------------------------------------------------

def split_physical(line):
    """Return (statements, trailing_comment); handles "" escapes and comments."""
    det = re.sub(r'^\d+(\.\d+)?\s*', '', line)
    whole_if = bool(re.match(r'^If\b', det) and find_word(det, 'Then') >= 0)
    stmts = []
    buf = []
    n = len(line)
    i = 0
    comment = None
    in_str = False
    while i < n:
        c = line[i]
        if c == '"':
            buf.append(c)
            if in_str and i + 1 < n and line[i + 1] == '"':
                buf.append('"')
                i += 2
                continue
            in_str = not in_str
            i += 1
            continue
        if c == "'" and not in_str:
            comment = line[i:].lstrip()
            break
        if c == ':' and not in_str and not whole_if \
                and not (i + 1 < n and line[i + 1] == '='):
            s = ''.join(buf).strip()
            if s:
                stmts.append(s)
            buf = []
            i += 1
            continue
        buf.append(c)
        i += 1
    s = ''.join(buf).strip()
    if s:
        stmts.append(s)
    return stmts, comment


def split_top_level(s, sep=','):
    parts, depth, cur, in_str = [], 0, [], False
    for ch in s:
        if ch == '"':
            in_str = not in_str
            cur.append(ch)
        elif in_str:
            cur.append(ch)
        elif ch in '([':
            depth += 1
            cur.append(ch)
        elif ch in ')]':
            depth -= 1
            cur.append(ch)
        elif ch == sep and depth == 0:
            parts.append(''.join(cur).strip())
            cur = []
        else:
            cur.append(ch)
    if cur:
        parts.append(''.join(cur).strip())
    return [p for p in parts if p]


def split_print_args(s):
    """Split VBA Print arg list on top-level ';' or ','; ignores spaces."""
    parts, cur = [], []
    depth = 0
    in_str = False
    for ch in s:
        if ch == '"':
            in_str = not in_str
            cur.append(ch)
        elif in_str:
            cur.append(ch)
        elif ch in '([':
            depth += 1
            cur.append(ch)
        elif ch in ')]':
            depth -= 1
            cur.append(ch)
        elif depth == 0 and ch in ';,':
            if ''.join(cur).strip():
                parts.append(''.join(cur).strip())
            cur = []
        else:
            cur.append(ch)
    if ''.join(cur).strip():
        parts.append(''.join(cur).strip())
    return parts


def find_word(text, word):
    """Index of first `word` outside strings/parens, or -1."""
    depth = 0
    in_str = False
    i = 0
    n = len(text)
    wl = len(word)
    while i <= n - wl:
        c = text[i]
        if c == '"':
            in_str = not in_str
            i += 1
            continue
        if not in_str:
            if c == '(':
                depth += 1
            elif c == ')':
                depth -= 1
            elif depth == 0 and text[i:i + wl] == word:
                before_ok = i == 0 or not (text[i - 1].isalnum() or text[i - 1] == '_')
                after_ok = i + wl >= n or not (text[i + wl].isalnum() or text[i + wl] == '_')
                if before_ok and after_ok:
                    return i
        i += 1
    return -1


def find_assign(t):
    """Index of the assignment '=' outside strings/parens, or -1."""
    depth = 0
    in_str = False
    i = 0
    n = len(t)
    while i < n:
        c = t[i]
        if c == '"':
            in_str = not in_str
        elif not in_str:
            if c in '([':
                depth += 1
            elif c in ')]':
                depth -= 1
            elif c == '=' and depth == 0:
                if i + 1 < n and t[i + 1] == '=':
                    i += 1
                elif i > 0 and t[i - 1] == ':':
                    i += 1
                else:
                    return i
        i += 1
    return -1


# ---------------------------------------------------------------------------
# expression rewriting (quote aware)
# ---------------------------------------------------------------------------

def code_parts(text):
    parts = []
    buf = []
    i = 0
    n = len(text)
    while i < n:
        c = text[i]
        if c == '"':
            if buf:
                parts.append(('C', ''.join(buf)))
                buf = []
            j = i + 1
            s = ['"']
            while j < n:
                if text[j] == '"':
                    if j + 1 < n and text[j + 1] == '"':
                        s.append('""')
                        j += 2
                        continue
                    s.append('"')
                    j += 1
                    break
                s.append(text[j])
                j += 1
            parts.append(('S', ''.join(s)))
            i = j
            continue
        buf.append(c)
        i += 1
    if buf:
        parts.append(('C', ''.join(buf)))
    return parts


def py_string_literal(s):
    if not (s.startswith('"') and s.endswith('"') and len(s) >= 2):
        return s
    body = s[1:-1]
    out = []
    while body:
        if body.startswith('""'):
            out.append('\\"')
            body = body[2:]
        elif body[0] == '\\':
            out.append('\\\\')
            body = body[1:]
        else:
            out.append(body[0])
            body = body[1:]
    return '"' + ''.join(out) + '"'


def _fix_code_segment(s, cond):
    s = re.sub(r'&H([0-9A-Fa-f]+)', r'0x\1', s)
    s = re.sub(r'&O([0-7]+)', r'0o\1', s)
    s = re.sub(r'\bNew\s+([A-Za-z_]\w*)\b', r'vba.new_object("\1")', s)
    s = re.sub(r'(?<=[A-Za-z0-9_])[&%#$!@](?![A-Za-z0-9_])', '', s)
    s = s.replace('&', '+')
    s = re.sub(r'\bMod\b', ' % ', s)
    s = re.sub(r'(?<=[\d\)\]])\s*\\\s*(?=[\d\(])', ' // ', s)
    s = s.replace('^', '**')
    s = s.replace('<>', '!=')
    s = re.sub(r'\bAnd\b', ' and ', s)
    s = re.sub(r'\bOr\b', ' or ', s)
    s = re.sub(r'\bNot\b', ' not ', s)
    s = re.sub(r'\bXor\b', ' ^ ', s)
    s = re.sub(r'\bIs\s+Not\b', ' is not ', s, flags=re.I)
    s = re.sub(r'\bIs\b', ' is ', s)
    for k, v in VB_STRING_CONST.items():
        s = re.sub(r'(?<![.\w])%s\b' % k, lambda m, r=v: r, s)
    s = re.sub(r'\bEmpty\b', 'None', s)
    s = re.sub(r'\bNothing\b', 'None', s)
    s = s.replace(':=', '=')
    if cond:
        s = re.sub(r'(?<![\=<>!])=(?![\=])', '==', s)
    return s


class ExprRewriter:
    PY_CALL_NAMES = {repl.split('.')[-1] for repl in BUILTIN_CALLS.values()}

    def __init__(self, proc_names, array_names):
        self.proc_names = set(proc_names)
        self.array_names = set(array_names)
        self._wv = None

    def _known_call(self, name):
        return (name in self.proc_names or name in BUILTIN_CALLS
                or name in EXCEL_TOP or name in EXCEL_CALL_NAMES
                or name in self.PY_CALL_NAMES)

    def _rewrite_code(self, s, cond):
        s = re.sub(r'(?<![.\w])UserForm(\d+)\b', r'vba.form(\1)', s)
        s = re.sub(r'(?<![.\w])(xl[A-Za-z0-9_]+|mso[A-Za-z0-9_]+)\b', r'vba.\1', s)
        for k in VB_CONST2:
            s = re.sub(r'(?<![.\w])%s\b' % k, 'vba.%s' % k, s)
        s = re.sub(r'(?<![.\w])(vb(?!a)[A-Za-z0-9_]+)\b', r'vba.\1', s)
        s = re.sub(r'\bApplication\b', 'vba.app', s)
        s = re.sub(r'\bThisWorkbook\b', 'vba.wb', s)
        s = re.sub(r'\bActiveWorkbook\b', 'vba.awb', s)
        s = re.sub(r'\bActiveSheet\b', 'vba.asheet', s)
        s = re.sub(r'\bActiveCell\b', 'vba.acell', s)
        s = re.sub(r'\bActiveWindow\b', 'vba.awin', s)
        s = re.sub(r'\bSelection\b', 'vba.sel', s)
        s = re.sub(r'\bWorksheetFunction\b(?=[.(])', 'vba.wf', s)
        s = re.sub(r'\bWorksheet\b(?=[.(])', 'vba.sheet', s)
        s = re.sub(r'\bWorkbooks\b(?=\s*\()', 'vba.workbooks', s)
        s = re.sub(r'(?<![.\w])Sheets\b(?=\s*\()', 'vba.wb.Sheets', s)
        s = re.sub(r'(?<![.\w])Cells\b(?=\s*\()', 'vba.asheet.Cells', s)
        s = re.sub(r'(?<![.\w])Range\b(?=\s*\()', 'vba.asheet.Range', s)
        s = re.sub(r'(?<![.\w])Charts\b(?=\s*\()', 'vba.wb.Charts', s)
        s = re.sub(r'\.(Select|Activate|Show|Hide|Delete|Insert|Copy|Clear|'
                   r'ClearContents|RemoveItem|AddItem|Cut|Merge|UnMerge|'
                   r'Calculate|Recalculate|Refresh|Quit)\b(?!\s*\()', r'.\1()', s)
        if self._wv:
            s = re.sub(r'(?<![A-Za-z0-9_.])\.([A-Za-z_]\w*)(?![A-Za-z0-9_])',
                       self._wv + r'.\1', s)
        if ' Like ' in s:
            s = re.sub(r'(.+?)\s+Like\s+("[^"]*"|[A-Za-z0-9_().]+)',
                       r'vba.like(\1, \2)', s)
        for name, repl in BUILTIN_CALLS.items():
            s = re.sub(r'(?<![\w$.])%s(?=\s*\()' % re.escape(name), repl, s)
        s = re.sub(r'(?<![.\w])MsgBox\s+(?!\()([^:]+)$', r'vba.MsgBox(\1)', s)
        return _fix_code_segment(self._unqual_bracket(s), cond)

    def _unqual_bracket(self, s):
        pat = re.compile(r'(?<![.\w])([A-Za-z_]\w*)(\s*'
                         r'\([^()]*(?:\([^()]*\))*[^()]*\))')

        def rep(m):
            nm, rest = m.group(1), m.group(2)
            if nm.lower() in PY_KEYWORDS or self._known_call(nm):
                return nm + rest
            return nm + '[' + rest[1:-1] + ']'

        prev = None
        while prev != s:
            prev = s
            s = pat.sub(rep, s)
        return s

    def rewrite(self, text, cond=False):
        if ' like ' in text.lower():
            text = self._wrap_like(text)
        if '"' not in text:
            return self._rewrite_code(text, cond)
        out = []
        for kind, part in code_parts(text):
            if kind == 'C':
                out.append(self._rewrite_code(part, cond))
            else:
                out.append(py_string_literal(part))
        return ''.join(out)

    # -- Like operator -------------------------------------------------
    def _find_like(self, text, start):
        n = len(text)
        i = start
        in_str = False
        while i < n:
            c = text[i]
            if c == '"':
                if in_str and i + 1 < n and text[i + 1] == '"':
                    i += 2
                    continue
                in_str = not in_str
                i += 1
                continue
            if in_str:
                i += 1
                continue
            if (text[i:i + 4].lower() == 'like'
                    and (i == 0 or not text[i - 1].isalnum())
                    and (i + 4 >= n or not text[i + 4].isalnum())
                    and (i == 0 or text[i - 1].isspace() or text[i - 1] in ',()')):
                return i
            i += 1
        return -1

    def _like_lhs_start(self, text, k, mask=None):
        depth = 0
        i = k - 1
        while i >= 0:
            if mask and mask[i]:
                i -= 1
                continue
            c = text[i]
            if c in ')]':
                depth += 1
            elif c in '([':
                if depth > 0:
                    depth -= 1
                else:
                    return i + 1
            elif depth == 0 and c in ',=<:<>':
                return i + 1
            elif depth == 0 and c.isspace():
                m = re.search(r'(And|Or|Then|Else|Not|Mod)\s*$', text[:i], re.I)
                if m and (i - len(m.group(1)) == 0
                          or not text[i - len(m.group(1)) - 1].isalnum()):
                    j = i
                    while j < k and text[j].isspace():
                        j += 1
                    return j
            i -= 1
        return 0

    def _like_rhs_end(self, text, start, mask=None):
        n = len(text)
        depth = 0
        i = start
        in_str = False
        while i < n:
            c = text[i]
            if c == '"':
                if in_str and i + 1 < n and text[i + 1] == '"':
                    i += 2
                    continue
                in_str = not in_str
                i += 1
                continue
            if in_str:
                i += 1
                continue
            if c in '([':
                depth += 1
            elif c == ')':
                if depth == 0:
                    return i
                depth -= 1
            elif depth == 0:
                if c in ',:<>=':
                    return i
                if c.isspace():
                    m = re.match(r'\s*(And|Or|Then|Else)\b', text[i:], re.I)
                    if m:
                        return i
            i += 1
        return n

    def _wrap_like(self, text):
        mask = [False] * len(text)
        in_str = False
        i = 0
        n = len(text)
        while i < n:
            c = text[i]
            if c == '"':
                if in_str and i + 1 < n and text[i + 1] == '"':
                    mask[i] = mask[i + 1] = True
                    i += 2
                    continue
                in_str = not in_str
                mask[i] = True
                i += 1
                continue
            if in_str:
                mask[i] = True
            i += 1
        res = []
        pos = 0
        while True:
            k = self._find_like(text, pos)
            if k < 0:
                res.append(text[pos:])
                return ''.join(res)
            lhs = self._like_lhs_start(text, k, mask)
            rhs = self._like_rhs_end(text, k + 4, mask)
            res.append(text[pos:lhs])
            res.append('vba.like(%s, (%s))' % (text[lhs:k].strip(),
                                               text[k + 4:rhs].strip()))
            pos = rhs


# ---------------------------------------------------------------------------
# module translator
# ---------------------------------------------------------------------------

class ModuleTranslator:
    def __init__(self, notes, proc_names, array_names, base=1):
        self.notes = notes
        self.proc_names = proc_names
        self.array_names = set(array_names)
        self.base = base
        self.out = []
        self.indent = 0
        self.stack = []
        self.sel_counter = 0
        self.with_counter = 0
        self.rewriter = ExprRewriter(proc_names, self.array_names)
        self.gosub_blocks = {}
        self.skip_indexes = set()
        self._in_type = False
        self._if_marks = []
        self.retvar = None
        self.try_wrap = False

    # -- basics ----------------------------------------------------
    def emit(self, text):
        self.out.append('    ' * self.indent + text if text else '')

    def peek(self):
        return self.stack[-1] if self.stack else None

    def pop(self):
        return self.stack.pop() if self.stack else None

    def _low_indent_for_chain(self):
        if self.peek() == 'if':
            self.indent -= 1
        elif self.peek() and isinstance(self.peek(), dict) \
                and self.peek()['type'] == 'select':
            self.indent -= 1

    # -- driver -----------------------------------------------------
    def run(self, stmts):
        i = 0
        while i < len(stmts):
            lno, text = stmts[i]
            if lno in self.skip_indexes:
                i += 1
                continue
            if isinstance(text, str) and text.strip() in self.gosub_blocks:
                self.emit('# VBA label; inlined at GoSub call sites: %s' % text.strip())
                i += 1
                continue
            gm = re.fullmatch(r'GoSub\s+([A-Za-z_]\w*)', text.strip())
            if gm:
                block = self.gosub_blocks.get(gm.group(1))
                if block is not None:
                    self.notes.add('gosub', 'GoSub %s inlined' % gm.group(1))
                    del stmts[i]
                    stmts[i:i] = block
                    self.skip_indexes = self.skip_indexes - {l for l, _ in block}
                    continue
                self.emit('# VBA GoSub %s (block not found — review)' % gm.group(1))
                i += 1
                continue
            self.translate_step(lno, text)
            i += 1

    # -- dispatcher -------------------------------------------------
    def translate_step(self, lno, text):
        t = text.strip()
        if t == '':
            self.emit('')
            return
        t = re.sub(r'^\d+(\.\d+)?\s*', '', t)  # VBA line-number labels
        t = t.strip()
        if t.startswith("'"):
            self.emit('#' + t[1:])
            return
        if re.search(r'^\s*(End\s+Sub|End\s+Function|End\s+Property)\b', t):
            self._close_proc(t)
            return
        if t.startswith('Option '):
            if re.match(r'^Option\s+Base\s+1$', t):
                self.base = 1
            elif re.match(r'^Option\s+Base\s+0$', t):
                self.base = 0
            self.emit('# VBA: %s' % t)
            return
        if t.startswith('Attribute '):
            self.emit('# VBA: %s' % t)
            return
        if re.match(r'^(Declare|Type|End\s+Type|Enum|End\s+Enum)\b', t):
            self.emit('# VBA: %s (Windows API / UDT — not converted)' % t)
            self.notes.add('declare', t)
            return
        if re.match(r'^On\s+Error\b', t):
            self.emit('# VBA: %s   (errors suppressed via try/except wrapper)' % t)
            return
        if re.match(r'^Resume\b', t):
            self.emit('# VBA: %s' % t)
            return
        m = re.match(r'^Exit\s+(Sub|Function|For|Do|While|Property)$', t)
        if m:
            k = m.group(1)
            if k in ('For', 'Do', 'While'):
                self.emit('break')
            elif k in ('Function', 'Property'):
                self.emit('return _out_')
            else:
                self.emit('return')
            return
        if re.fullmatch(r'End', t):
            self.emit('sys.exit()  # VBA End')
            return
        # --- If / ElseIf / Else / End If ---
        if re.match(r'^End\s+If\b', t):
            if self.peek() == 'if':
                mark = self._if_marks.pop() if self._if_marks else len(self.out)
                if not any(x.strip() and not x.strip().startswith('#')
                           for x in self.out[mark:]):
                    self.emit('pass')
                self.indent -= 1
                self.pop()
            else:
                self.emit('# VBA End If (no matching If)')
            return
        if re.match(r'^If\s+', t):
            self._handle_if(t)
            return
        if re.match(r'^ElseIf\s+', t):
            cm = re.match(r'^ElseIf\s+(.+?)\s+Then\s*$', t)
            if cm and self.peek() == 'if':
                self.indent -= 1
                self.emit('elif %s:' % self.rewriter.rewrite(cm.group(1), cond=True))
                self.indent += 1
            else:
                self.emit('# REVIEW ElseIf')
            return
        if re.fullmatch(r'Else', t):
            if self.peek() == 'if':
                self.indent -= 1
                self.emit('else:')
                self.indent += 1
            else:
                self.emit('# REVIEW Else')
            return
        # --- For / Next ---
        m = re.match(r'^For\s+Each\s+(\w+)\s+In\s+(.+)$', t)
        if m:
            self.emit('for %s in %s:' % (m.group(1), self.rewriter.rewrite(m.group(2))))
            self.stack.append('foreach')
            self.indent += 1
            return
        m = re.match(r'^For\s+(\w+)\s*=\s*(.+?)\s+To\s+(.+?)(?:\s+Step\s+(.+))?$', t)
        if m:
            v, lo, hi, st = m.group(1), m.group(2), m.group(3), m.group(4)
            lo, hi = self.rewriter.rewrite(lo), self.rewriter.rewrite(hi)
            if st is not None and re.match(r'^\s*-\s*', st):
                stn = self.rewriter.rewrite(re.sub(r'^\s*-\s*', '', st))
                self.emit('for %s in range(%s, (%s) - 1, -%s):' % (v, lo, hi, stn))
            else:
                self.emit('for %s in range(%s, (%s) + 1, %s):' % (v, lo, hi,
                                                                   self.rewriter.rewrite(st or '1')))
            self.stack.append('for')
            self.indent += 1
            return
        m = re.fullmatch(r'Next(?:\s+([\w, ]+)\s*)?', t)
        if m:
            n = len(split_top_level(m.group(1))) if m.group(1) and m.group(1).strip() else 1
            cnt = 0
            while cnt < n and self.peek() in ('for', 'foreach'):
                self.indent -= 1
                self.pop()
                cnt += 1
            for _ in range(max(0, n - cnt)):
                self.emit('pass  # VBA Next without matching For')
            return
        # --- Do / Loop / While / Wend ---
        if re.fullmatch(r'Do', t):
            self.emit('while True:')
            self.stack.append('do')
            self.indent += 1
            return
        m = re.match(r'^Do\s+(While|Until)\s+(.+)$', t)
        if m:
            kw, cond = m.group(1), self.rewriter.rewrite(m.group(2), cond=True)
            self.emit('while %s:' % (cond if kw == 'While' else 'not (%s)' % cond))
            self.stack.append('do')
            self.indent += 1
            return
        if re.fullmatch(r'Loop', t):
            if self.peek() == 'do':
                self.indent -= 1
                self.pop()
            else:
                self.emit('pass  # VBA Loop')
            return
        m = re.match(r'^Loop\s+(While|Until)\s+(.+)$', t)
        if m:
            kw, cond = m.group(1), self.rewriter.rewrite(m.group(2), cond=True)
            self.emit('if %s: break' % (cond if kw == 'Until' else 'not (%s)' % cond))
            if self.peek() == 'do':
                self.indent -= 1
                self.pop()
            return
        m = re.match(r'^While\s+(.+)$', t)
        if m:
            self.emit('while %s:' % self.rewriter.rewrite(m.group(1), cond=True))
            self.stack.append('while')
            self.indent += 1
            return
        if re.fullmatch(r'Wend', t):
            if self.peek() == 'while':
                self.indent -= 1
                self.pop()
            return
        # --- Select Case ---
        m = re.match(r'^Select\s+Case\s+(.+)$', t)
        if m:
            self.sel_counter += 1
            sv = '_sel_%d' % self.sel_counter
            self.emit('%s = %s' % (sv, self.rewriter.rewrite(m.group(1))))
            self.stack.append({'type': 'select', 'selvar': sv, 'any': False})
            return
        m = re.match(r'^Case\s*(.*)$', t)
        if m and isinstance(self.peek(), dict) and self.peek()['type'] == 'select':
            b = self.peek()
            first = not b['any']
            b['any'] = True
            if not first:
                self.indent -= 1
            rest = m.group(1).strip()
            if rest.lower() == 'else':
                self.emit('else:')
            else:
                conds = [self._case_item_cond(b['selvar'], it)
                         for it in split_top_level(rest)]
                self.emit(('if ' if first else 'elif ') + ' or '.join(conds) + ':')
            self.indent += 1
            return
        if re.match(r'^End\s+Select\b', t):
            if isinstance(self.peek(), dict) and self.peek()['type'] == 'select':
                self.indent -= 1
                self.pop()
            return
        # --- With ---
        m = re.match(r'^With\s+(.+)$', t)
        if m:
            self.with_counter += 1
            wv = 'vw_%d' % self.with_counter
            self.emit('%s = %s' % (wv, self.rewriter.rewrite(m.group(1))))
            self.stack.append({'type': 'with', 'var': wv})
            self.rewriter._wv = wv
            return
        if re.match(r'^End\s+With\b', t):
            if isinstance(self.peek(), dict) and self.peek()['type'] == 'with':
                self.rewriter._wv = None
                self.pop()
            return
        # --- Type / End Type ---
        if re.match(r'^(?:Private\s+|Public\s+)?Type\s+\w+', t):
            self.emit('# VBA Type block: %s' % t)
            self._in_type = True
            return
        if self._in_type:
            if re.match(r'^End\s+Type\b', t):
                self._in_type = False
            return
        # --- declarations ---
        m = re.match(r'^ReDim\s+(Preserve\s+)?(\w+)\s*(?:\((.*)\))?(?:\s+As\s+.+)?$', t, re.S)
        if m:
            self._alloc_array(m.group(2), m.group(3) or '')
            return
        m = re.match(r'^(?:Dim|Private|Public|Global|Static)\s+(.+)$', t)
        if m:
            self._handle_dim(m.group(1))
            return
        m = re.match(r'^(?:Public|Private)?\s*Const\s+(.+)$', t)
        if m:
            self._handle_const(m.group(1))
            return
        if re.match(r'^(?:Public|Private|Global)\s+', t) and ' As ' in t:
            self.emit('# VBA: %s   (declaration)' % t)
            return
        # --- file I/O ---
        if re.match(r'^Kill\s+', t):
            self.emit('vba.Kill(%s)' % self.rewriter.rewrite(re.sub(r'^Kill\s+', '', t)))
            return
        if re.match(r'^FileCopy\s+', t):
            parts = split_top_level(re.sub(r'^FileCopy\s+', '', t))
            if len(parts) >= 2:
                self.emit('shutil.copy(%s, %s)' % (self.rewriter.rewrite(parts[0]),
                                                   self.rewriter.rewrite(parts[1])))
            return
        m = re.match(r'^Open\s+(.+?)\s+For\s+(Input|Output|Append|Binary|Random)\s*'
                 r'As\s+#?\s*([A-Za-z0-9_]+)(.*)$', t)
        if m:
            path, mode, num, extra = m.groups()
            py_mode = {'Input': 'r', 'Output': 'w', 'Append': 'a',
                       'Binary': 'rb', 'Random': 'rb+'}[mode]
            num = self.rewriter.rewrite(num)
            self.emit('vba.file_open(%s, %s, %r)' % (
                num, self.rewriter.rewrite(path), py_mode))
            return
        m = re.fullmatch(r'Close(?:\s*#?\s*([A-Za-z0-9_]+))?', t)
        if m:
            if m.group(1):
                self.emit('vba.file_close(%s)' % self.rewriter.rewrite(m.group(1)))
            else:
                self.emit('vba.file_close_all()')
            return
        m = re.match(r'^Line\s+Input\s+#?\s*(\d+|[A-Za-z_]\w*)\s*,\s*(.+)$', t)
        if m:
            self.emit('%s = vba.file_line_input(%s)' % (m.group(2), m.group(1)))
            return
        m = re.match(r'^Input\s+#?\s*(\d+|[A-Za-z_]\w*)\s*,\s*(.+)$', t)
        if m:
            self.emit('%s = vba.file_input(%s)' % (m.group(2), m.group(1)))
            return
        m = re.match(r'^Print\s+#?\s*(\d+|[A-Za-z_]\w*)\s*,\s*(.*)$', t)
        if m:
            vals = [self.rewriter.rewrite(v) for v in split_print_args(m.group(2))]
            argstr = ', '.join(vals)
            self.emit('vba.file_print(%s%s)' % (m.group(1), ', ' + argstr if argstr else ''))
            return
        m = re.match(r'^Write\s+#?\s*(\d+|[A-Za-z_]\w*)\s*,\s*(.*)$', t)
        if m:
            vals = [self.rewriter.rewrite(v) for v in split_top_level(m.group(2))]
            argstr = ', '.join(vals)
            self.emit('vba.file_write(%s%s)' % (m.group(1), ', ' + argstr if argstr else ''))
            return
        m = re.match(r'^(Get|Put)\s+#?\s*(\d+|[A-Za-z_]\w*)\s*,.*$', t)
        if m:
            self.emit('# VBA %s #%s (binary I/O — review manually)' % (m.group(1).upper(),
                                                                        m.group(2)))
            return
        if re.match(r'^Unload\s+', t):
            self.emit('vba.form_unload(%s)' % re.sub(r'^Unload\s+', '', t).strip())
            return
        m = re.fullmatch(r'GoTo\s+([A-Za-z_]\w*)', t)
        if m:
            self.emit('# REVIEW: VBA GoTo %s (not convertible — control flow)' % m.group(1))
            self.notes.add('goto', 'GoTo %s not converted' % m.group(1))
            return
        if re.match(r'^Debug\.Print\s*(.*)$', t):
            rest = re.match(r'^Debug\.Print\s*(.*)$', t).group(1).strip()
            self.emit('print(%s)' % self.rewriter.rewrite(rest, cond=False) if rest
                      else 'print()')
            return
        m = re.match(r'^\?\s*(.+)$', t)
        if m:
            self.emit('print(%s)' % self.rewriter.rewrite(m.group(1)))
            return
        if re.match(r'^Let\s+', t):
            t = re.sub(r'^Let\s+', '', t)
        if t.startswith('Call '):
            t = t[len('Call '):]
            if not t.endswith(')'):
                t += '()'
        if t.startswith('.') and isinstance(self.peek(), dict) and self.peek()['type'] == 'with':
            t = self.peek()['var'] + t
        self._emit_statement(t)

    # -- If ---------------------------------------------------------
    def _handle_if(self, t):
        then_pos = find_word(t, 'Then')
        if then_pos < 0:
            self.emit('# REVIEW unparsed If: %s' % t)
            return
        cond = t[3:then_pos].strip()
        rest = t[then_pos + 4:].strip()
        if rest == '':
            self.emit('if %s:' % self.rewriter.rewrite(cond, cond=True))
            self.stack.append('if')
            self._if_marks.append(len(self.out))
            self.indent += 1
            return
        else_pos = find_word(rest, 'Else')
        if else_pos < 0:
            then_part, else_part = rest, None
        else:
            then_part = rest[:else_pos].strip()
            else_part = rest[else_pos + 4:].strip()
        self.emit('if %s:' % self.rewriter.rewrite(cond, cond=True))
        self.indent += 1
        mark = len(self.out)
        self.run([(0, s) for s in self._split_into(then_part)])
        if not any(x.strip() and not x.strip().startswith('#')
                   for x in self.out[mark:]):
            self.emit('pass')
        self.indent -= 1
        if else_part is not None:
            self.emit('else:')
            self.indent += 1
            mark = len(self.out)
            self.run([(0, s) for s in self._split_into(else_part)])
            if not any(x.strip() and not x.strip().startswith('#')
                       for x in self.out[mark:]):
                self.emit('pass')
            self.indent -= 1

    @staticmethod
    def _split_into(part):
        stmts, _ = split_physical(part)
        return stmts

    # -- Select Case helper ------------------------------------------
    def _case_item_cond(self, selvar, item):
        item = item.strip()
        m = re.match(r'^Is\s*(<>|<=|>=|<|>|=)\s*(.+)$', item)
        if m:
            op = {'=': '==', '<>': '!='}.get(m.group(1), m.group(1))
            return '%s %s (%s)' % (selvar, op, self.rewriter.rewrite(m.group(2), cond=True))
        m = re.match(r'^(.+?)\s+To\s+(.+)$', item)
        if m:
            lo = self.rewriter.rewrite(m.group(1), cond=True)
            hi = self.rewriter.rewrite(m.group(2), cond=True)
            return 'vba.between(%s, %s, %s)' % (selvar, lo, hi)
        return '%s == (%s)' % (selvar, self.rewriter.rewrite(item, cond=True))

    # -- declarations ------------------------------------------------
    def _alloc_array(self, name, dims):
        dims = dims.strip()
        if not dims:
            return
        exprs = split_top_level(dims)
        if len(exprs) == 1 and re.search(r'\bto\b', exprs[0], re.I):
            m = re.match(r'^(.*?)\s+[Tt][Oo]\s+(.+)$', exprs[0])
            if m:
                lo = self.rewriter.rewrite(m.group(1))
                hi = self.rewriter.rewrite(m.group(2))
                self.emit('%s = [None] * ((%s) - (%s) + 1)' % (name, hi, lo))
                self.emit('# NOTE: VBA array %s is %s-based (VBA %s..%s)' % (
                    name, m.group(1).strip(), m.group(1).strip(), m.group(2).strip()))
                self.array_names.add(name)
                return
        base = 1 if self.base == 1 else 0
        if len(exprs) == 1:
            e = exprs[0]
            m = re.match(r'^(.*?)\s+[Tt][Oo]\s+(.+)$', e)
            if m:
                lo = self.rewriter.rewrite(m.group(1))
                hi = self.rewriter.rewrite(m.group(2))
                self.emit('%s = [None] * ((%s) - (%s) + 1)' % (name, hi, lo))
            else:
                self.emit('%s = [None] * ((%s) + 1)' % (name, self.rewriter.rewrite(e)))
            if base == 1:
                self.emit('# NOTE: VBA array %s is 1-based; index 0 stays unused' % name)
        else:
            sizes = []
            for e in exprs:
                m = re.match(r'^(.*?)\s+[Tt][Oo]\s+(.+)$', e)
                if m:
                    lo = self.rewriter.rewrite(m.group(1))
                    hi = self.rewriter.rewrite(m.group(2))
                    sizes.append('((%s) - (%s) + 1)' % (hi, lo))
                else:
                    sizes.append('((%s) + 1)' % self.rewriter.rewrite(e))
            inner = 'None'
            for s in reversed(sizes):
                inner = '[%s for _ in range(%s)]' % (inner, s)
            self.emit('%s = %s' % (name, inner))
            self.emit('# NOTE: VBA array %s is declared with Option Base %d; 1-based '
                      'indexing is handled by accessing indices 1..N' % (name, 1 if base else 0))
        self.array_names.add(name)

    def _handle_dim(self, body):
        made = False
        for sp in split_top_level(body):
            sp = sp.strip()
            if not sp:
                continue
            m = re.match(r'^(\w+)\s*\((.+)\)(?:\s+As\s+.+)?$', sp, re.S)
            if m:
                self._alloc_array(m.group(1), m.group(2))
                made = True
                continue
            m = re.match(r'^(\w+)(?:\s+As\s+.+)?$', sp)
            if m:
                made = True
        if not made:
            self.emit('# VBA Dim: %s' % body)

    def _handle_const(self, body):
        for sp in split_top_level(body):
            sp = sp.strip()
            m = re.match(r'^(\w+)\s*(?:As\s+\w+\s+)?\s*=\s*(.+)$', sp)
            if m:
                name, val = _fix_ident(m.group(1)), self.rewriter.rewrite(m.group(2))
                self.emit('%s = %s' % (name, val))
            else:
                self.emit('# VBA Const: %s' % sp)

    # -- assignment / expression -------------------------------------
    def _emit_statement(self, t):
        t = re.sub(r'^\d+\s*', '', t)
        if t.startswith('Set '):
            t = re.sub(r'^Set\s+', '', t)
        eq = find_assign(t)
        if eq < 0:
            depth = 0
            in_str = False
            sp = -1
            for j, c in enumerate(t):
                if c == '"':
                    in_str = not in_str
                elif not in_str:
                    if c in '([':
                        depth += 1
                    elif c in ')]':
                        depth -= 1
                    elif c.isspace() and depth == 0 and j > 0 \
                            and t[j - 1] not in '([.,+-*/=<>':
                        sp = j
                        break
            if sp > 0:
                target = t[:sp].strip()
                args = t[sp:].strip()
                if ':=' in t:
                    t = '%s(%s)' % (target,
                                    re.sub(r'\b([A-Za-z_]\w*):=', r'\1=', args))
                elif not target.endswith(')'):
                    t = '%s(%s)' % (target, args)
            self.emit(self.rewriter.rewrite(t))
            return
        lhs, rhs = t[:eq].strip(), t[eq + 1:].strip()
        if self.retvar is not None and lhs == self.retvar:
            lhs_rw = '_out_'
        else:
            lhs_rw = self.rewriter.rewrite(_fix_ident(lhs) if lhs in PY_KEYWORDS else lhs)
        if self.retvar is None and lhs == '_out_':
            pass
        if lhs_rw.endswith(')') and re.search(r'\.(Cells|Range|Offset)\s*\([^()]*\)$', lhs_rw):
            lhs_rw += '.Value'
        elif lhs_rw.endswith(')') and re.fullmatch(r'(vba\.)?(Cells|Range)\s*\([^()]*\)', lhs_rw):
            lhs_rw += '.Value'
        lhs_rw = re.sub(r'\.List\(\)$', 'List', lhs_rw)
        if not lhs_rw.endswith('.Value'):
            lhs_rw = re.sub(r'\(([^()]*)\)(\.[A-Za-z_]\w*)$', r'[\1]\2', lhs_rw)
        if lhs_rw.endswith(')') and not lhs_rw.endswith('.Value'):
            lhs_rw = re.sub(r'\(([^()]*)\)$', r'[\1]', lhs_rw)
        self.emit('%s = %s' % (lhs_rw, self.rewriter.rewrite(rhs, cond=True)))
        m = re.match(r'^(\w+)\s*=\s*(?:Array|Split)\s*\(', t)
        if m:
            self.array_names.add(m.group(1))

    # -- procedure framing -------------------------------------------
    def _close_proc(self, t):
        if self.retvar is not None and re.search(r'End\s+(Function|Property)', t):
            self.emit('return _out_')
        if self.peek() == 'proc':
            self.indent -= 1
            self.pop()
        self.retvar = None
        self.try_wrap = False


def _parse_logical(lines):
    logical = []
    buf = None
    start = 0
    for i, ln in enumerate(lines):
        stripped = ln.strip()
        if stripped.startswith("'"):
            if buf is None:
                logical.append((i, stripped))
            else:
                buf += ' ' + stripped
                logical.append((start, buf.strip()))
                buf = None
            continue
        cont = ln.rstrip().endswith('_')
        if buf is None:
            if cont:
                buf = ln.rstrip()[:-1].strip()
                start = i
            else:
                logical.append((i, stripped))
        else:
            buf += ' ' + (stripped[:-1] if cont else stripped)
            if not cont:
                logical.append((start, buf.strip()))
                buf = None
    if buf is not None:
        logical.append((start, buf.strip()))
    return logical


def _split_procs(logical):
    blocks = []
    cur = []
    n = len(logical)
    i = 0
    while i < n:
        lno, t = logical[i]
        if re.match(r'^\s*(?:Public|Private|Friend|Global)?\s*'
                    r'(?:Sub|Function|Property\s+\w+)\b', t):
            if cur:
                blocks.append(('module', cur))
                cur = []
            header = logical[i]
            body = []
            j = i + 1
            while j < n:
                lt = logical[j][1]
                if re.match(r'^\s*(?:Public|Private|Friend|Global)?\s*'
                            r'(?:Sub|Function|Property\s+\w+)\b', lt):
                    break
                body.append(logical[j])
                if re.match(r'^\s*End\s+(?:Sub|Function|Property)\b', lt):
                    j += 1
                    break
                j += 1
            blocks.append(('proc', [header, body]))
            i = j
            continue
        cur.append(logical[i])
        i += 1
    if cur:
        blocks.append(('module', cur))
    return blocks


def _gosub_blocks(body):
    labels = {}
    for i, (lno, t) in enumerate(body):
        m = re.fullmatch(r'([A-Za-z_]\w*)\s*:', t.strip())
        if m:
            labels.setdefault(m.group(1), i)
    gosub = {}
    ordered = sorted(labels.items(), key=lambda kv: kv[1])
    for name, idx in ordered:
        blk = []
        j = idx + 1
        while j < len(body):
            lt = body[j][1]
            if re.fullmatch(r'[A-Za-z_]\w*\s*:', lt.strip()):
                break
            if re.match(r'^\s*Return\s*$', lt):
                break
            for s_ in split_physical(lt)[0]:
                blk.append((body[j][0], s_))
            j += 1
        gosub[name] = blk
    used = set()
    for _, t in body:
        m = re.match(r'GoSub\s+([A-Za-z_]\w*)', t.strip())
        if m:
            used.add(m.group(1))
    skip = set()
    for name in used:
        if name in gosub:
            skip |= {ln for ln, _ in gosub[name]}
    return gosub, skip


def _python_args(args):
    out = []
    for a in split_top_level(args):
        a = a.strip()
        if not a:
            continue
        if a.startswith('ParamArray '):
            a = a[len('ParamArray '):]
        a = re.sub(r'^(?:Optional|ByVal|ByRef|ByVal\s+|ByRef\s+)\s*', '', a)
        if a.endswith('()'):
            a = a[:-2]
        m = re.match(r'^(\w+)\s*(?:As\s+\w+(?:\s*\(\d+\))?)?\s*(?:=\s*(.+))?$', a)
        if not m:
            continue
        name, default = m.group(1), m.group(2)
        out.append('%s=%s' % (_fix_ident(name), default if default is not None else 'None'))
    return ', '.join(out)


def _translate_proc_body(mt, header, body):
    m = re.match(r'^\s*(?:Public|Private|Friend|Global)?\s*'
                 r'((?:Sub|Function|Property\s+(?:Get|Let|Set)))\s+'
                 r'([A-Za-z_]\w*)\s*(?:\((.*)\))?(?:\s+As\s+.+)?$', header)
    if not m:
        m = re.match(r'^\s*(?:Public|Private|Friend|Global)?\s*'
                     r'((?:Sub|Function|Property\s+\w+))\s+'
                     r'([A-Za-z_]\w*)\s*(?:\((.*)\))?(?:\s+As\s+.+)?$', header)
    if not m:
        return ['# REVIEW (unparsed header): %s' % header]
    prok, name, args = m.group(1), m.group(2), m.group(3)
    retvar = name if prok in ('Function', 'Property Get') else None
    mt.retvar = retvar
    mt.gosub_blocks, mt.skip_indexes = _gosub_blocks(body)
    mt.try_wrap = any(re.match(r'^On\s+Error\b', b[1].strip()) for b in body)
    mt.stack = []
    mt.retvar = None
    mt.rewriter._wv = None
    mt.indent = 1
    stmts = []
    for b in body:
        s, cmt = split_physical(b[1])
        stmts.extend((b[0], s_) for s_ in s)
    lines = ['def %s(%s):' % (_fix_ident(name), _python_args(args) if args and args.strip() else '')]
    if retvar is not None:
        lines.append('    _out_ = None')
    start = len(mt.out)
    mt.run(stmts)
    lines += mt.out[start:]
    mt.out = mt.out[:start]
    if mt.try_wrap:
        lines = [lines[0], '    try:'] + \
                ['    ' + (ln if ln == '' else ln) for ln in lines[1:]] + \
                ['    except Exception:', '        pass']
    return lines


def translate_module(code, modname, proc_names, notes):
    logical = _parse_logical(code.splitlines())
    base = 1
    for _, t in logical:
        if re.match(r'^Option\s+Base\s+1\b', t):
            base = 1
        elif re.match(r'^Option\s+Base\s+0\b', t):
            base = 0
    arrays = set()
    for _, t in logical:
        for m in re.finditer(r'\b(?:Dim|Private|Public|ReDim)\s+([A-Za-z_]\w*)\s*\([^)]*\)', t):
            arrays.add(m.group(1))
    mt = ModuleTranslator(notes, proc_names, arrays, base)
    out = ['# ===== module: %s =====' % modname]
    for kind, items in _split_procs(logical):
        if kind == 'module':
            stmts = [(l, t) for l, t in items]
            mt.run(stmts)
            out.extend(mt.out)
            mt.out = []
        else:
            header, body = items
            htext = header[1]
            if re.search(r':\s*$', htext) or not re.match(r'^\s*(?:Public|Private|Friend|Global)?\s*'
                                                          r'(Sub|Function|Property\s+\w+)', htext):
                out.append('# REVIEW header (continued): %s' % htext)
                continue
            out.extend(_translate_proc_body(mt, htext, body))
    return '\n'.join(out)


# ---------------------------------------------------------------------------
# workbook driver
# ---------------------------------------------------------------------------

def _scan_proc_names(code):
    names = set()
    for ln in code.splitlines():
        m = re.match(r'^\s*(?:Public|Private|Friend|Global)?\s*'
                     r'(?:Sub|Function|Property\s+\w+)\s+([A-Za-z_]\w*)', ln.strip())
        if m:
            names.add(m.group(1))
    return names


def convert_one(path, outdir, notes, no_runtime):
    base = os.path.splitext(os.path.basename(path))[0]
    os.makedirs(outdir, exist_ok=True)
    if not no_runtime:
        _emit_runtime(os.path.join(outdir, 'vba_runtime.py'))
    try:
        parser = VBA_Parser(str(path))
    except Exception as e:  # noqa
        notes.add('extract', '%s: parse error %r' % (os.path.basename(path), e))
        return
    modules = []
    for _f, _s, vfn, code in parser.extract_macros():
        lower = vfn.lower()
        if lower in ('thisworkbook',) or re.fullmatch(r'sheet\d*', lower):
            kind = 'sheet'
        elif lower.endswith('.frm') or lower.startswith('userform'):
            kind = 'form'
        elif lower.endswith('.cls'):
            kind = 'class'
        else:
            kind = 'module'
        name = vfn.split('/')[-1]
        modules.append((kind, name, code))
    parser.close()

    proc_names = set()
    for _, _, code in modules:
        proc_names |= _scan_proc_names(code)
    decl_names = set()
    for _, _, code in modules:
        for ln in code.splitlines():
            m = re.match(r'^\s*Declare\s+(?:Function|Sub)\s+([A-Za-z_]\w*)', ln)
            if m:
                decl_names.add(m.group(1))
    proc_names |= decl_names

    bas_out = []
    singles = []
    for kind, name, code in modules:
        notes_mod = Notes()
        txt = translate_module(code, name, proc_names, notes_mod)
        header = '# ==== %s [%s] ====' % (name, kind)
        if kind == 'module':
            bas_out.append(header + '\n' + txt + '\n')
        else:
            singles.append((kind, name, header + '\n' + txt + '\n'))
    with open(os.path.join(outdir, '%s.py' % base), 'w', encoding='utf-8') as fh:
        fh.write('# AUTO-CONVERTED from %s by vba2py (best-effort port).\n'
                 '# Review the code and the _report.txt before use.\n\n'
                 % os.path.basename(path))
        fh.write('import math\nimport os\nimport random\nimport shutil\n'
                 'import subprocess\nimport sys\nimport vba_runtime as vba\n\n')
        fh.writelines(bas_out)
    for kind, name, txt in singles:
        with open(os.path.join(outdir, '%s__%s.py' % (base, name)), 'w',
                  encoding='utf-8') as fh:
            fh.write('# AUTO-CONVERTED from %s :: %s [%s module]\n'
                     '# VBA class/UserForm modules carry UI/event context; the\n'
                     '# procedures below are kept as module-level functions.\n\n'
                     '' % (os.path.basename(path), name, kind))
            fh.write('import math\nimport os\nimport random\nimport shutil\n'
                     'import subprocess\nimport sys\nimport vba_runtime as vba\n\n')
            fh.write(txt)
            fh.write('\n')
    # report
    report = []
    report.append('Conversion report for %s\n' % os.path.basename(path))
    report.append('Generated by vba2py -- best-effort VBA->Python translation.\n')
    for kind, msg in notes.items:
        report.append('  [%-8s] %s\n' % (kind, msg))
    with open(os.path.join(outdir, '%s_report.txt' % base), 'w', encoding='utf-8') as fh:
        fh.writelines(report)


def _emit_runtime(target):
    here = os.path.dirname(os.path.abspath(__file__))
    src = os.path.join(here, 'vba_runtime.py')
    with open(src, 'r', encoding='utf-8') as fh:
        data = fh.read()
    with open(target, 'w', encoding='utf-8') as fh:
        fh.write(data)


def find_excel_files(paths):
    found = []
    for p in paths:
        ap = os.path.abspath(p)
        if os.path.isfile(ap):
            if ap.lower().endswith(('.xls', '.xlsm', '.xlsb')):
                found.append(ap)
        elif os.path.isdir(ap):
            for dirpath, _, fns in os.walk(ap):
                for fn in fns:
                    if fn.lower().endswith(('.xls', '.xlsm', '.xlsb')):
                        found.append(os.path.join(dirpath, fn))
    return sorted(set(found))


def main(argv=None):
    ap = argparse.ArgumentParser(prog='vba2py', description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('paths', nargs='*', help='.xls/.xlsm files or directories')
    ap.add_argument('-o', '--outdir', help='output directory (default: input file directory)')
    ap.add_argument('--no-runtime', action='store_true',
                    help='do not copy vba_runtime.py into the output directory')
    ap.add_argument('--quiet', action='store_true')
    args = ap.parse_args(argv)

    paths = args.paths or ['.']
    files = find_excel_files(paths)
    if not files:
        print('No .xls/.xlsm workbooks found under: %s' % ' '.join(paths))
        return 1
    notes = Notes()
    for f in files:
        odir = args.outdir or os.path.dirname(f)
        convert_one(f, odir, notes, args.no_runtime)
        if not args.quiet:
            print('converted %s' % f)
    if not args.quiet:
        agg = {}
        for kind, _ in notes.items:
            agg[kind] = agg.get(kind, 0) + 1
        if agg:
            print('\nNotes summary (see *_report.txt):')
            for kind, n in sorted(agg.items()):
                print('  %-10s %d' % (kind, n))
    return 0


if __name__ == '__main__':
    sys.exit(main())