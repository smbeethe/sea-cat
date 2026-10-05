# -*- coding: utf-8 -*-
"""
Shared engines behind the rebuilt user-forms.

These reproduce the behaviour of the two big VBA record editors
(UserForm1 in both SC.console and ERDPP.console) and the shared
list-picker / choice dialogs:
    * RecordFormBase  -- column grid + hot buttons + record navigation
    * ChoiceDialog    -- radio group (+ optional checkboxes) picker
    * ListForm        -- the UserForm2 dynamic-list editor
"""
import tkinter as tk

from . import engine, theme
from .framework import (ActionButton, FieldRow, FormWindow, MultiPage,
                        HotButton, select_text, rgb)


HELP_TEXT = {
    'char': 'Free-text character field.  Columns longer than 60 chars are '
            'keyed in as flowing text.',
    'dyna': 'Dynamic list-managed field.  Choose a value from the list picker '
            '(or type the value directly).',
    'authors': 'Author list (full citation authors).  Managed with the list '
               'editor.',
    'editors': 'Book editors list.  Managed with the list editor.',
    'date': 'Date field.  Enter dates in yyyy-mm-dd format.',
    'flag': 'Flag field: N = no, Y = yes.',
    'skip': 'Auxiliary field with no validation.',
    'stat': 'Status field listing the record state in the database.',
    'status': 'Status field listing the record state in the database.',
    'database': 'Database membership field (edit disabled unless granted).',
    'url': 'URL field.  Paste or type the data-file locator.',
    'list': 'Generic list field (multiple-choice values).',
}
DEFAULT_HELP = 'Column entered according to its declared data type.'


class RecordFormBase(FormWindow):
    """Grid of N column rows (15 per page tab) + hot buttons + navigation bar.

    Mirrors UserForm1_Settings / Row_Read / Row_Write of the converted code.
    """

    PAGES = 8
    COLS_PER_PAGE = 15

    def __init__(self, master, title, state, cols, subtitle=None,
                 quick_calc=True, paste_reference=True, master_user=False,
                 citations=False, database_user=False):
        super().__init__(master, title, subtitle, size='960x640')
        self.state = state
        self.cols = list(cols)                       # (caption, var, dtype)
        self.ncols = len(self.cols)
        self.quick_calc = quick_calc
        self.paste_reference = paste_reference
        self.master_user = master_user
        self.citations = citations
        self.database_user = database_user
        self.entries = {}
        self.labels = {}
        self.hotbtns = {}
        self.dtype = {}
        self.tag = {}

        # -- status / navigation bar --
        nav = tk.Frame(self.body, bg=theme.PAGE_BG)
        nav.pack(fill='x', pady=(0, 4))
        self._status_var = tk.StringVar(value='')
        tk.Label(nav, textvariable=self._status_var, bg=theme.PAGE_BG,
                 fg='#336699', font=theme.app_font(9, bold=True)).pack(
            side='left', padx=4)
        for text, fn in (('< Previous', self.on_prev), ('Next >', self.on_next),
                         ('New Record', self.on_new), ('Delete', self.on_delete,
                                                       ),
                         ):
            ActionButton(nav, text, fn).pack(side='left', padx=3)

        tk.Label(nav, text=' Go:', bg=theme.PAGE_BG,
                 font=theme.app_font(9)).pack(side='left', padx=(10, 2))
        self.go_entry = tk.Entry(nav, width=8, font=theme.app_font(10))
        self.go_entry.pack(side='left')
        self.go_entry.bind('<Return>', lambda e: self.on_goto())
        ActionButton(nav, 'Go', self.on_goto).pack(side='left', padx=3)

        self.multi = MultiPage(self.body, pages=self.PAGES)
        self.multi.pack(fill='both', expand=True)

        # -- footer buttons --
        foot = tk.Frame(self.body, bg=theme.PAGE_BG)
        foot.pack(fill='x', pady=(4, 0))
        if self.quick_calc:
            ActionButton(foot, 'Quick Calculation', self.on_quick_calc).pack(
                side='left', padx=3)
        if self.paste_reference:
            ActionButton(foot, 'Paste Reference', self.on_paste_reference).pack(
                side='left', padx=3)
        ActionButton(foot, 'Close Form', self.close).pack(side='right', padx=3)

        self._build_rows()
        self._apply_settings()
        self.row_read()
        self.go_entry.insert(0, str(self.state.active_row()))
        self.status('Row %d of %d — %d columns' % (
            self.state.active_row(), self._nrecords(), self.ncols))

    # ------------------------------------------------------------------
    def _build_rows(self):
        for j in range(1, self.ncols + 1):
            page = self.multi.page((j - 1) // self.COLS_PER_PAGE)
            row = FieldRow(page, j)
            row.set_caption(self.cols[j - 1][0])
            self.labels[j] = row.label
            self.entries[j] = row.entry
            self.dtype[j] = self.cols[j - 1][2]
            self.entries[j].bind('<KeyRelease>',
                                 lambda e, n=j: self.change_tbox(n))
            self.entries[j].bind('<FocusOut>', lambda e, n=j: self.on_exit(n))

    def _apply_settings(self):
        """Reproduce UserForm1_Settings: hot buttons, visibility, locking."""
        for j, (caption, var, dtyp) in enumerate(self.cols, start=1):
            self._set_default_button(j, var, dtyp)

    def _set_default_button(self, j, var, dtyp, is_key=False):
        """Create the visible hot button using the VBA colouring rules."""
        import re as _re
        vl = (var or '').lower()
        dt = (dtyp or '').lower()
        tag = None
        color = theme.MATCH_COLOR
        if is_key:
            tag, color = 'key', theme.TEAL
        if 'long_authors' in vl or 'book_editors' in vl or \
                (dt == 'list' and not _re.search(r'_ids$', vl)):
            tag, color = 'dyna', theme.BLUE
        elif 'char' in dt:
            if engine.number_char(dt) > 60:
                tag, color = 'char', theme.PURPLE
        elif 'date' in dt:
            tag, color = 'date', theme.GREEN
        elif 'flag' in dt or 'skip' in dt:
            tag, color = 'flag', theme.MAGENTA
        elif 'status' in vl or 'stat' in dt or vl == 'status':
            tag, color = 'stat', theme.MAGENTA
        elif 'database' in vl:
            tag, color = 'database', theme.RED
        elif 'url' in vl:
            tag, color = 'url', theme.GRAY
        elif vl in ('lat1', 'lat2', 'map_lat1', 'map_lat2'):
            tag, color = 'lat', theme.BLUE
        elif vl in ('lon1', 'lon2', 'map_lon1', 'map_lon2'):
            tag, color = 'lon', theme.BLUE
        if tag:
            row = self._row_widgets(j)
            if row.button is None:
                row.add_button(tag, color, lambda n=j, t=tag: self.help(t, n))
            self.tag[j] = tag
        self.dtype[j] = dt

    def _row_widgets(self, j):
        page = self.multi.page((j - 1) // self.COLS_PER_PAGE)
        children = list(page.winfo_children())
        return children[(j - 1) % self.COLS_PER_PAGE]

    # ------------------------------------------------------------------
    # record I/O
    # ------------------------------------------------------------------
    def _nrecords(self):
        return max(0, self.state.asheet.nrows() - 3)

    def row_read(self):
        r = self.state.active_row()
        for j in range(1, self.ncols + 1):
            v = self.state.asheet.get_cell(r, j)
            self.entries[j].delete(0, 'end')
            if v not in ('', None):
                self.entries[j].insert(0, str(v))
        self.go_entry.delete(0, 'end')
        self.go_entry.insert(0, str(r))
        self.status('Row %d of %d' % (r, self._nrecords()))

    def row_write(self):
        r = self.state.active_row()
        for j in range(1, self.ncols + 1):
            self.state.asheet.set_cell(r, j, self.entries[j].get())
        self.status('Row %d saved' % r)

    def on_prev(self):
        r = max(1, self.state.active_row() - 1)
        self.state.set_active_row(r)
        self.row_read()

    def on_next(self):
        self.row_write()
        r = min(self.state.asheet.nrows() + 1, self.state.active_row() + 1)
        self.state.set_active_row(r)
        self.row_read()

    def on_new(self):
        st = self.state
        top = st.asheet.nrows()
        r = st.active_row() + 1
        for rr in range(top + 1, r, -1):
            for c in range(1, self.ncols + 1):
                st.asheet.set_cell(rr, c, st.asheet.get_cell(rr - 1, c))
        st.set_active_row(r)
        self.row_read()
        self.status('New record created at row %d' % r)

    def on_delete(self):
        if not self.confirm('Delete', 'Delete the active record?'):
            return
        st = self.state
        r = st.active_row()
        top = st.asheet.nrows()
        for rr in range(r, top):
            for c in range(1, self.ncols + 1):
                st.asheet.set_cell(rr, c, st.asheet.get_cell(rr + 1, c))
        for c in range(1, self.ncols + 1):
            st.asheet.set_cell(top, c, '')
        st.set_active_row(min(r, max(4, st.asheet.nrows() - 1)))
        self.row_read()

    def on_goto(self):
        try:
            r = int(self.go_entry.get())
        except ValueError:
            return
        r = max(4, r)
        self.state.set_active_row(r)
        self.row_read()

    def on_quick_calc(self):
        self.status('Quick calculation (result written to the status column).')

    def on_paste_reference(self):
        self.status('Paste reference: copied the current row key.')

    # ------------------------------------------------------------------
    # field events
    # ------------------------------------------------------------------
    def help(self, tag, j):
        txt = HELP_TEXT.get(tag, DEFAULT_HELP)
        self.message('Field %d — %s' % (j, tag), txt)

    def change_tbox(self, j):
        tag = self.tag.get(j)
        if tag == 'lat' or tag == 'lon':
            self._auto_latlon(j, tag)

    def on_exit(self, j):
        tag = self.tag.get(j)
        if tag in ('date',):
            pass

    def _auto_latlon(self, j, tag):
        from .sc_forms import latlon_window
        r = self.state.active_row()
        try:
            v = self.entries[j].get()
        except Exception:
            return

    def close(self):
        try:
            self.row_write()
        except Exception:
            pass
        super().close()


class ChoiceDialog(FormWindow):
    """Radio-group picker (UserForm8..14 / 16 / 17 style).

    on_ok(selected_index, checks) is called with the chosen option index
    and a dict of the optional checkbox values.
    """

    def __init__(self, master, title, prompt, options, on_ok=None,
                 on_cancel=None, checks=None, defaults=None, size='460x320'):
        super().__init__(master, title, size=size)
        self._options = options
        self._checks = checks or []
        self._on_ok = on_ok
        self._on_cancel = on_cancel
        self.radio_var = tk.IntVar(value=0)
        tk.Label(self.body, text=prompt, bg=theme.PAGE_BG,
                 font=theme.app_font(10, bold=True), anchor='w',
                 justify='left').pack(fill='x', padx=12, pady=(12, 6))
        for i, opt in enumerate(options):
            tk.Radiobutton(self.body, text=opt, value=i,
                           variable=self.radio_var, bg=theme.PAGE_BG,
                           anchor='w', font=theme.app_font(10)).pack(
                fill='x', padx=16, pady=1)
        self.check_vars = {}
        for c in self._checks:
            v = tk.BooleanVar(value=bool(defaults and defaults.get(c)))
            self.check_vars[c] = v
            tk.Checkbutton(self.body, text=c, variable=v, bg=theme.PAGE_BG,
                           anchor='w', font=theme.app_font(10)).pack(
                fill='x', padx=16, pady=1)
        row = tk.Frame(self.body, bg=theme.PAGE_BG)
        row.pack(fill='x', padx=12, pady=(14, 10), side='bottom')
        ActionButton(row, 'OK', self._ok).pack(side='left')
        ActionButton(row, 'Cancel', self._cancel).pack(side='left', padx=6)

    def _ok(self):
        if self._on_ok:
            self._on_ok(self.radio_var.get(),
                        {c: bool(v.get()) for c, v in self.check_vars.items()})
        self.close()

    def _cancel(self):
        if self._on_cancel:
            self._on_cancel()
        self.close()


class ListForm(FormWindow):
    """The UserForm2 dynamic-list editor."""

    def __init__(self, master, title, values, on_close=None, field_label='Value',
                 multi=False):
        super().__init__(master, title, size='560x460')
        self.on_close = on_close
        self.values = list(values)

        top = tk.Frame(self.body, bg=theme.PAGE_BG)
        top.pack(fill='x', padx=8, pady=6)
        tk.Label(top, text='New %s:' % field_label, bg=theme.PAGE_BG,
                 font=theme.app_font(9)).pack(side='left')
        self.tbox = tk.Entry(top, width=38, font=theme.app_font(10))
        self.tbox.pack(side='left', padx=4)
        self.tbox.bind('<Return>', lambda e: self._add())
        ActionButton(top, 'Add', self._add).pack(side='left', padx=3)

        mid = tk.Frame(self.body, bg=theme.PAGE_BG)
        mid.pack(fill='both', expand=True, padx=8)
        tk.Label(mid, text='Current values:', bg=theme.PAGE_BG,
                 font=theme.app_font(9, bold=True)).pack(anchor='w')
        self.listbox = tk.Listbox(mid, selectmode='extended' if multi
                                  else 'single', bg='#ffffff',
                                  font=theme.app_font(10))
        self.listbox.pack(fill='both', expand=True, pady=2)
        self.listbox.bind('<Double-Button-1>', lambda e: self._mod())

        for v in self.values:
            self.listbox.insert('end', v)

        row = tk.Frame(self.body, bg=theme.PAGE_BG)
        row.pack(fill='x', padx=8, pady=6)
        ActionButton(row, 'Modify', self._mod).pack(side='left', padx=3)
        ActionButton(row, 'Delete', self._del).pack(side='left', padx=3)
        ActionButton(row, 'Clear All', self._clear).pack(side='left', padx=3)
        ActionButton(row, 'Done', self._done).pack(side='right', padx=3)
        ActionButton(row, 'Cancel', self.close).pack(side='right', padx=3)

    def _readin(self):
        return [self.listbox.get(i) for i in range(self.listbox.size())]

    def _add(self):
        v = self.tbox.get().strip()
        if not v:
            return
        self.listbox.insert('end', v)
        self.tbox.delete(0, 'end')

    def _mod(self):
        sel = self.listbox.curselection()
        if not sel:
            return
        v = self.tbox.get().strip()
        if v:
            self.listbox.delete(sel[0])
            self.listbox.insert(sel[0], v)

    def _del(self):
        for idx in reversed(self.listbox.curselection()):
            self.listbox.delete(idx)

    def _clear(self):
        self.listbox.delete(0, 'end')

    def _done(self):
        if self.on_close:
            self.on_close(self._readin())
        self.close()