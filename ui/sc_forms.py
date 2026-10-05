# -*- coding: utf-8 -*-
"""
Smart Coast (SC.console) rebuilt user-forms.

    UserForm1  ScRecordForm      -- main data-entry editor
    UserForm2  ListForm          -- dynamic list editor (shared)
    UserForm5  LatLonForm        -- lat/lon decimal recast dialog
    UserForm7  BatchUpdateForm   -- batch number-insertion dialog
    UserForm8  RecordStatusDialog -- record status picker
    UserForm9  IndexStatusDialog  -- index status picker
    UserForm10 CreationStatusDialog -- creation status picker
    UserForm11 CalculationStatusDialog -- calculation status picker
    UserForm12 DeleteStatusDialog -- delete status picker
    UserForm13 MapStatusDialog    -- map status + checkbox flags
    UserForm14 AutoMapUpdateDialog -- auto map-update picker
    UserForm15 MapSetupForm       -- map setup / option matrix editor
    UserForm16 LocaStatusDialog   -- localisation status + overwrite flag
    UserForm17 CustomGridStatusDialog -- custom grid picker
"""
import re

import tkinter as tk

from . import engine, theme
from .framework import ActionButton, FieldRow, FormWindow, HotButton, select_text
from .record_form import RecordFormBase, ChoiceDialog

# ---------------------------------------------------------------------------
# UserForm1 -- the SC record editor
# ---------------------------------------------------------------------------

SC_STATUS_HELP = ('Records leaving this screen are marked with a status: '
                  '"Temporary" while they are being worked on, "Frozen" once '
                  'the record is ready for QC, and various database states '
                  'after that.')


class ScRecordForm(RecordFormBase):
    def __init__(self, master, state, **kw):
        cols = [(c[0], c[1], c[2]) for c in engine.SC_COLUMNS]
        kw.setdefault('quick_calc', True)
        kw.setdefault('paste_reference', True)
        kw.setdefault('master_user', True)
        super().__init__(master, 'SC.console — Data Entry', state, cols,
                         subtitle='Smart Coast Data Manager', **kw)


# ---------------------------------------------------------------------------
# UserForm5 -- lat / lon decimal recast dialog
# ---------------------------------------------------------------------------
class LatLonForm(FormWindow):
    def __init__(self, master, callback=None):
        super().__init__(master, 'Latitude / Longitude Recast', size='520x360')
        self.callback = callback
        body = self.body
        tk.Label(body, text='Enter decimal lat/lon below, then choose the '
                 'recast action.', bg=theme.PAGE_BG, fg='#222222',
                 font=theme.app_font(9, bold=True), anchor='w').pack(
            fill='x', padx=12, pady=(10, 4))

        grid = tk.Frame(body, bg=theme.PAGE_BG)
        grid.pack(fill='x', padx=16, pady=4)
        tk.Label(grid, text='Latitude :', bg=theme.PAGE_BG,
                 font=theme.app_font(10)).grid(row=0, column=0, sticky='e')
        self.lat0 = tk.Entry(grid, width=16, font=theme.app_font(10))
        self.lat0.grid(row=0, column=1, padx=6, pady=2)
        ActionButton(grid, 'Start Calculation Lat', self.calc_lat).grid(
            row=0, column=2, padx=6)
        ActionButton(grid, 'Copy Value', lambda: self._copy('lat')).grid(
            row=0, column=3, padx=6)
        tk.Label(grid, text='Longitude:', bg=theme.PAGE_BG,
                 font=theme.app_font(10)).grid(row=1, column=0, sticky='e')
        self.lon0 = tk.Entry(grid, width=16, font=theme.app_font(10))
        self.lon0.grid(row=1, column=1, padx=6, pady=2)
        ActionButton(grid, 'Start Calculation Lon', self.calc_lon).grid(
            row=1, column=2, padx=6)
        ActionButton(grid, 'Copy Value', lambda: self._copy('lon')).grid(
            row=1, column=3, padx=6)

        tk.Label(body, text='Recast :', bg=theme.PAGE_BG,
                 font=theme.app_font(10, bold=True)).pack(anchor='w',
                                                          padx=16, pady=(8, 0))
        self.outcome = tk.Text(body, height=2, width=56, bg='#ffffff',
                               relief='sunken', bd=2, font=theme.app_font(10))
        self.outcome.pack(fill='both', expand=True, padx=16, pady=4)
        row = tk.Frame(body, bg=theme.PAGE_BG)
        row.pack(fill='x', padx=16, pady=6)
        ActionButton(row, 'Done', self.close).pack(side='right')

    def _eprint(self, txt):
        self.outcome.delete('0.0', 'end')
        self.outcome.insert('0.0', txt)

    def calc_lat(self):
        self._eprint('  %.6f  ->  %s' % (
            engine.latlon2dec(self.lat0.get()) if self.lat0.get().strip() else 0.0,
            engine.recast_dec_lat(self.lat0.get()) if self.lat0.get().strip() else ''))

    def calc_lon(self):
        self._eprint('  %.6f  ->  %s' % (
            engine.latlon2dec(self.lon0.get()) if self.lon0.get().strip() else 0.0,
            engine.recast_dec_lon(self.lon0.get()) if self.lon0.get().strip() else ''))

    def _copy(self, which):
        e = self.lat0 if which == 'lat' else self.lon0
        select_text(e)


def latlon_window(master, callback=None):
    return LatLonForm(master, callback)


# ---------------------------------------------------------------------------
# UserForm7 -- batch number insertion with progress bar
# ---------------------------------------------------------------------------
class BatchUpdateForm(FormWindow):
    def __init__(self, master, state, on_done=None):
        super().__init__(master, 'Batch Update Rows', size='540x380')
        self.state = state
        self.on_done = on_done
        body = self.body
        tk.Label(body, text='Apply a number to every data row in the active '
                 'table.', bg=theme.PAGE_BG, font=theme.app_font(10, bold=True),
                 anchor='w').pack(fill='x', padx=14, pady=(10, 4))

        top = tk.Frame(body, bg=theme.PAGE_BG)
        top.pack(fill='x', padx=16)
        tk.Label(top, text='Number :', bg=theme.PAGE_BG,
                 font=theme.app_font(10)).pack(side='left')
        self.number = tk.Entry(top, width=12, font=theme.app_font(10))
        self.number.pack(side='left', padx=6)
        tk.Label(top, text='Column :', bg=theme.PAGE_BG,
                 font=theme.app_font(10)).pack(side='left', padx=(16, 2))
        self.column = tk.Entry(top, width=6, font=theme.app_font(10))
        self.column.insert(0, '1')
        self.column.pack(side='left', padx=2)

        tk.Label(body, text='Insert mode :', bg=theme.PAGE_BG,
                 font=theme.app_font(9, bold=True)).pack(anchor='w', padx=16,
                                                         pady=(8, 0))
        self.mode = tk.StringVar(value='front')
        for text, val in (('Front', 'front'), ('Both sides', 'both'),
                          ('Behind', 'after'), ('Overwrite rows',
                                                'overwrite'), ('Empty rows only',
                                                               'empty')):
            tk.Radiobutton(body, text=text, value=val, variable=self.mode,
                           bg=theme.PAGE_BG, font=theme.app_font(9)).pack(
                anchor='w', padx=24)

        self.pbar = tk.Frame(body, bg=theme.PAGE_BG)
        self.pbar.pack(fill='x', padx=16, pady=(10, 2))
        import tkinter.ttk as ttk
        self.progress = ttk.Progressbar(self.pbar, length=420, maximum=100,
                                        mode='determinate')
        self.progress.pack(fill='x')
        self.progress_label = tk.Label(body, text='', bg=theme.PAGE_BG,
                                       font=theme.app_font(9))
        self.progress_label.pack(anchor='w', padx=16)

        row = tk.Frame(body, bg=theme.PAGE_BG)
        row.pack(fill='x', padx=16, pady=(8, 10), side='bottom')
        ActionButton(row, 'Start Batch Update', self.start).pack(side='left')
        ActionButton(row, 'Cancel', self.close).pack(side='right')

    def start(self):
        st = self.state
        try:
            number = float(self.number.get() or 0)
            col = max(1, int(self.column.get() or 1))
        except ValueError:
            self.message('Batch Update', 'Enter a valid number and column.',
                         'error')
            return
        mode = self.mode.get()
        start_row, end_row = 4, st.asheet.nrows()
        total = max(1, end_row - start_row + 1)
        for i, r in enumerate(range(start_row, end_row + 1)):
            old = st.asheet.get_cell(r, col)
            shown = str(number)
            if mode == 'front':
                new = shown if engine.is_empty(old) else shown + str(old)
            elif mode == 'after':
                new = shown if engine.is_empty(old) else str(old) + shown
            elif mode == 'both':
                new = str(old) + shown if not engine.is_empty(old) else shown
            elif mode == 'overwrite':
                new = shown
            else:  # empty
                new = shown if engine.is_empty(old) else old
            st.asheet.set_cell(r, col, new)
            self.progress['value'] = 100.0 * (i + 1) / total
            self.progress_label.configure(text='Row %d processed' % r)
            self.update_idletasks()
        self.progress_label.configure(
            text='Done — %d rows updated.' % (end_row - start_row + 1))
        if self.on_done:
            self.on_done()


# ---------------------------------------------------------------------------
# UserForm8..17 -- small choice dialogs
# ---------------------------------------------------------------------------
def record_status_dialog(master, state, on_ok=None):
    d = ChoiceDialog(master, 'Record Status', 'Set the record status',
                     ['Status 0', 'Status 1', 'Status 2', 'Status 3',
                      'Status 4', 'Status 5'], on_ok=on_ok)
    return d


def index_status_dialog(master, state, on_ok=None):
    return ChoiceDialog(master, 'Index Status', 'Set the index status',
                        ['Temporary (work in progress)', 'Frozen (ready for QC)'],
                        on_ok=on_ok)


def creation_status_dialog(master, state, on_ok=None):
    return ChoiceDialog(master, 'Creation Status',
                        'Set the creation status for the data copy',
                        ['New (to be built)', 'Existing (keep as is)'],
                        on_ok=on_ok)


def calculation_status_dialog(master, state, on_ok=None):
    return ChoiceDialog(master, 'Calculation Status',
                        'Choose the calculation scope',
                        ['Selected rows only', 'Entire table'], on_ok=on_ok)


def delete_status_dialog(master, state, on_ok=None):
    return ChoiceDialog(master, 'Delete Status',
                        'How should this record be removed?',
                        ['Delete entirely', 'Mark as deleted only'],
                        on_ok=on_ok)


def map_status_dialog(master, state, on_ok=None):
    return ChoiceDialog(master, 'Map Status',
                        'Which maps should be (re)built?',
                        ['All maps', 'Only selected schools'], on_ok=on_ok,
                        checks=['Overwrite settings', 'Back scatter generation',
                                'Gravity maps generation'])


def auto_map_update_dialog(master, state, on_ok=None):
    return ChoiceDialog(master, 'Auto Map Update',
                        'Request an automatic update check?',
                        ['Yes, update now', 'No, skip for this session'],
                        on_ok=on_ok)


def loca_status_dialog(master, state, on_ok=None):
    return ChoiceDialog(master, 'Localisation Status',
                        'Set the localisation status',
                        ['Localise all active rows', 'Localise selected rows only'],
                        on_ok=on_ok, checks=['Overwrite existing settings'])


def custom_grid_status_dialog(master, state, on_ok=None):
    return ChoiceDialog(master, 'Custom Grid Status',
                        'Choose the custom grid intensity',
                        ['Coarse (fast)', 'Medium', 'Fine (slow)'], on_ok=on_ok)


# ---------------------------------------------------------------------------
# UserForm15 -- map setup editor
# ---------------------------------------------------------------------------
MAP_FIELDS = [
    ('Grid Size (minutes)', 'grid_size_min', 'grid_size_min'),
    ('Map Scale', 'gmt_scale', 'GMT_SCALE'),
    ('Sun', 'gmt_sun', 'GMT_SUN'),
    ('Normalization', 'gmt_normalization', 'GMT_NORMALIZATION'),
    ('Minor Contour Interval', 'gmt_ci1', 'GMT_CI1'),
    ('Major Contour Interval', 'gmt_ci2', 'GMT_CI2'),
    ('Grid Line Tick', 'gridline_interval', 'GRIDLINE_INTERVAL'),
    ('Label Interval', 'label_interval', 'LABEL_INTERVAL'),
    ('Frame Interval', 'frame_interval', 'FRAME_INTERVAL'),
    ('Map Caption', 'file_caption', 'FILE_CAPTION'),
    ('Seamount Name', 'seamount_name', 'SEAMOUNT_NAME'),
    ('Region Name', 'region_name', 'REGION_NAME'),
    ('Ocean Name', 'ocean_name', 'OCEAN_NAME'),
    ('Map File Name', 'file_name', 'FILE_NAME'),
]
BOUND_FIELDS = [('North limit', 'map_lat2', 'MAP_LAT2'),
                ('South limit', 'map_lat1', 'MAP_LAT1'),
                ('West limit', 'map_lon2', 'MAP_LON2'),
                ('East limit', 'map_lon1', 'MAP_LON1')]
MAP_KINDS = ['SSGRID', 'SSCPT', 'SSASC', 'SSMBGRID', 'SSTACK']


class MapSetupForm(FormWindow):
    def __init__(self, master, state, on_update=None):
        super().__init__(master, 'SC Map Setup', size='760x600')
        self.state = state
        self.on_update = on_update
        self.maps = state.awb.Sheets('SC_maps')

        # locate the data row and header positions
        self.row = 4
        headers = {}
        for c in range(1, self.maps.ncols() + 1):
            v = self.maps.get_cell(1, c)
            if v:
                headers[str(v).upper()] = c
        self.headers = headers

        body = self.body
        left = tk.Frame(body, bg=theme.PAGE_BG)
        left.pack(side='left', fill='both', expand=True, padx=4)
        right = tk.Frame(body, bg=theme.PAGE_BG)
        right.pack(side='right', fill='both', expand=True, padx=4)

        self.field_entries = {}
        g = tk.Frame(left, bg=theme.PAGE_BG)
        g.pack(fill='x')
        for i, (lab, var, _h) in enumerate(MAP_FIELDS):
            tk.Label(g, text=lab + ':', bg=theme.PAGE_BG,
                     font=theme.app_font(9), anchor='e').grid(
                row=i, column=0, sticky='e', padx=(4, 4), pady=1)
            e = tk.Entry(g, width=30, bg=theme.FIELD_BG, relief='sunken',
                         bd=2, font=theme.app_font(9))
            e.grid(row=i, column=1, sticky='w', pady=1)
            self.field_entries[var] = e
        tk.Label(g, text='Limits:', bg=theme.PAGE_BG,
                 font=theme.app_font(9, bold=True)).grid(
            row=len(MAP_FIELDS), column=0, columnspan=2, sticky='w',
            padx=4, pady=(8, 2))
        for i, (lab, var, _h) in enumerate(BOUND_FIELDS, start=len(MAP_FIELDS) + 1):
            tk.Label(g, text=lab + ':', bg=theme.PAGE_BG,
                     font=theme.app_font(9), anchor='e').grid(
                row=i, column=0, sticky='e', padx=(4, 4), pady=1)
            e = tk.Entry(g, width=30, bg=theme.FIELD_BG, relief='sunken',
                         bd=2, font=theme.app_font(9))
            e.grid(row=i, column=1, sticky='w', pady=1)
            self.field_entries['bound_' + var] = e

        tk.Label(right, text='Map type:', bg=theme.PAGE_BG,
                 font=theme.app_font(9, bold=True)).pack(anchor='w', padx=6,
                                                         pady=(4, 0))
        self.mapkind = tk.StringVar()
        for k in MAP_KINDS:
            tk.Radiobutton(right, text=k, value=k, variable=self.mapkind,
                           bg=theme.PAGE_BG, font=theme.app_font(9)).pack(
                anchor='w', padx=18)
        self.autocpt = tk.BooleanVar(value=True)
        tk.Checkbutton(right, text='Auto CPT generation', variable=self.autocpt,
                       bg=theme.PAGE_BG, font=theme.app_font(9)).pack(
            anchor='w', padx=18, pady=(6, 0))

        tk.Label(right, text='Map options:', bg=theme.PAGE_BG,
                 font=theme.app_font(9, bold=True)).pack(anchor='w', padx=6,
                                                         pady=(8, 0))
        self.opt_vars = []
        om = tk.Frame(right, bg=theme.PAGE_BG)
        om.pack(fill='both', expand=True, padx=6)
        for i, name in enumerate(engine.MAP_OPTIONS):
            v = tk.BooleanVar(value=False)
            tk.Checkbutton(om, text=name, variable=v, bg=theme.PAGE_BG,
                           font=theme.app_font(8)).pack(anchor='w')
            self.opt_vars.append(v)

        foot = tk.Frame(body, bg=theme.PAGE_BG)
        foot.pack(fill='x', side='bottom', pady=6)
        ActionButton(foot, 'Rewrite Map Setup', self.update_row).pack(side='left', padx=6)
        ActionButton(foot, 'New Map Setup', self.new_row).pack(side='left', padx=6)
        ActionButton(foot, 'Close', self.close).pack(side='right', padx=6)

        self.read_row()

    # -------- I/O against the SC_maps sheet ----------------
    def _cell(self, header, default=''):
        c = self.headers.get(header)
        if not c:
            return default
        return self.maps.get_cell(self.row, c)

    def _set_cell(self, header, value):
        c = self.headers.get(header)
        if c:
            self.maps.set_cell(self.row, c, value)

    def read_row(self):
        for lab, var, h in MAP_FIELDS:
            self.field_entries[var].delete(0, 'end')
            self.field_entries[var].insert(0, self._cell(h))
        for lab, var, h in BOUND_FIELDS:
            self.field_entries['bound_' + var].delete(0, 'end')
            self.field_entries['bound_' + var].insert(0, self._cell(h))
        kind = self._cell('CREATION_TYPE')
        if kind in MAP_KINDS:
            self.mapkind.set(kind)
        opts = str(self._cell('MAP_OPTIONS', ''))
        for i, v in enumerate(self.opt_vars):
            v.set(i < len(opts) and opts[i] == 'Y')

    def update_row(self):
        for lab, var, h in MAP_FIELDS:
            self._set_cell(h, self.field_entries[var].get())
        for lab, var, h in BOUND_FIELDS:
            self._set_cell(h, self.field_entries['bound_' + var].get())
        self._set_cell('CREATION_TYPE', self.mapkind.get())
        self._set_cell('MAP_OPTIONS',
                       ''.join('Y' if v.get() else 'N' for v in self.opt_vars))
        if self.on_update:
            self.on_update()
        self.status('Map setup row %d saved.' % self.row)

    def new_row(self):
        for e in self.field_entries.values():
            e.delete(0, 'end')
        for v in self.opt_vars:
            v.set(False)
        self.mapkind.set(MAP_KINDS[0])
        self.status('Cleared — press "Rewrite Map Setup" to save as row %d.'
                    % (self.row + 1))