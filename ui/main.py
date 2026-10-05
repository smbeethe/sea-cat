# -*- coding: utf-8 -*-
"""
Sea-Cat legacy workbook launcher.

Run from the repository root with:

    python -m ui.main
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from . import engine, theme
from .framework import ActionButton, FormWindow
from .sc_forms import (BatchUpdateForm, LatLonForm, MapSetupForm,
                       ScRecordForm,
                       auto_map_update_dialog, calculation_status_dialog,
                       creation_status_dialog, custom_grid_status_dialog,
                       delete_status_dialog, index_status_dialog,
                       loca_status_dialog, map_status_dialog,
                       record_status_dialog)
from .record_form import ListForm
from .erdpp_forms import ErdpRecordForm, ErdpUploadForm

DESCRIPTION = (
    'Privacy Notice: this tool reads and writes only the sample table that is'
    '\nstored in memory by this launcher. It does not contact Excel, the '
    '\nsea-cat files, or the network. Use "Export table" to save a CSV copy '
    '\nof your edits on your own machine.'
)


class Launcher(FormWindow):
    def __init__(self, root):
        super().__init__(root, 'Sea-Cat Legacy Workbook Launcher',
                         subtitle='Anthony A.P. Koppers - OSU', size='720x560')
        self.state = engine.build_sample_state()
        self.subforms = []

        tk.Label(self.body, text=DESCRIPTION, bg=theme.PAGE_BG, fg='#333333',
                 justify='left', font=theme.app_font(8)).pack(fill='x',
                                                              padx=10, pady=4)

        # ---------- SC.console ----------
        sc = tk.LabelFrame(self.body, text=' SC.console ', bg=theme.PAGE_BG,
                           fg=theme.HEADER_BG, font=theme.app_font(10, bold=True),
                           padx=6, pady=4)
        sc.pack(fill='x', padx=10, pady=4)
        for text, cmd, prompt in (
                ('Open Data Entry editor (UserForm1)', self.sc_record,
                 'Smart Coast data editor'),):
            ActionButton(sc, text, cmd).pack(anchor='w', pady=1)
        row = tk.Frame(sc, bg=theme.PAGE_BG)
        row.pack(anchor='w')
        ActionButton(row, 'Record status ...', self.sc_status).pack(side='left', padx=2)
        ActionButton(row, 'Index status ...', lambda: self._choice(
            index_status_dialog, 'Temporary index marked.')).pack(side='left', padx=2)
        ActionButton(row, 'Map status ...', self._map_status).pack(side='left', padx=2)
        ActionButton(row, 'Batch Update ...', self.sc_batch).pack(side='left', padx=2)

        # ---------- ERDPP.console ----------
        er = tk.LabelFrame(self.body, text=' ERDPP.console ', bg=theme.PAGE_BG,
                           fg=theme.HEADER_BG, font=theme.app_font(10, bold=True),
                           padx=6, pady=4)
        er.pack(fill='x', padx=10, pady=4)
        ActionButton(er, 'Open Data Entry editor (UserForm1)',
                     self.erdp_record).pack(anchor='w', pady=1)
        ActionButton(er, 'Citation manager (UserForm2)', self.erdp_list
                     ).pack(anchor='w', pady=1)
        ActionButton(er, 'Citation Uploader (UserForm3)', self.erdp_upload
                     ).pack(anchor='w', pady=1)

        # ---------- common tools ----------
        tools = tk.LabelFrame(self.body, text=' Tools ', bg=theme.PAGE_BG,
                              fg=theme.HEADER_BG,
                              font=theme.app_font(10, bold=True), padx=6, pady=4)
        tools.pack(fill='x', padx=10, pady=4)
        ActionButton(tools, 'Lat / Lon recast (UserForm5)',
                     self.latlon).pack(anchor='w', pady=1)
        ActionButton(tools, 'Map Setup (UserForm15)', self.map_setup
                     ).pack(anchor='w', pady=1)
        ActionButton(tools, 'List editor', self.list_editor).pack(anchor='w',
                                                                  pady=1)

        foot = tk.Frame(self.body, bg=theme.PAGE_BG)
        foot.pack(fill='x', pady=6, side='bottom')
        ActionButton(foot, 'Export table (CSV)', self.export_csv).pack(side='left', padx=4)
        ActionButton(foot, 'Import table (CSV)', self.import_csv).pack(side='left', padx=4)
        ActionButton(foot, 'Reset sample data', self.reset_data).pack(side='left', padx=4)
        ActionButton(foot, 'Exit', self.close).pack(side='right', padx=6)

    def _new(self, cls, *a, **k):
        f = cls(self, self.state, *a, **k)
        self.subforms.append(f)
        return f

    # ---- SC forms ----
    def sc_record(self):
        self._new(ScRecordForm)

    def sc_batch(self):
        self._new(BatchUpdateForm).transient(self)

    def sc_status(self):
        def ok(i, checks):
            self.status('Record status set to %d.' % i)
        self._new(record_status_dialog, self.state, on_ok=ok) \
            .transient(self)

    def _choice(self, factory, msg):
        def ok(i, checks):
            self.status(msg)
        self._new(factory, self.state, on_ok=ok).transient(self)

    def _map_status(self):
        def ok(i, checks):
            self.status('Map status: option %d, flags=%s' % (
                i, {k: v for k, v in checks.items()}))
        self._new(map_status_dialog, self.state, on_ok=ok).transient(self)

    # ---- ERDPP forms ----
    def erdp_record(self):
        self._new(ErdpRecordForm)

    def erdp_list(self):
        values = ["Smith et al.", "Jones and Lee", "O'Hara B", "Koppers et al."]
        self._new(ListForm, 'ERDPP — Citation Key List', values,
                  on_close=lambda v: self.status(
                      'List updated: %d entries.' % len(v)),
                  field_label='Citation').transient(self)

    def erdp_upload(self):
        self._new(ErdpUploadForm).transient(self)

    # ---- tools ----
    def latlon(self):
        self._new(LatLonForm).transient(self)

    def map_setup(self):
        self._new(MapSetupForm).transient(self)

    def list_editor(self):
        values = ['SSGRID', 'SSCPT', 'SSASC', 'SSMBGRID', 'SSTACK']
        self._new(ListForm, 'Dynamic List Editor', values,
                  on_close=lambda v: self.status('List: %s' % (v,))).transient(self)

    # ---- data utilities ----
    def export_csv(self):
        path = filedialog.asksaveasfilename(defaultextension='.csv',
                                            filetypes=[('CSV files', '*.csv')])
        if not path:
            return
        engine.save_csv_sheet(self.state.asheet, path)
        self.status('Exported active table to %s' % path)

    def import_csv(self):
        path = filedialog.askopenfilename(filetypes=[('CSV files', '*.csv')])
        if not path:
            return
        engine.load_csv_sheet(self.state.asheet, path)
        self.status('Imported table from %s' % path)

    def reset_data(self):
        self.state = engine.build_sample_state()
        self.status('Sample data reset.')

    def close(self):
        for f in list(self.subforms):
            if f.winfo_exists():
                f.destroy()
        super().close()


def main():
    root = tk.Tk()
    root.withdraw()
    Launcher(root)
    root.mainloop()


if __name__ == '__main__':
    main()