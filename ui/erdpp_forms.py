# -*- coding: utf-8 -*-
"""
ERDPP.console rebuilt user-forms.

    UserForm1  ErdpRecordForm  -- data editor (same engine as SC)
    UserForm2  ListForm        -- dynamic list editor (shared)
    UserForm3  ErdpUploadForm  -- citation-file uploader
"""
import tkinter as tk

from . import engine, theme
from .framework import ActionButton, FormWindow
from .record_form import RecordFormBase

FILE_KINDS = [
    ('Abstract (.abs)', '.abs'),
    ('Text file (.txt)', '.txt'),
    ('Table (.tab)', '.tab'),
    ('Tech copy (.tch)', '.tch'),
    ('Appendix (.app)', '.app'),
    ('Document (.doc)', '.doc'),
    ('Figure (.fig)', '.fig'),
    ('Temporary (.tmp)', '.tmp'),
]


class ErdpRecordForm(RecordFormBase):
    def __init__(self, master, state, **kw):
        cols = [(c[0], c[1], c[2]) for c in engine.ERDPP_COLUMNS]
        kw.setdefault('quick_calc', True)
        kw.setdefault('paste_reference', True)
        super().__init__(master, 'ERDPP.console — Data Entry', state, cols,
                         subtitle='Earth REFERENCE Database Project Parser',
                         **kw)


class ErdpUploadForm(FormWindow):
    """UserForm3: make a folder/file name and upload a citation file."""

    def __init__(self, master, state=None, on_upload=None):
        super().__init__(master, 'ERDPP — Citation File Uploader',
                         size='580x440')
        self.state = state
        self.on_upload = on_upload
        body = self.body

        top = tk.Frame(body, bg=theme.PAGE_BG)
        top.pack(fill='x', padx=12, pady=(10, 2))
        tk.Label(top, text='Citation :', bg=theme.PAGE_BG,
                 font=theme.app_font(10)).pack(side='left')
        self.open_btn = ActionButton(top, 'Open ...', self.open_file)
        self.open_btn.pack(side='right')
        self.tb_cit = tk.Entry(top, width=40, bg=theme.FIELD_BG,
                               relief='sunken', bd=2, font=theme.app_font(9))
        self.tb_cit.pack(side='left', padx=6, fill='x', expand=True)
        tk.Label(top, text='Citation KB :', bg=theme.PAGE_BG,
                 font=theme.app_font(9)).pack(anchor='w', padx=12)

        mid = tk.Frame(body, bg=theme.PAGE_BG)
        mid.pack(fill='x', padx=12, pady=4)
        tk.Label(mid, text='Copy content as :', bg=theme.PAGE_BG,
                 font=theme.app_font(9, bold=True)).pack(anchor='w')
        self.kind = tk.StringVar(value=FILE_KINDS[0][1])
        for text, ext in FILE_KINDS:
            tk.Radiobutton(mid, text=text, value=ext, variable=self.kind,
                           bg=theme.PAGE_BG, font=theme.app_font(9)).pack(
                anchor='w', padx=20)

        nm = tk.Frame(body, bg=theme.PAGE_BG)
        nm.pack(fill='x', padx=12, pady=6)
        tk.Label(nm, text='File name :', bg=theme.PAGE_BG,
                 font=theme.app_font(9)).grid(row=0, column=0, sticky='e')
        self.tb_name = tk.Entry(nm, width=34, font=theme.app_font(9))
        self.tb_name.grid(row=0, column=1, padx=6, pady=2)
        self.make_name = ActionButton(nm, 'Make', self.make_name_click)
        self.make_name.grid(row=0, column=2, padx=6)
        self.make_folder = ActionButton(nm, 'Make Folder', self.make_folder_click)
        self.make_folder.grid(row=0, column=3, padx=6)
        tk.Label(nm, text='Serial # :', bg=theme.PAGE_BG,
                 font=theme.app_font(9)).grid(row=1, column=0, sticky='e')
        self.sb_num = tk.Spinbox(nm, from_=1, to=9999, width=8,
                                 font=theme.app_font(9))
        self.sb_num.grid(row=1, column=1, sticky='w', padx=6, pady=2)
        self.sb_num.delete(0, 'end')
        self.sb_num.insert(0, '1')

        tk.Label(body, text='File content (paste the citation text):',
                 bg=theme.PAGE_BG, font=theme.app_font(9, bold=True)).pack(
            anchor='w', padx=12, pady=(4, 0))
        self.txt = tk.Text(body, height=8, bg='#ffffff', relief='sunken',
                           bd=2, font=theme.app_font(9))
        self.txt.pack(fill='both', expand=True, padx=12, pady=4)

        row = tk.Frame(body, bg=theme.PAGE_BG)
        row.pack(fill='x', padx=12, pady=6, side='bottom')
        ActionButton(row, 'Upload', self.upload).pack(side='left')
        ActionButton(row, 'Close', self.close).pack(side='right')

    def open_file(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title='Choose a citation file',
            filetypes=[('Text files', '*.txt *.abs *.tab'), ('All files', '*.*')])
        if path:
            self.tb_cit.delete(0, 'end')
            self.tb_cit.insert(0, path)

    def _base_name(self):
        base = self.tb_cit.get().strip()
        if not base:
            base = self.tb_name.get().strip() or 'citation'
        return base

    def make_name_click(self):
        base = self._base_name()
        name = engine.custom_file_name(base) + self.kind.get()
        self.tb_name.delete(0, 'end')
        self.tb_name.insert(0, name)
        self.status('File name made: %s' % name)

    def make_folder_click(self):
        self.status('Folder name: %s' %
                    engine.custom_file_name(self._base_name()))

    def upload(self):
        content = self.txt.get('1.0', 'end-1c').strip()
        name = self.tb_name.get().strip()
        if self.on_upload:
            self.on_upload({'name': name, 'kind': self.kind.get(),
                            'serial': self.sb_num.get(),
                            'content': content})
        self.status('Uploaded "%s" as %s' % (name or self._base_name(),
                                             self.kind.get()))