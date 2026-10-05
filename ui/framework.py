# -*- coding: utf-8 -*-
"""
Base tkinter widgets used by all rebuilt user-forms.

The legacy forms were plain VBA UserForms; these classes give the same
feel (blocky coloured command buttons, labelled data rows, page tabs) using
only the standard library.
"""
import tkinter as tk
from tkinter import ttk, messagebox

from . import theme


def rgb(*args):
    """vba.rgb(r, g, b) equivalent returning a tkinter hex colour."""
    r, g, b = args[:3]
    return '#%02x%02x%02x' % (int(r) & 255, int(g) & 255, int(b) & 255)


class FormWindow(tk.Toplevel):
    """Base window: header strip + themed background + close-on-Escape."""

    def __init__(self, master, title, subtitle=None, size='740x560'):
        super().__init__(master)
        self.title(title)
        self.configure(bg=theme.PAGE_BG)
        self.geometry(size)
        self.resizable(True, True)
        self.protocol('WM_DELETE_WINDOW', self.close)
        header = tk.Frame(self, bg=theme.HEADER_BG)
        header.pack(fill='x')
        tk.Label(header, text=title, bg=theme.HEADER_BG, fg=theme.HEADER_FG,
                 font=theme.app_font(12, bold=True), anchor='w').pack(
            side='left', padx=10, pady=6)
        if subtitle:
            tk.Label(header, text=subtitle, bg=theme.HEADER_BG,
                     fg=theme.HEADER_FG, font=theme.app_font(9), anchor='e'
                     ).pack(side='right', padx=10, pady=6)
        self.body = tk.Frame(self, bg=theme.PAGE_BG)
        self.body.pack(fill='both', expand=True, padx=6, pady=6)
        self.bind('<Escape>', lambda e: self.close())
        self._sub_windows = []

    def child(self, cls, *a, **k):
        w = cls(self, *a, **k)
        self._sub_windows.append(w)
        return w

    def status(self, text):
        if getattr(self, '_status_var', None) is not None:
            self._status_var.set(str(text))

    def message(self, title, text, kind='info'):
        if kind == 'error':
            messagebox.showerror(title, text, parent=self)
        elif kind == 'warning':
            messagebox.showwarning(title, text, parent=self)
        else:
            messagebox.showinfo(title, text, parent=self)

    def confirm(self, title, text):
        return messagebox.askyesno(title, text, parent=self)

    def close(self):
        for w in list(self._sub_windows):
            if w.winfo_exists():
                w.destroy()
        if self.winfo_exists():
            self.destroy()


class HotButton(tk.Button):
    """Coloured command button used beside each data column."""

    def __init__(self, parent, text, color, command, width=9):
        super().__init__(parent, text=text, command=command, width=width,
                         relief='raised', bd=2, cursor='hand2')
        self.set_color(color)

    def set_color(self, color):
        self.configure(bg=color, activebackground=color)
        self['fg'] = '#ffffff'


class ActionButton(tk.Button):
    def __init__(self, parent, text, command):
        super().__init__(parent, text=text, command=command, relief='raised',
                         bd=2, bg=theme.BUTTON_BG, fg=theme.BUTTON_FG,
                         activebackground=theme.BUTTON_BG,
                         activeforeground='#ffffff', padx=10,
                         font=theme.app_font(9, bold=False), cursor='hand2')


class FieldRow(tk.Frame):
    """One data-column row: caption label + value entry + optional hot button."""

    def __init__(self, parent, j, read_only=False):
        super().__init__(parent, bg=theme.PAGE_BG_ALT
                         if j % 2 else theme.PAGE_BG)
        self.pack(fill='x', pady=1)
        self.label = tk.Label(self, text='', width=26, anchor='w', bg=self['bg'],
                              font=theme.app_font(9, bold=True))
        self.label.pack(side='left', padx=(8, 4), pady=2)
        self.entry = tk.Entry(self, width=34, bg=theme.FIELD_BG,
                              relief='sunken', bd=2, justify='left',
                              font=theme.app_font(10))
        self.entry.pack(side='left', padx=4, pady=2, fill='x', expand=True)
        if read_only:
            self.entry.configure(state='readonly', readonlybackground=theme.FIELD_DISABLED_BG)
        self.button = None

    def set_caption(self, text):
        self.label.configure(text=str(text))

    def set_value(self, text):
        if text is None:
            return
        self.entry.delete(0, 'end')
        self.entry.insert(0, '' if str(text) == 'None' else str(text))

    def get_value(self):
        return self.entry.get()

    def add_button(self, text, color, command, width=9):
        self.button = HotButton(self, text, color, command, width=width)
        self.button.pack(side='left', padx=4, pady=2)

    def remove_button(self):
        if self.button is not None:
            self.button.destroy()
            self.button = None

    def set_enabled(self, enabled):
        state = 'normal' if enabled else 'disabled'
        self.entry.configure(state=state)

    def bind_key(self, seq, fn):
        self.entry.bind(seq, fn)


class MultiPage(ttk.Notebook):
    """Tab widget; index 0 is the first page (VBA MultiPage also starts at 0)."""

    def __init__(self, parent, pages=8):
        super().__init__(parent)
        self._frames = []
        for p in range(pages):
            f = tk.Frame(self, bg=theme.PAGE_BG)
            self.add(f, text='Page %d' % (p + 1))
            self._frames.append(f)

    def page(self, idx):
        return self._frames[idx]

    def current(self):
        return self.index(self.select())

    def set_page(self, idx):
        self.select(self._frames[idx])


def select_text(entry):
    entry.focus_set()
    entry.selection_range(0, 'end')