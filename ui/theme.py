# The Sea-Cat legacy workbooks (SC.console, ERDPP.console) were Excel 2000-2003
# data-entry tools by Anthony A.P. Koppers (CEOAS, Oregon State University).
# The frozenset below reproduces their familiar colors and font styling.
"""
Theme constants matching the legacy Smart Coast / ERDPP Excel user-forms.
"""
from tkinter import font as _font

# ---- Koppers Smart Coast palette (from vbColor usage in the converted VBA) --
TEAL = '#33ccff'          # vb_rgb(51, 204, 255) : matching tag-name fields
PURPLE = '#9900cc'        # vb_rgb(153, 0, 204)  : long char fields
BLUE = '#0000ff'          # vb_rgb(0, 0, 255)    : dynamic list / authors / editors
GREEN = '#22c822'         # vb_rgb(34, 200, 34)  : date fields
MAGENTA = '#ff66ff'       # vb_rgb(255, 102, 255): flag / skip / status fields
RED = '#ff1414'           # vb_rgb(255, 20, 20)  : database fields
GRAY = '#464646'          # vb_rgb(70, 70, 70)   : url fields

PAGE_BG = '#e8f0f8'
PAGE_BG_ALT = '#dbe7f2'
HEADER_BG = '#0f3f7a'
HEADER_FG = '#ffffff'
FIELD_BG = '#ffffff'
FIELD_DISABLED_BG = '#e0e0e0'
BUTTON_BG = '#4a90d9'
BUTTON_FG = '#ffffff'
OUTLINE = '#90a8c0'

# Tagname -> HotButton colour (mirrors UserForm1_Settings / Show_HotButton).
TAG_COLORS = {
    'char': PURPLE,
    'dyna': BLUE,
    'authors': BLUE,
    'editors': BLUE,
    'date': GREEN,
    'flag': MAGENTA,
    'skip': MAGENTA,
    'stat': MAGENTA,
    'status': MAGENTA,
    'database': RED,
    'url': GRAY,
}
# Default colour used for columns that appear in the "tagnames" range.
MATCH_COLOR = TEAL


def base_font():
    try:
        return _font.nametofont('TkDefaultFont').copy()
    except Exception:
        f = _font.Font()
        return f


def app_font(size=10, bold=False):
    f = base_font()
    f.configure(size=size, weight='bold' if bold else 'normal')
    return f