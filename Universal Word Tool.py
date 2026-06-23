# -*- coding: utf-8 -*-
"""
===============================================================================
 UNIVERSAL WORD TOOL  (IronPython 2.7)
===============================================================================
 A single GUI front-end that bundles all the individual document-editing
 scripts in this folder. Workflow:

   1. SELECT  - tick the tools you want to run (hover for a description).
   2. CONFIGURE - a tabbed window: a General tab (page / row / column range
      + run order) and one tab per selected tool with its own settings.
   3. SAVE & RUN - settings are written to 'config.json' (next to this file)
      and reloaded next launch. A working *copy* of your .docx is processed,
      so the original is never touched.

-------------------------------------------------------------------------------
 HOW THIS FILE IS ORGANISED  (so you can edit one tool without hunting):

   SECTION A  - Imports & Word constants
   SECTION B  - Shared helpers (text cleaning, currency, colour, COM ctx)
   SECTION C  - TOOL DEFINITIONS  <-- edit individual tools here.
                Each tool is ONE labelled block containing, in order:
                    * its process_*() function (what it does to the doc)
                    * its TOOL_* dict (id, label, tooltip, settings fields)
   SECTION D  - Tool registry (the ordered list of all tools)
   SECTION E  - Config load / save (JSON)
   SECTION F  - GUI: settings-field rendering
   SECTION G  - GUI: selection window
   SECTION H  - GUI: settings window (tabs)
   SECTION I  - Runner (opens Word, runs tools in order, saves, reports)
   SECTION J  - Entry point

 FIELD TYPES available to a tool's "fields" list (SECTION C):
   text       - single-line string
   multitext  - multi-line string (taller box)
   int        - whole number (required)
   float      - decimal number (required)
   optint     - whole number, blank allowed (-> None)
   bool       - checkbox
   choice     - dropdown; supply "choices": [...]
   keymap     - editable keyword -> hex-colour table (Highlight Cells)
   stringlist - editable list of strings (Insert Images patterns)
===============================================================================
"""

from __future__ import print_function

# ===========================================================================
# SECTION A - IMPORTS & WORD CONSTANTS
# ===========================================================================
import os
import sys
import traceback

# --- Crash reporting -------------------------------------------------------
# If anything goes wrong (even during imports) we want the error to be
# READABLE rather than flashing past as the console window closes. These two
# helpers hold the window open, write an error_log.txt next to the script, and
# (when possible) show a blocking dialog with the full traceback.

def _hold_console():
    try:
        raw_input("\nPress Enter to close this window...")
    except Exception:
        try:
            input("\nPress Enter to close this window...")  # py3 fallback
        except Exception:
            pass


def _report_fatal(context):
    tb = traceback.format_exc()
    print("\n" + "=" * 70)
    print("FATAL ERROR during: %s" % context)
    print("=" * 70)
    print(tb)
    try:
        log_dir = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        log_dir = os.getcwd()
    try:
        with open(os.path.join(log_dir, "error_log.txt"), "w") as f:
            f.write("FATAL ERROR during: %s\n\n%s" % (context, tb))
        print("(A copy was written to error_log.txt)")
    except Exception:
        pass
    # A blocking dialog survives the console closing - try it if WinForms loaded
    try:
        from System.Windows.Forms import (
            MessageBox, MessageBoxButtons, MessageBoxIcon)
        MessageBox.Show(tb, "Universal Word Tool - Fatal Error (%s)" % context,
                        MessageBoxButtons.OK, MessageBoxIcon.Error)
    except Exception:
        pass
    _hold_console()


try:
    import clr
    import re
    import json
    import shutil

    clr.AddReference("System.Windows.Forms")
    clr.AddReference("System.Drawing")
    from System.Windows.Forms import (
        Application, Form, Label, Button, TextBox, CheckBox, ComboBox,
        CheckedListBox, ListBox, TabControl, TabPage, Panel, ToolTip,
        DataGridView, DataGridViewAutoSizeColumnsMode,
        OpenFileDialog, MessageBox, MessageBoxButtons, MessageBoxIcon,
        DialogResult, FormBorderStyle, FormStartPosition, SelectionMode,
        ComboBoxStyle, AnchorStyles, ScrollBars, BorderStyle, DockStyle
    )
    from System.Drawing import Size, Point, Font, FontStyle, Color

    clr.AddReference("Microsoft.Office.Interop.Word")
    import Microsoft.Office.Interop.Word as Word
except Exception:
    _report_fatal("startup imports (IronPython / Word interop not available?)")
    raise SystemExit(1)

# Word row-height rule constants
WD_ROW_HEIGHT_AT_LEAST = 1
WD_ROW_HEIGHT_EXACTLY = 2
# Centimetres -> points (Word measures height in points)
CM_TO_POINTS = 28.3464567

# Path of the config file (always next to this script)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "config.json")


# ===========================================================================
# SECTION B - SHARED HELPERS
# ===========================================================================

def clean(text):
    """Strip Word's cell-end (\\x07) and carriage returns, then whitespace."""
    return text.replace("\r", "").replace("\x07", "").strip()


def hex_to_rgb(h):
    """'63BE7B' or '#63BE7B' -> (99, 190, 123)."""
    h = h.strip().lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def is_valid_hex(h):
    h = h.strip().lstrip("#")
    if len(h) != 6:
        return False
    try:
        int(h, 16)
        return True
    except ValueError:
        return False


def rgb_to_wdcolor(r, g, b):
    """Word stores cell shading colour as r + g*256 + b*65536."""
    return r + (g * 256) + (b * 65536)


def format_as_currency(value):
    """'4800' -> '$4,800'; '1200.5' -> '$1,200.50'. None on failure."""
    try:
        clean_val = value.replace(",", "").replace("$", "").strip()
        number = float(clean_val)
        if number == int(number):
            return "$%s" % "{:,.0f}".format(number)
        return "$%s" % "{:,.2f}".format(number)
    except (ValueError, TypeError):
        return None


def is_plain_number(text):
    """True only for an unformatted number (no existing '$' or ',')."""
    clean_val = text.replace(",", "").replace("$", "").strip()
    try:
        float(clean_val)
        if "$" in text or "," in text:
            return False
        return True
    except ValueError:
        return False


def table_start_page(doc, table):
    """Return the page number a table starts on (needs pagination on)."""
    try:
        rng = doc.Range(table.Range.Start, table.Range.Start)
        return rng.Information(Word.WdInformation.wdActiveEndPageNumber)
    except Exception:
        return None


class Ctx(object):
    """Per-tool execution context. Wraps the open document and the global
    page/row/column range so every tool obeys the same scoping rules.

      ctx.settings - this tool's own settings dict
      ctx.tables() - yields (index, table) for tables inside the page range
      ctx.rows(table, lo, hi) - row indices inside the global row range
      ctx.col_ok(idx) - is this column index inside the global column range?
    """

    def __init__(self, word_app, doc, settings, glob):
        self.word_app = word_app
        self.doc = doc
        self.settings = settings
        self.glob = glob
        self._page_filter = bool(glob.get("start_page") or glob.get("finish_page"))

    def tables(self):
        for i in range(1, self.doc.Tables.Count + 1):
            table = self.doc.Tables.Item(i)
            if self._page_ok(table):
                yield i, table

    def _page_ok(self, table):
        if not self._page_filter:
            return True
        page = table_start_page(self.doc, table)
        if page is None:
            return True
        sp = self.glob.get("start_page")
        fp = self.glob.get("finish_page")
        if sp and page < sp:
            return False
        if fp and page > fp:
            return False
        return True

    def rows(self, table, lo=1, hi=None):
        n = table.Rows.Count
        if hi is None:
            hi = n
        start = max(1, lo, self.glob.get("start_row") or 1)
        end = min(n, hi, self.glob.get("end_row") or n)
        if start > end:
            return range(0)
        return range(start, end + 1)

    def col_ok(self, idx):
        sc = self.glob.get("start_col")
        ec = self.glob.get("end_col")
        if sc and idx < sc:
            return False
        if ec and idx > ec:
            return False
        return True


# ===========================================================================
# SECTION C - TOOL DEFINITIONS
# ---------------------------------------------------------------------------
# Each block below is self-contained: the processing function followed by the
# tool's definition dict. To tweak a tool, edit only its block. To add a new
# tool, copy a block, then add the new TOOL_* dict to the TOOLS list in
# SECTION D.
# ===========================================================================


# --------------------------------------------------------------------------
# TOOL 1 - DELETE ROWS  (by trigger text)
# --------------------------------------------------------------------------
def process_delete_rows(ctx):
    s = ctx.settings
    trigger = s["trigger_text"]
    ci = s["case_insensitive"]
    needle = trigger.lower() if ci else trigger
    deleted = 0

    for ti, table in ctx.tables():
        table_text = table.Range.Text
        hay = table_text.lower() if ci else table_text
        if needle not in hay:
            continue

        to_delete = []
        for ri in ctx.rows(table):
            row = table.Rows.Item(ri)
            row_text = row.Range.Text
            if needle not in (row_text.lower() if ci else row_text):
                continue
            # Confirm the trigger sits in a cell within the column range
            for cidx in range(1, row.Cells.Count + 1):
                if not ctx.col_ok(cidx):
                    continue
                try:
                    cell_text = clean(row.Cells.Item(cidx).Range.Text)
                except Exception:
                    continue
                if needle in (cell_text.lower() if ci else cell_text):
                    to_delete.append(ri)
                    break

        for ri in reversed(to_delete):
            try:
                table.Rows.Item(ri).Delete()
                deleted += 1
            except Exception as e:
                print("  Error deleting row %d: %s" % (ri, str(e)))

        if ti % 20 == 0:
            ctx.doc.UndoClear()

    return "Delete Rows: removed %d row(s)" % deleted


TOOL_DELETE_ROWS = {
    "id": "delete_rows",
    "label": "1. Delete Rows (by trigger text)",
    "description": ("Deletes any entire table row that contains the trigger "
                    "text (default 'delete line'). Searches all columns "
                    "within the global column range."),
    "process": process_delete_rows,
    "fields": [
        {"key": "trigger_text", "type": "text", "label": "Trigger text",
         "default": "delete line"},
        {"key": "case_insensitive", "type": "bool", "label": "Case-insensitive match",
         "default": True},
    ],
}


# --------------------------------------------------------------------------
# TOOL 2 - DELETE EMPTY ROWS
# --------------------------------------------------------------------------
def _row_is_empty(row, s):
    for cidx in range(1, row.Cells.Count + 1):
        try:
            cell = row.Cells.Item(cidx)
            if s["check_text"] and clean(cell.Range.Text):
                return False
            if s["check_inline_shapes"] and cell.Range.InlineShapes.Count > 0:
                return False
            if s["check_floating_shapes"] and cell.Range.ShapeRange.Count > 0:
                return False
        except Exception as e:
            # Unreadable cell -> treat row as non-empty to be safe
            print("  Error checking a cell: %s" % str(e))
            return False
    return True


def process_delete_empty_rows(ctx):
    s = ctx.settings
    deleted = 0

    for ti, table in ctx.tables():
        to_delete = []
        for ri in ctx.rows(table):
            try:
                row = table.Rows.Item(ri)
                if _row_is_empty(row, s):
                    to_delete.append(ri)
            except Exception as e:
                print("  Error checking row %d: %s" % (ri, str(e)))

        for ri in reversed(to_delete):
            try:
                table.Rows.Item(ri).Delete()
                deleted += 1
            except Exception as e:
                print("  Error deleting row %d: %s" % (ri, str(e)))

        if ti % 20 == 0:
            ctx.doc.UndoClear()

    return "Delete Empty Rows: removed %d row(s)" % deleted


TOOL_DELETE_EMPTY = {
    "id": "delete_empty_rows",
    "label": "2. Delete Empty Rows",
    "description": ("Deletes table rows where every cell is empty. Choose what "
                    "counts as 'empty': text, inline images, floating shapes."),
    "process": process_delete_empty_rows,
    "fields": [
        {"key": "check_text", "type": "bool",
         "label": "Empty requires: no text", "default": True},
        {"key": "check_inline_shapes", "type": "bool",
         "label": "Empty requires: no inline images", "default": True},
        {"key": "check_floating_shapes", "type": "bool",
         "label": "Empty requires: no floating shapes", "default": True},
    ],
}


# --------------------------------------------------------------------------
# TOOL 3 - HIGHLIGHT CELLS
# --------------------------------------------------------------------------
def process_highlight(ctx):
    s = ctx.settings
    ci = s["case_insensitive"]

    # Build lookup: keyword -> (r, g, b)
    colour_map = {}
    for pair in s["color_map"]:
        kw, hex_val = pair[0], pair[1]
        key = kw.lower() if ci else kw
        colour_map[key] = hex_to_rgb(hex_val)

    highlighted = 0
    for ti, table in ctx.tables():
        table_text = table.Range.Text
        table_hay = table_text.lower() if ci else table_text
        if not any(k in table_hay for k in colour_map):
            continue

        for ri in ctx.rows(table):
            row = table.Rows.Item(ri)
            for cidx in range(1, row.Cells.Count + 1):
                if not ctx.col_ok(cidx):
                    continue
                try:
                    cell = row.Cells.Item(cidx)
                    cell_text = clean(cell.Range.Text)
                    key = cell_text.lower() if ci else cell_text
                    if key in colour_map:
                        r, g, b = colour_map[key]
                        cell.Shading.BackgroundPatternColor = rgb_to_wdcolor(r, g, b)
                        highlighted += 1
                except Exception as e:
                    print("  Error at Row %d Cell %d: %s" % (ri, cidx, str(e)))

        if ti % 20 == 0:
            ctx.doc.UndoClear()

    return "Highlight Cells: coloured %d cell(s)" % highlighted


TOOL_HIGHLIGHT = {
    "id": "highlight_cells",
    "label": "3. Highlight Cells",
    "description": ("Fills a cell with a colour when the cell's exact text "
                    "matches a keyword (e.g. P1, C3 - average, Yes). Edit the "
                    "keyword -> hex-colour table below."),
    "process": process_highlight,
    "fields": [
        {"key": "case_insensitive", "type": "bool",
         "label": "Case-insensitive match", "default": True},
        {"key": "color_map", "type": "keymap",
         "label": "Keyword -> Hex colour",
         "default": [
             ["p5", "63BE7B"], ["p4", "B1D47F"], ["p3", "FFEB84"],
             ["p2", "FCAA78"], ["p1", "F8696B"],
             ["p5 (5-10y)", "63BE7B"], ["p4 (2-5y)", "B1D47F"],
             ["p3 (12-24m)", "FFEB84"], ["p2 (0-12m)", "FCAA78"],
             ["p1 (immediate)", "F8696B"],
             ["no", "FF9B9B"], ["yes", "A8D08D"],
             ["c1 - excellent", "63BE7B"], ["c2 - good", "B1D47F"],
             ["c3 - average", "FFEB84"], ["c4 - poor", "FCAA78"],
             ["c5 - end of life", "F8696B"], ["c5 - very poor", "F8696B"],
         ]},
    ],
}


# --------------------------------------------------------------------------
# TOOL 4 - REMOVE TEXT
# --------------------------------------------------------------------------
def process_remove_text(ctx):
    s = ctx.settings
    target = int(s["target_column"])
    target_text = s["text_to_remove"]
    removed = 0

    for ti, table in ctx.tables():
        if target_text not in table.Range.Text:
            continue
        for ri in ctx.rows(table):
            row = table.Rows.Item(ri)
            if row.Cells.Count < target:
                continue
            try:
                cell = row.Cells.Item(target)
                cell_text = clean(cell.Range.Text)
                if target_text in cell_text:
                    cell.Range.Text = cell_text.replace(target_text, "").strip()
                    removed += 1
            except Exception as e:
                print("  Error at Row %d: %s" % (ri, str(e)))

        if ti % 20 == 0:
            ctx.doc.UndoClear()

    return "Remove Text: cleaned %d cell(s)" % removed


TOOL_REMOVE_TEXT = {
    "id": "remove_text",
    "label": "4. Remove Text",
    "description": ("Removes a specific block of text from the chosen column, "
                    "leaving any other content in the cell intact."),
    "process": process_remove_text,
    "fields": [
        {"key": "target_column", "type": "int", "label": "Target column",
         "default": 3},
        {"key": "text_to_remove", "type": "multitext", "label": "Text to remove",
         "default": ("Note: Replacement of the switchboard will remediate all "
                     "risk items and the associated cost in the risk register "
                     "relating to the switchboard.")},
    ],
}


# --------------------------------------------------------------------------
# TOOL 5 - FORMAT CURRENCY  (merges 'Below Cost Estimate' + 'Row 10 onwards')
# --------------------------------------------------------------------------
def process_currency(ctx):
    s = ctx.settings
    mode = s["mode"]              # "after_marker" | "fixed_start_row"
    target = int(s["target_column"])
    skip_last = int(s["skip_last_rows"])
    formatted = 0

    for ti, table in ctx.tables():
        n = table.Rows.Count

        if mode == "after_marker":
            marker = s["marker_text"]
            marker_row = None
            for ri in ctx.rows(table):
                if marker in table.Rows.Item(ri).Range.Text:
                    marker_row = ri
                    break
            if marker_row is None:
                continue
            lo = marker_row + 1
        else:  # fixed_start_row
            lo = int(s["start_row"])

        hi = n - skip_last

        for ri in ctx.rows(table, lo, hi):
            row = table.Rows.Item(ri)
            if row.Cells.Count < target:
                continue
            try:
                col = min(target, row.Cells.Count)
                cell = row.Cells.Item(col)
                cell_text = clean(cell.Range.Text)
                if cell_text and is_plain_number(cell_text):
                    new_val = format_as_currency(cell_text)
                    if new_val:
                        cell.Range.Text = new_val
                        formatted += 1
            except Exception as e:
                print("  Error at Table %d Row %d: %s" % (ti, ri, str(e)))

        if ti % 20 == 0:
            ctx.doc.UndoClear()

    return "Format Currency: formatted %d cell(s)" % formatted


TOOL_CURRENCY = {
    "id": "format_currency",
    "label": "5. Format Currency",
    "description": ("Formats plain numbers in a column as currency "
                    "(4800 -> $4,800). Mode 'after_marker' starts after a row "
                    "containing the marker text; 'fixed_start_row' starts at a "
                    "set row. 'Skip last rows' excludes trailing rows."),
    "process": process_currency,
    "fields": [
        {"key": "mode", "type": "choice", "label": "Mode",
         "choices": ["after_marker", "fixed_start_row"],
         "default": "after_marker"},
        {"key": "marker_text", "type": "text",
         "label": "Marker text (after_marker mode)", "default": "Cost Estimate"},
        {"key": "start_row", "type": "int",
         "label": "Start row (fixed_start_row mode)", "default": 10},
        {"key": "skip_last_rows", "type": "int", "label": "Skip last N rows",
         "default": 2},
        {"key": "target_column", "type": "int", "label": "Target column",
         "default": 4},
    ],
}


# --------------------------------------------------------------------------
# TOOL 6 - BOLD / UNDERLINE
# --------------------------------------------------------------------------
def process_bold_underline(ctx):
    s = ctx.settings
    tag = s["tag"]
    apply_mode = s["apply"]       # "both" | "bold" | "underline"
    target = int(s["target_column"])
    ci = s["case_insensitive"]
    needle = tag.lower() if ci else tag
    processed = 0

    tag_re = re.compile(re.escape(tag), re.IGNORECASE if ci else 0)

    for ti, table in ctx.tables():
        table_text = table.Range.Text
        if needle not in (table_text.lower() if ci else table_text):
            continue
        for ri in ctx.rows(table):
            row = table.Rows.Item(ri)
            if row.Cells.Count < target:
                continue
            try:
                cell = row.Cells.Item(target)
                cell_text = clean(cell.Range.Text)
                if needle not in (cell_text.lower() if ci else cell_text):
                    continue
                cell.Range.Text = tag_re.sub("", cell_text).strip()
                if apply_mode in ("both", "bold"):
                    cell.Range.Bold = True
                if apply_mode in ("both", "underline"):
                    cell.Range.Underline = Word.WdUnderline.wdUnderlineSingle
                processed += 1
            except Exception as e:
                print("  Error at Row %d: %s" % (ri, str(e)))

        if ti % 20 == 0:
            ctx.doc.UndoClear()

    return "Bold/Underline: formatted %d cell(s)" % processed


TOOL_BOLD = {
    "id": "bold_underline",
    "label": "6. Bold / Underline",
    "description": ("Finds a tag (default '(bold)') in the chosen column, "
                    "removes it, and applies bold and/or underline to the cell."),
    "process": process_bold_underline,
    "fields": [
        {"key": "tag", "type": "text", "label": "Tag to find", "default": "(bold)"},
        {"key": "apply", "type": "choice", "label": "Apply",
         "choices": ["both", "bold", "underline"], "default": "both"},
        {"key": "target_column", "type": "int", "label": "Target column",
         "default": 4},
        {"key": "case_insensitive", "type": "bool",
         "label": "Case-insensitive match", "default": True},
    ],
}


# --------------------------------------------------------------------------
# TOOL 7 - DELETE NUMBERED LINES
# --------------------------------------------------------------------------
def process_delete_numbered(ctx):
    s = ctx.settings
    require = s["require_text"]
    pattern = s["pattern"]
    target = int(s["target_column"])
    cleaned = 0

    for ti, table in ctx.tables():
        if require and require not in table.Range.Text:
            continue
        for ri in ctx.rows(table):
            row = table.Rows.Item(ri)
            if row.Cells.Count < target:
                continue
            try:
                cell = row.Cells.Item(target)
                # Keep \r for the regex; only strip the cell-end marker
                cell_text = cell.Range.Text.replace("\x07", "")
                if require and require not in cell_text:
                    continue
                new_text = re.sub(pattern, "", cell_text).strip()
                if new_text != cell_text.strip():
                    cell.Range.Text = new_text
                    cleaned += 1
            except Exception as e:
                print("  Error at Row %d: %s" % (ri, str(e)))

        if ti % 20 == 0:
            ctx.doc.UndoClear()

    return "Delete Numbered Lines: cleaned %d cell(s)" % cleaned


TOOL_DELETE_NUMBERED = {
    "id": "delete_numbered_lines",
    "label": "7. Delete Numbered Lines",
    "description": ("Removes a numbered line (e.g. '1. ...') from the chosen "
                    "column, but only in cells that contain the required marker "
                    "text (default 'Note: ')."),
    "process": process_delete_numbered,
    "fields": [
        {"key": "require_text", "type": "text", "label": "Required marker text",
         "default": "Note: "},
        {"key": "pattern", "type": "text", "label": "Regex pattern to remove",
         "default": r"[\r\n]+\d+\.[^\r\n]*"},
        {"key": "target_column", "type": "int", "label": "Target column",
         "default": 2},
    ],
}


# --------------------------------------------------------------------------
# TOOL 8 - SET MIN ROW HEIGHT
# --------------------------------------------------------------------------
def _set_height_by_cells(table, target_points):
    """Fallback for vertically-merged tables: iterate cells, one row each."""
    set_count = 0
    processed_rows = set()
    cells = table.Range.Cells
    for i in range(1, cells.Count + 1):
        try:
            cell = cells.Item(i)
            row_index = cell.RowIndex
            if row_index in processed_rows:
                continue
            processed_rows.add(row_index)

            rule = cell.HeightRule
            height = cell.Height
            if rule in (WD_ROW_HEIGHT_AT_LEAST, WD_ROW_HEIGHT_EXACTLY):
                if height is not None and 0 < height < 9000 and height >= target_points:
                    continue
            cell.HeightRule = WD_ROW_HEIGHT_AT_LEAST
            cell.Height = target_points
            set_count += 1
        except Exception as e:
            print("    Cell %d skipped (likely vertical merge): %s" % (i, str(e)))
    return set_count


def process_row_height(ctx):
    s = ctx.settings
    target_cm = float(s["height_cm"])
    target_points = target_cm * CM_TO_POINTS
    rows_set = 0

    for ti, table in ctx.tables():
        for ri in ctx.rows(table):
            try:
                row = table.Rows.Item(ri)
                rule = row.HeightRule
                height = row.Height
                if rule in (WD_ROW_HEIGHT_AT_LEAST, WD_ROW_HEIGHT_EXACTLY):
                    if height is not None and 0 < height < 9000 and height >= target_points:
                        continue
                row.HeightRule = WD_ROW_HEIGHT_AT_LEAST
                row.Height = target_points
                rows_set += 1
            except Exception as e:
                if "vertically merged" in str(e).lower():
                    rows_set += _set_height_by_cells(table, target_points)
                else:
                    print("  Table %d error: %s" % (ti, str(e)))

        if ti % 20 == 0:
            ctx.doc.UndoClear()

    return "Set Row Height: set %d row(s) to >= %.2f cm" % (rows_set, target_cm)


TOOL_ROW_HEIGHT = {
    "id": "row_height",
    "label": "8. Set Min Row Height",
    "description": ("Sets every table row to AT LEAST the given height in "
                    "centimetres. Rows that already need more space are left "
                    "alone."),
    "process": process_row_height,
    "fields": [
        {"key": "height_cm", "type": "float", "label": "Minimum height (cm)",
         "default": 0.50},
    ],
}


# --------------------------------------------------------------------------
# TOOL 9 - INSERT IMAGES
# --------------------------------------------------------------------------
def process_insert_images(ctx):
    s = ctx.settings
    patterns = s["match_patterns"]
    prefix = s["path_prefix"]
    suffix = s["path_suffix"]
    inserted = 0
    missing = 0

    for ti, table in ctx.tables():
        table_text = table.Range.Text
        if not any(p in table_text for p in patterns):
            continue
        for ri in ctx.rows(table):
            row = table.Rows.Item(ri)
            for cidx in range(1, row.Cells.Count + 1):
                if not ctx.col_ok(cidx):
                    continue
                try:
                    cell = row.Cells.Item(cidx)
                    cell_text = cell.Range.Text.strip()
                    if not any(p in cell_text for p in patterns):
                        continue
                    img_path = prefix + clean(cell_text) + suffix
                    if os.path.exists(img_path):
                        cell.Range.Text = ""
                        cell.Range.InlineShapes.AddPicture(
                            FileName=img_path, LinkToFile=False,
                            SaveWithDocument=True)
                        inserted += 1
                    else:
                        missing += 1
                        print("  File not found: %s" % img_path)
                except Exception as e:
                    print("  Error at Row %d Cell %d: %s" % (ri, cidx, str(e)))

        if ti % 20 == 0:
            ctx.doc.UndoClear()

    return "Insert Images: inserted %d, missing %d" % (inserted, missing)


TOOL_IMAGES = {
    "id": "insert_images",
    "label": "9. Insert Images",
    "description": ("Replaces file-path text in a cell with the actual image. "
                    "A cell matches if it contains any match pattern. An "
                    "optional prefix/suffix is added to the path before loading "
                    "(e.g. suffix '.jpg')."),
    "process": process_insert_images,
    "fields": [
        {"key": "match_patterns", "type": "stringlist",
         "label": "Match patterns (cell contains any)",
         "default": ["\\LCE", "\\\\"]},
        {"key": "path_prefix", "type": "text", "label": "Path prefix",
         "default": ""},
        {"key": "path_suffix", "type": "text", "label": "Path suffix",
         "default": ""},
    ],
}


# ===========================================================================
# SECTION D - TOOL REGISTRY
# ---------------------------------------------------------------------------
# The order here is the default order shown in the selection list. Run order
# is set per-run on the General tab.
# ===========================================================================
TOOLS = [
    TOOL_DELETE_ROWS,
    TOOL_DELETE_EMPTY,
    TOOL_HIGHLIGHT,
    TOOL_REMOVE_TEXT,
    TOOL_CURRENCY,
    TOOL_BOLD,
    TOOL_DELETE_NUMBERED,
    TOOL_ROW_HEIGHT,
    TOOL_IMAGES,
]

TOOLS_BY_ID = dict((t["id"], t) for t in TOOLS)

# Global settings shown on the General tab. (Page/row/column range.)
GLOBAL_FIELDS = [
    {"key": "start_page", "type": "optint", "label": "Start page (blank = first)",
     "default": None},
    {"key": "finish_page", "type": "optint", "label": "Finish page (blank = last)",
     "default": None},
    {"key": "start_row", "type": "optint", "label": "Start row (blank = first)",
     "default": None},
    {"key": "end_row", "type": "optint", "label": "End row (blank = last)",
     "default": None},
    {"key": "start_col", "type": "optint", "label": "Start column (blank = first)",
     "default": None},
    {"key": "end_col", "type": "optint", "label": "End column (blank = last)",
     "default": None},
]


# ===========================================================================
# SECTION E - CONFIG LOAD / SAVE
# ===========================================================================

def _copy_default(value):
    """Deep-copy a default value (lists/dicts) so edits don't mutate it."""
    if isinstance(value, (list, dict)):
        return json.loads(json.dumps(value))
    return value


def build_default_config():
    cfg = {"selected": [], "run_order": [], "global": {}, "tools": {}}
    for f in GLOBAL_FIELDS:
        cfg["global"][f["key"]] = _copy_default(f.get("default"))
    for tool in TOOLS:
        d = {}
        for f in tool["fields"]:
            d[f["key"]] = _copy_default(f["default"])
        cfg["tools"][tool["id"]] = d
    return cfg


def load_config():
    """Load config.json, filling any missing keys from defaults so the file
    stays forward-compatible when new fields are added to the code."""
    defaults = build_default_config()
    if not os.path.exists(CONFIG_PATH):
        return defaults
    try:
        with open(CONFIG_PATH, "r") as f:
            saved = json.load(f)
    except Exception as e:
        print("Could not read config.json (%s); using defaults." % str(e))
        return defaults

    cfg = defaults
    cfg["selected"] = saved.get("selected", [])
    cfg["run_order"] = saved.get("run_order", [])
    for k, v in saved.get("global", {}).items():
        if k in cfg["global"]:
            cfg["global"][k] = v
    for tid, tvals in saved.get("tools", {}).items():
        if tid in cfg["tools"]:
            for k, v in tvals.items():
                if k in cfg["tools"][tid]:
                    cfg["tools"][tid][k] = v
    return cfg


def save_config(cfg):
    with open(CONFIG_PATH, "w") as f:
        json.dump(cfg, f, indent=2)
    print("Saved config: %s" % CONFIG_PATH)


# ===========================================================================
# SECTION F - GUI: SETTINGS-FIELD RENDERING
# ---------------------------------------------------------------------------
# build_field() lays out one field's label + control on a panel and returns a
# zero-arg getter that reads (and validates) the current value. A getter may
# raise ValueError(message) which the save handler reports to the user.
# ===========================================================================

LABEL_X = 12
INPUT_X = 280      # moved right to give labels more width (220 -> 256 px)
INPUT_W = 400      # increased to keep inputs from feeling cramped
ROW_H = 30


def _add_label(panel, text, y):
    lbl = Label()
    lbl.Text = text
    lbl.Location = Point(LABEL_X, y + 3)
    lbl.Size = Size(INPUT_X - LABEL_X - 8, 40)
    panel.Controls.Add(lbl)


def build_field(panel, field, value, y):
    """Render one field at vertical position y. Returns (getter, next_y)."""
    ftype = field["type"]
    label = field["label"]

    if ftype in ("text", "int", "float", "optint"):
        _add_label(panel, label, y)
        tb = TextBox()
        tb.Location = Point(INPUT_X, y)
        tb.Size = Size(INPUT_W, 22)
        tb.Text = "" if value is None else str(value)
        panel.Controls.Add(tb)

        def getter(_tb=tb, _t=ftype, _l=label):
            raw = _tb.Text.strip()
            if _t == "text":
                return _tb.Text
            if _t == "optint":
                if raw == "":
                    return None
                try:
                    return int(raw)
                except ValueError:
                    raise ValueError("'%s' must be a whole number." % _l)
            if _t == "int":
                try:
                    return int(raw)
                except ValueError:
                    raise ValueError("'%s' must be a whole number." % _l)
            if _t == "float":
                try:
                    return float(raw.replace(",", "."))
                except ValueError:
                    raise ValueError("'%s' must be a number." % _l)
        return getter, y + ROW_H

    if ftype == "multitext":
        _add_label(panel, label, y)
        tb = TextBox()
        tb.Multiline = True
        tb.ScrollBars = ScrollBars.Vertical
        tb.Location = Point(INPUT_X, y)
        tb.Size = Size(INPUT_W, 70)
        tb.Text = "" if value is None else str(value)
        panel.Controls.Add(tb)
        return (lambda _tb=tb: _tb.Text), y + 78

    if ftype == "bool":
        cb = CheckBox()
        cb.Text = label
        cb.Location = Point(LABEL_X, y)
        cb.Size = Size(INPUT_X + INPUT_W - LABEL_X, 22)
        cb.Checked = bool(value)
        panel.Controls.Add(cb)
        return (lambda _cb=cb: _cb.Checked), y + ROW_H

    if ftype == "choice":
        _add_label(panel, label, y)
        combo = ComboBox()
        combo.DropDownStyle = ComboBoxStyle.DropDownList
        combo.Location = Point(INPUT_X, y)
        combo.Size = Size(INPUT_W, 22)
        for choice in field["choices"]:
            combo.Items.Add(choice)
        if value in field["choices"]:
            combo.SelectedItem = value
        elif combo.Items.Count > 0:
            combo.SelectedIndex = 0
        panel.Controls.Add(combo)
        return (lambda _c=combo: _c.SelectedItem), y + ROW_H

    if ftype == "keymap":
        _add_label(panel, label, y)
        grid = DataGridView()
        grid.Location = Point(INPUT_X, y)
        grid.Size = Size(INPUT_W, 220)
        grid.AllowUserToAddRows = True
        grid.AllowUserToDeleteRows = True
        grid.RowHeadersVisible = True
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        grid.Columns.Add("kw", "Keyword")
        grid.Columns.Add("hex", "Hex (RRGGBB)")
        grid.Columns[1].FillWeight = 60
        for pair in (value or []):
            grid.Rows.Add(pair[0], pair[1])
        panel.Controls.Add(grid)

        def getter(_g=grid, _l=label):
            result = []
            for row in _g.Rows:
                if row.IsNewRow:
                    continue
                kw = row.Cells[0].Value
                hx = row.Cells[1].Value
                if kw is None or str(kw).strip() == "":
                    continue
                hx = "" if hx is None else str(hx).strip()
                if not is_valid_hex(hx):
                    raise ValueError(
                        "'%s': '%s' has an invalid hex colour '%s' "
                        "(need 6 hex digits)." % (_l, kw, hx))
                result.append([str(kw), hx.lstrip("#")])
            return result
        return getter, y + 228

    if ftype == "stringlist":
        _add_label(panel, label, y)
        grid = DataGridView()
        grid.Location = Point(INPUT_X, y)
        grid.Size = Size(INPUT_W, 140)
        grid.AllowUserToAddRows = True
        grid.AllowUserToDeleteRows = True
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        grid.Columns.Add("val", "Pattern")
        for item in (value or []):
            grid.Rows.Add(item)
        panel.Controls.Add(grid)

        def getter(_g=grid):
            result = []
            for row in _g.Rows:
                if row.IsNewRow:
                    continue
                val = row.Cells[0].Value
                if val is not None and str(val) != "":
                    result.append(str(val))
            return result
        return getter, y + 148

    # Unknown type - skip gracefully
    _add_label(panel, "%s (unsupported field)" % label, y)
    return (lambda v=value: v), y + ROW_H


# ===========================================================================
# SECTION G - GUI: SELECTION WINDOW
# ===========================================================================
class SelectionForm(Form):
    def __init__(self, config):
        self.proceed = False
        self.selected_ids = []
        self._config = config

        self.Text = "Universal Word Tool - Select Tools"
        self.Size = Size(520, 470)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition = FormStartPosition.CenterScreen
        self.MaximizeBox = False

        lbl = Label()
        lbl.Text = "Tick the tools you want to run (hover for a description):"
        lbl.Location = Point(14, 12)
        lbl.Size = Size(480, 20)
        self.Controls.Add(lbl)

        self.clb = CheckedListBox()
        self.clb.Location = Point(14, 38)
        self.clb.Size = Size(478, 330)
        self.clb.CheckOnClick = True
        self.clb.IntegralHeight = False
        previously = set(config.get("selected", []))
        for tool in TOOLS:
            idx = self.clb.Items.Add(tool["label"])
            if tool["id"] in previously:
                self.clb.SetItemChecked(idx, True)
        self.Controls.Add(self.clb)

        # Per-item tooltip on hover
        self._tip = ToolTip()
        self._last_idx = [-1]
        self.clb.MouseMove += self._on_mouse_move

        btn_next = Button()
        btn_next.Text = "Next: Configure ->"
        btn_next.Font = Font("Segoe UI", 9, FontStyle.Bold)
        btn_next.Location = Point(14, 380)
        btn_next.Size = Size(478, 40)
        btn_next.Click += self._on_next
        self.Controls.Add(btn_next)

    def _on_mouse_move(self, sender, e):
        idx = self.clb.IndexFromPoint(e.Location)
        if idx != self._last_idx[0]:
            self._last_idx[0] = idx
            if 0 <= idx < len(TOOLS):
                self._tip.SetToolTip(self.clb, TOOLS[idx]["description"])
            else:
                self._tip.SetToolTip(self.clb, "")

    def _on_next(self, sender, e):
        ids = []
        for i in range(self.clb.Items.Count):
            if self.clb.GetItemChecked(i):
                ids.append(TOOLS[i]["id"])
        if not ids:
            MessageBox.Show("Select at least one tool.", "Nothing selected",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        self.selected_ids = ids
        self.proceed = True
        self.Close()


# ===========================================================================
# SECTION H - GUI: SETTINGS WINDOW (TABS)
# ===========================================================================
class SettingsForm(Form):
    def __init__(self, selected_ids, config):
        self.do_run = False
        self.config = config
        self.selected_ids = selected_ids
        self._tool_getters = {}    # tool_id -> {field_key: getter}
        self._global_getters = {}  # field_key -> getter

        self.Text = "Universal Word Tool - Settings"
        self.Size = Size(720, 600)
        self.FormBorderStyle = FormBorderStyle.Sizable
        self.StartPosition = FormStartPosition.CenterScreen

        self.tabs = TabControl()
        self.tabs.Location = Point(8, 8)
        self.tabs.Size = Size(696, 500)
        self.tabs.Anchor = (AnchorStyles.Top | AnchorStyles.Bottom |
                            AnchorStyles.Left | AnchorStyles.Right)
        self.Controls.Add(self.tabs)

        self._build_general_tab()
        for tid in selected_ids:
            self._build_tool_tab(TOOLS_BY_ID[tid])

        btn_run = Button()
        btn_run.Text = "Save & Run"
        btn_run.Font = Font("Segoe UI", 9, FontStyle.Bold)
        btn_run.Location = Point(540, 516)
        btn_run.Size = Size(164, 40)
        btn_run.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        btn_run.Click += self._on_run
        self.Controls.Add(btn_run)

        btn_cancel = Button()
        btn_cancel.Text = "Cancel"
        btn_cancel.Location = Point(8, 516)
        btn_cancel.Size = Size(120, 40)
        btn_cancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Left
        btn_cancel.Click += self._on_cancel
        self.Controls.Add(btn_cancel)

    # ---- General tab: page/row/col range + run order -------------------
    def _build_general_tab(self):
        tab = TabPage()
        tab.Text = "General"
        panel = Panel()
        panel.Dock = DockStyle.Fill
        panel.AutoScroll = True
        tab.Controls.Add(panel)

        y = 12
        hdr = Label()
        hdr.Text = "Global range (applies to every tool)"
        hdr.Font = Font("Segoe UI", 9, FontStyle.Bold)
        hdr.Location = Point(LABEL_X, y)
        hdr.Size = Size(420, 20)
        panel.Controls.Add(hdr)
        y += 28

        gvals = self.config["global"]
        for field in GLOBAL_FIELDS:
            getter, y = build_field(panel, field, gvals.get(field["key"]), y)
            self._global_getters[field["key"]] = getter

        y += 10
        hdr2 = Label()
        hdr2.Text = "Run order (top runs first)"
        hdr2.Font = Font("Segoe UI", 9, FontStyle.Bold)
        hdr2.Location = Point(LABEL_X, y)
        hdr2.Size = Size(420, 20)
        panel.Controls.Add(hdr2)
        y += 28

        self.order_list = ListBox()
        self.order_list.Location = Point(LABEL_X, y)
        self.order_list.Size = Size(440, 180)
        self.order_list.SelectionMode = SelectionMode.One
        self._order_ids = self._initial_order()
        for tid in self._order_ids:
            self.order_list.Items.Add(TOOLS_BY_ID[tid]["label"])
        panel.Controls.Add(self.order_list)

        btn_up = Button()
        btn_up.Text = "Up"
        btn_up.Location = Point(LABEL_X + 452, y)
        btn_up.Size = Size(80, 30)
        btn_up.Click += self._move_up
        panel.Controls.Add(btn_up)

        btn_down = Button()
        btn_down.Text = "Down"
        btn_down.Location = Point(LABEL_X + 452, y + 38)
        btn_down.Size = Size(80, 30)
        btn_down.Click += self._move_down
        panel.Controls.Add(btn_down)

        self.tabs.TabPages.Add(tab)

    def _initial_order(self):
        """Selected tools, ordered by the saved run_order where possible."""
        saved = [tid for tid in self.config.get("run_order", [])
                 if tid in self.selected_ids]
        rest = [tid for tid in self.selected_ids if tid not in saved]
        return saved + rest

    def _move_up(self, sender, e):
        idx = self.order_list.SelectedIndex
        if idx <= 0:
            return
        self._swap_order(idx, idx - 1)

    def _move_down(self, sender, e):
        idx = self.order_list.SelectedIndex
        if idx < 0 or idx >= self.order_list.Items.Count - 1:
            return
        self._swap_order(idx, idx + 1)

    def _swap_order(self, a, b):
        self._order_ids[a], self._order_ids[b] = self._order_ids[b], self._order_ids[a]
        item_a = self.order_list.Items[a]
        self.order_list.Items[a] = self.order_list.Items[b]
        self.order_list.Items[b] = item_a
        self.order_list.SelectedIndex = b

    # ---- One tab per tool ----------------------------------------------
    def _build_tool_tab(self, tool):
        tab = TabPage()
        tab.Text = tool["label"].split(".", 1)[0].strip() or tool["id"]
        panel = Panel()
        panel.Dock = DockStyle.Fill
        panel.AutoScroll = True
        tab.Controls.Add(panel)

        y = 12
        desc = Label()
        desc.Text = tool["description"]
        desc.Location = Point(LABEL_X, y)
        desc.Size = Size(640, 50)
        desc.Font = Font("Segoe UI", 8, FontStyle.Italic)
        panel.Controls.Add(desc)
        y += 58

        tvals = self.config["tools"][tool["id"]]
        getters = {}
        for field in tool["fields"]:
            getter, y = build_field(panel, field, tvals.get(field["key"]), y)
            getters[field["key"]] = getter
        self._tool_getters[tool["id"]] = getters

        self.tabs.TabPages.Add(tab)

    # ---- Save & Run -----------------------------------------------------
    def _collect(self):
        """Read every getter into the config. Raises ValueError on bad input."""
        for key, getter in self._global_getters.items():
            self.config["global"][key] = getter()
        for tid, getters in self._tool_getters.items():
            for key, getter in getters.items():
                self.config["tools"][tid][key] = getter()
        self.config["selected"] = list(self.selected_ids)
        self.config["run_order"] = list(self._order_ids)

    def _on_run(self, sender, e):
        try:
            self._collect()
        except ValueError as ex:
            MessageBox.Show(str(ex), "Invalid setting",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        save_config(self.config)
        self.do_run = True
        self.Close()

    def _on_cancel(self, sender, e):
        self.do_run = False
        self.Close()


# ===========================================================================
# SECTION I - RUNNER
# ===========================================================================

def run_tools(ordered_ids, config):
    """Pick a .docx, copy it, run each selected tool in order on the copy,
    save, and report a summary. The original document is never modified."""
    dialog = OpenFileDialog()
    dialog.Title = "Select Word Document to Process"
    dialog.Filter = "Word Documents (*.docx)|*.docx|All Files (*.*)|*.*"
    if dialog.ShowDialog() != DialogResult.OK:
        print("No document selected. Exiting.")
        return

    input_doc = dialog.FileName
    directory = os.path.dirname(input_doc)
    base = os.path.splitext(os.path.basename(input_doc))[0]
    output_doc = os.path.join(directory, base + "_processed.docx")

    glob = config["global"]
    page_filter = bool(glob.get("start_page") or glob.get("finish_page"))

    print("Input:  %s" % input_doc)
    print("Output: %s" % output_doc)
    print("Pagination (page-range filtering): %s" % ("ON" if page_filter else "OFF"))

    word_app = None
    doc = None
    results = []

    try:
        # Work on a copy so the original is untouched
        shutil.copy2(input_doc, output_doc)

        word_app = Word.ApplicationClass()
        word_app.Visible = False
        doc = word_app.Documents.Open(output_doc, ReadOnly=False)

        word_app.ScreenUpdating = False
        word_app.Options.CheckSpellingAsYouType = False
        word_app.Options.CheckGrammarAsYouType = False
        # Smart pagination: only pay the cost when a page range is set
        word_app.Options.Pagination = page_filter
        if page_filter:
            doc.Repaginate()

        for tid in ordered_ids:
            tool = TOOLS_BY_ID[tid]
            print("\n-- Running: %s --" % tool["label"])
            try:
                ctx = Ctx(word_app, doc, config["tools"][tid], glob)
                summary = tool["process"](ctx)
                print("   %s" % summary)
                results.append("OK   " + summary)
            except Exception as ex:
                traceback.print_exc()
                results.append("ERR  %s -- %s" % (tool["label"], str(ex)))

        word_app.Options.Pagination = True
        word_app.ScreenUpdating = True
        word_app.Options.CheckSpellingAsYouType = True
        word_app.Options.CheckGrammarAsYouType = True

        print("\nSaving...")
        doc.Save()
        doc.Close()
        doc = None
        print("Saved: %s" % output_doc)

        MessageBox.Show(
            "Run complete.\n\n" + "\n".join(results) +
            "\n\nSaved to:\n" + output_doc,
            "Done", MessageBoxButtons.OK, MessageBoxIcon.Information)

    except Exception:
        tb = traceback.format_exc()
        traceback.print_exc()
        try:
            log_dir = os.path.dirname(os.path.abspath(__file__))
            with open(os.path.join(log_dir, "error_log.txt"), "w") as f:
                f.write(tb)
        except Exception:
            pass
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        MessageBox.Show("Error while processing the document:\n\n" + tb,
                        "Universal Word Tool - Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error)
    finally:
        if word_app is not None:
            try:
                word_app.Options.Pagination = True
                word_app.ScreenUpdating = True
                word_app.Quit()
            except Exception:
                pass


# ===========================================================================
# SECTION J - ENTRY POINT
# ===========================================================================

def main():
    Application.EnableVisualStyles()
    config = load_config()

    # Step 1: select tools
    selection = SelectionForm(config)
    Application.Run(selection)
    if not selection.proceed:
        print("Cancelled at selection.")
        return

    # Step 2: configure settings (tabbed)
    settings = SettingsForm(selection.selected_ids, config)
    Application.Run(settings)
    if not settings.do_run:
        print("Cancelled at settings.")
        return

    # Step 3: run on a working copy
    run_tools(settings.config["run_order"], settings.config)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        _report_fatal("run")
