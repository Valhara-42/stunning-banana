import clr
import os
import sys

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Windows.Forms import (
    Form, Label, Button, ListBox, MessageBox,
    MessageBoxButtons, MessageBoxIcon, DialogResult,
    OpenFileDialog, FormBorderStyle, FormStartPosition,
    SelectionMode, Application
)
from System.Drawing import Size, Point, Font, FontStyle

# ---------------------------------------------------------------------------
# Map each script filename to its processing function.
# Add new scripts here as you create them.
# ---------------------------------------------------------------------------
FUNCTION_MAP = {
    "1 Delete Rows.py":                         "delete_rows_with_delete_line",
    "2 Highlight_Table_Cells.py":               "highlight_table_cells",
    "3 removeTextColumn3.py":                   "remove_target_text",
    "5 formatAsCurrency.py":                    "format_currency_cells",
    "6 Underline and Bold (Bold) cells.py":     "apply_bold_formatting",
    "Image Inserter 2.py":                      "replace_paths_with_images",
}

def run_script(script_path, input_path, output_path):
    script_name = os.path.basename(script_path)
    func_name = FUNCTION_MAP.get(script_name)
    if func_name is None:
        raise RuntimeError(
            "Script not in FUNCTION_MAP: %s\n"
            "Add it to FUNCTION_MAP at the top of 0 Master Script.py" % script_name
        )

    script_dir = os.path.dirname(os.path.abspath(script_path))
    if script_dir not in sys.path:
        sys.path.insert(0, script_dir)

    import imp

    mod = imp.load_source("_step_module", script_path)

    # Suppress MessageBox popups by replacing the reference inside the loaded module
    import System.Windows.Forms as _swf
    class _SilentMessageBox(object):
        @staticmethod
        def Show(*args):
            # Print the first two args (text, caption) to console
            text = args[0] if len(args) > 0 else ""
            caption = args[1] if len(args) > 1 else ""
            print("  [%s] %s" % (caption, text))
            return _swf.DialogResult.OK
    mod.MessageBox = _SilentMessageBox()

    func = getattr(mod, func_name)

    if func_name == "replace_paths_with_images":
        json_path = os.path.splitext(output_path)[0] + "_image_paths.json"
        func(input_path, output_path, json_path)
    else:
        func(input_path, output_path)


# ---------------------------------------------------------------------------
# Order window
# ---------------------------------------------------------------------------

class MasterForm(Form):
    def __init__(self, script_paths):
        self.Text = "Master Script"
        self.Size = Size(540, 480)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition = FormStartPosition.CenterScreen
        self.MaximizeBox = False

        lbl = Label()
        lbl.Text = "Scripts to run — use arrows to set order, then click Run:"
        lbl.Font = Font("Segoe UI", 9, FontStyle.Regular)
        lbl.Location = Point(16, 14)
        lbl.Size = Size(480, 20)

        self.listbox = ListBox()
        self.listbox.Location = Point(16, 42)
        self.listbox.Size = Size(380, 300)
        self.listbox.SelectionMode = SelectionMode.One
        self.listbox.HorizontalScrollbar = True
        for path in script_paths:
            self.listbox.Items.Add(path)

        btn_up = Button()
        btn_up.Text = "Up"
        btn_up.Location = Point(408, 42)
        btn_up.Size = Size(90, 32)
        btn_up.Click += self.move_up

        btn_down = Button()
        btn_down.Text = "Down"
        btn_down.Location = Point(408, 82)
        btn_down.Size = Size(90, 32)
        btn_down.Click += self.move_down

        btn_remove = Button()
        btn_remove.Text = "Remove"
        btn_remove.Location = Point(408, 130)
        btn_remove.Size = Size(90, 32)
        btn_remove.Click += self.remove_item

        lbl_note = Label()
        lbl_note.Text = (
            "New scripts must be added to FUNCTION_MAP\n"
            "at the top of 0 Master Script.py to work here."
        )
        lbl_note.Font = Font("Segoe UI", 8, FontStyle.Italic)
        lbl_note.Location = Point(16, 352)
        lbl_note.Size = Size(480, 36)

        btn_run = Button()
        btn_run.Text = "Select document and run"
        btn_run.Font = Font("Segoe UI", 9, FontStyle.Bold)
        btn_run.Location = Point(16, 398)
        btn_run.Size = Size(482, 40)
        btn_run.Click += self.on_run

        for ctrl in [lbl, self.listbox, btn_up, btn_down, btn_remove, lbl_note, btn_run]:
            self.Controls.Add(ctrl)

    def move_up(self, sender, e):
        idx = self.listbox.SelectedIndex
        if idx <= 0:
            return
        item = self.listbox.Items[idx]
        self.listbox.Items.RemoveAt(idx)
        self.listbox.Items.Insert(idx - 1, item)
        self.listbox.SelectedIndex = idx - 1

    def move_down(self, sender, e):
        idx = self.listbox.SelectedIndex
        if idx < 0 or idx >= self.listbox.Items.Count - 1:
            return
        item = self.listbox.Items[idx]
        self.listbox.Items.RemoveAt(idx)
        self.listbox.Items.Insert(idx + 1, item)
        self.listbox.SelectedIndex = idx + 1

    def remove_item(self, sender, e):
        idx = self.listbox.SelectedIndex
        if idx < 0:
            return
        self.listbox.Items.RemoveAt(idx)

    def on_run(self, sender, e):
        if self.listbox.Items.Count == 0:
            MessageBox.Show(
                "No scripts in the list.", "Nothing to run",
                MessageBoxButtons.OK, MessageBoxIcon.Warning
            )
            return

        dialog = OpenFileDialog()
        dialog.Title = "Select Word Document to Process"
        dialog.Filter = "Word Documents (*.docx)|*.docx|All Files (*.*)|*.*"
        if dialog.ShowDialog() != DialogResult.OK:
            return

        input_doc = dialog.FileName
        directory = os.path.dirname(input_doc)
        base = os.path.splitext(os.path.basename(input_doc))[0]

        self.Hide()

        current_input = input_doc
        results = []

        for step_num in range(self.listbox.Items.Count):
            script_path = self.listbox.Items[step_num]
            script_name = os.path.splitext(os.path.basename(script_path))[0]
            output_doc = os.path.join(
                directory,
                "%s_step%d_%s.docx" % (base, step_num + 1, script_name)
            )

            print("\n-- Step %d: %s --" % (step_num + 1, script_name))
            print("  Input:  %s" % current_input)
            print("  Output: %s" % output_doc)

            try:
                run_script(script_path, current_input, output_doc)
                results.append("OK   %s" % script_name)
                current_input = output_doc
            except Exception as ex:
                import traceback
                traceback.print_exc()
                results.append("ERR  %s  --  %s" % (script_name, str(ex)))
                answer = MessageBox.Show(
                    "Error in step %d (%s):\n\n%s\n\nContinue with remaining steps?" % (
                        step_num + 1, script_name, str(ex)),
                    "Step failed",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Error
                )
                if answer == DialogResult.No:
                    break

        MessageBox.Show(
            "Run complete.\n\n" + "\n".join(results) + "\n\nFinal output:\n" + current_input,
            "Done",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )
        self.Close()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    Application.EnableVisualStyles()

    # Step 1: pick scripts
    script_dialog = OpenFileDialog()
    script_dialog.Title = "Select scripts to run (hold Ctrl for multiple)"
    script_dialog.Filter = "Python Scripts (*.py)|*.py|All Files (*.*)|*.*"
    script_dialog.Multiselect = True

    if script_dialog.ShowDialog() != DialogResult.OK:
        print("No scripts selected. Exiting.")
    else:
        print("Selected %d script(s)." % len(script_dialog.FileNames))
        form = MasterForm(list(script_dialog.FileNames))
        Application.Run(form)