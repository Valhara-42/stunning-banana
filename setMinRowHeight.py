import clr
import sys
import os
import time

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import OpenFileDialog, MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult
clr.AddReference("Microsoft.VisualBasic")
from Microsoft.VisualBasic import Interaction
clr.AddReference("Microsoft.Office.Interop.Word")
import Microsoft.Office.Interop.Word as Word

# Word measures row height in points. 1 cm = 28.3464567 points.
CM_TO_POINTS = 28.3464567

# Word's "at least" row-height rule constant (wdRowHeightAtLeast)
WD_ROW_HEIGHT_AUTO     = 0
WD_ROW_HEIGHT_AT_LEAST = 1
WD_ROW_HEIGHT_EXACTLY  = 2

def com_retry(func, retries=5, delay=1.0):
    """Retry a COM call if Word isn't ready yet."""
    for attempt in range(retries):
        try:
            return func()
        except SystemError as e:
            if "RPC_E_CALL_REJECTED" in str(e) or "0x80010001" in str(e):
                print("  Word busy, retrying (%d/%d)..." % (attempt + 1, retries))
                time.sleep(delay)
            else:
                raise
    raise Exception("Word did not become ready after %d retries" % retries)
    
def process_table_by_cells(table, target_points):
    """Fallback for tables with vertical merges - work through cells instead
    of rows, since row indexing is blocked on merged tables."""
    set_count = 0
    skipped_count = 0
    processed_rows = set()

    cells = com_retry(lambda: table.Range.Cells)
    cell_count = com_retry(lambda: cells.Count)

    for i in range(1, cell_count + 1):
        try:
            cell = com_retry(lambda: cells.Item(i))

            # Only handle each row once
            row_index = com_retry(lambda: cell.RowIndex)
            if row_index in processed_rows:
                continue
            processed_rows.add(row_index)

            rule = com_retry(lambda: cell.HeightRule)
            height = com_retry(lambda: cell.Height)

            # Leave rows already explicitly set at or above target
            if rule in (WD_ROW_HEIGHT_AT_LEAST, WD_ROW_HEIGHT_EXACTLY):
                if height is not None and 0 < height < 9000 and height >= target_points:
                    skipped_count += 1
                    continue

            com_retry(lambda: setattr(cell, "HeightRule", WD_ROW_HEIGHT_AT_LEAST))
            com_retry(lambda: setattr(cell, "Height", target_points))
            set_count += 1

        except Exception as e:
            # Vertically merged cells raise here - skip them and keep going
            print("    Cell %d skipped (likely vertical merge): %s" % (i, str(e)))

    return set_count, skipped_count

def set_min_row_heights(doc_path, output_path, target_cm):
    target_points = target_cm * CM_TO_POINTS

    word_app = Word.ApplicationClass()
    word_app.Visible = False

    try:
        doc = word_app.Documents.Open(doc_path, ReadOnly=True)
        doc.SaveAs2(output_path)
        doc.Close()

        doc = word_app.Documents.Open(output_path, ReadOnly=False)

        word_app.ScreenUpdating = False
        word_app.Options.CheckSpellingAsYouType = False
        word_app.Options.CheckGrammarAsYouType = False
        word_app.Options.Pagination = False

        print("Minimum height: %.2f cm (%.2f points)" % (target_cm, target_points))

        table_count = com_retry(lambda: doc.Tables.Count)
        print("Total tables: %d" % table_count)

        rows_set = 0
        rows_skipped = 0

        for table_idx in range(1, table_count + 1):
            table = com_retry(lambda: doc.Tables.Item(table_idx))
            row_count = com_retry(lambda: table.Rows.Count)
            rows_set_in_table = 0
            for row_idx in range(1, row_count + 1):
                try:
                    row = com_retry(lambda: table.Rows.Item(row_idx))

                    rule = com_retry(lambda: row.HeightRule)
                    height = com_retry(lambda: row.Height)

                    # If the row already has an explicit height rule (at least
                    # or exactly) AND that height is a valid value >= target,
                    # leave it alone.
                    if rule in (WD_ROW_HEIGHT_AT_LEAST, WD_ROW_HEIGHT_EXACTLY):
                        if height is not None and 0 < height < 9000 and height >= target_points:
                            rows_skipped += 1
                            continue

                    # Otherwise bump it up to at-least the target
                    com_retry(lambda: setattr(row, "HeightRule", WD_ROW_HEIGHT_AT_LEAST))
                    com_retry(lambda: setattr(row, "Height", target_points))
                    rows_set += 1
                    rows_set_in_table += 1

                except Exception as e:
                    if "vertically merged" in str(e).lower():
                        # Switch to cell-based handling for this table
                        print("Table %d: vertical merges present, using cell-based pass..." % table_idx)
                        s, sk = process_table_by_cells(table, target_points)
                        rows_set += s
                        rows_skipped += sk
                    else:
                        print("  Table %d error: %s" % (table_idx, str(e)))
                    
            print("Set %d rows in table # %d" % (rows_set_in_table, table_idx))
            
            if table_idx % 20 == 0:
                doc.UndoClear()
                print("  [Cleared undo buffer at table %d]" % table_idx)

        print("\n=== COMPLETE ===")
        print("Set %d rows | Left %d rows unchanged (already >= target)" % (rows_set, rows_skipped))

        word_app.Options.Pagination = True
        word_app.ScreenUpdating = True
        word_app.Options.CheckSpellingAsYouType = True
        word_app.Options.CheckGrammarAsYouType = True

        print("Saving...")
        doc.Save()
        doc.Close()
        print("Saved: %s" % output_path)

        MessageBox.Show(
            "Set %d rows to a minimum of %.2f cm.\nLeft %d rows unchanged (already at or above target).\n\nSaved to:\n%s" % (
                rows_set, target_cm, rows_skipped, output_path),
            "Complete",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )

    except Exception as e:
        print("FATAL ERROR: %s" % str(e))
        import traceback
        traceback.print_exc()
        MessageBox.Show(
            "Error: %s" % str(e),
            "Error",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        )

    finally:
        word_app.Options.Pagination = True
        word_app.ScreenUpdating = True
        word_app.Options.CheckSpellingAsYouType = True
        word_app.Options.CheckGrammarAsYouType = True
        word_app.Quit()

def main():
    response = Interaction.InputBox(
        "Enter the minimum row height in cm.\n\nEvery table row will be set to at least this height "
        "(taller rows that need more space are left alone).",
        "Minimum Row Height",
        "0.50"
    )

    response = response.strip()
    if not response:
        print("Cancelled - no height entered.")
        return

    try:
        target_cm = float(response.replace(",", "."))
    except:
        MessageBox.Show(
            "'%s' is not a valid number." % response,
            "Invalid Input",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        )
        return

    if target_cm <= 0:
        MessageBox.Show("Height must be greater than 0.", "Invalid Input",
                        MessageBoxButtons.OK, MessageBoxIcon.Error)
        return

    dialog = OpenFileDialog()
    dialog.Title = "Select Word Document to Process"
    dialog.Filter = "Word Documents (*.docx)|*.docx|All Files (*.*)|*.*"
    dialog.FilterIndex = 1

    if dialog.ShowDialog() == DialogResult.OK:
        input_doc = dialog.FileName
        directory = os.path.dirname(input_doc)
        name_without_ext = os.path.splitext(os.path.basename(input_doc))[0]
        output_doc = os.path.join(directory, name_without_ext + "_row_heights.docx")

        print("Input:  %s" % input_doc)
        print("Output: %s" % output_doc)
        print("Height: %.2f cm" % target_cm)
        print("\nProcessing...")

        set_min_row_heights(input_doc, output_doc, target_cm)
    else:
        print("No file selected. Exiting.")

if __name__ == "__main__":
    main()