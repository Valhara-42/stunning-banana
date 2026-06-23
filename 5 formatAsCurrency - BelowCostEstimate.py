import clr
import os
import shutil
import re

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import OpenFileDialog, MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult

clr.AddReference("Microsoft.Office.Interop.Word")
import Microsoft.Office.Interop.Word as Word

def format_as_currency(value):
    """Convert a plain number string to currency format. eg 4800 -> $4,800"""
    try:
        clean = value.replace(",", "").replace("$", "").strip()
        number = float(clean)
        if number == int(number):
            return "$%s" % "{:,.0f}".format(number)
        else:
            return "$%s" % "{:,.2f}".format(number)
    except:
        return None

def is_plain_number(text):
    """Returns True if the text is a plain unformatted number eg 4800, 1200.5"""
    clean = text.replace(",", "").replace("$", "").strip()
    try:
        float(clean)
        if "$" in text or "," in text:
            return False
        return True
    except:
        return False

def find_cost_estimate_row(table):
    """Find the row index containing 'Cost Estimate' past row 6. Returns None if not found."""
    for row_idx in range(7, table.Rows.Count + 1):
        try:
            row = table.Rows.Item(row_idx)
            if "Cost Estimate" in row.Range.Text:
                return row_idx
        except:
            pass
    return None

def format_currency_cells(doc_path, output_path):
    word_app = Word.ApplicationClass()
    word_app.Visible = False
    doc = None
    temp_path = doc_path + ".working.docx"
    formatted = 0

    try:
        shutil.copy2(doc_path, temp_path)
        print("Working copy created.")

        doc = word_app.Documents.Open(temp_path, ReadOnly=False)

        word_app.ScreenUpdating = False
        word_app.Options.CheckSpellingAsYouType = False
        word_app.Options.CheckGrammarAsYouType = False
        word_app.Options.Pagination = False

        print("Total tables: %d" % doc.Tables.Count)

        # Pre-find all Cost Estimate row positions once
        # so we don't re-scan on every pass
        cost_rows = {}
        for table_idx in range(1, doc.Tables.Count + 1):
            table = doc.Tables.Item(table_idx)
            cost_row = find_cost_estimate_row(table)
            if cost_row is not None:
                cost_rows[table_idx] = cost_row
                print("Table %d: 'Cost Estimate' at row %d" % (table_idx, cost_row))

        print("Found %d tables with 'Cost Estimate'." % len(cost_rows))

        max_passes = 5
        for pass_num in range(1, max_passes + 1):
            pass_formatted = 0
            print("\n--- Pass %d ---" % pass_num)

            for table_idx, cost_row in cost_rows.items():
                table = doc.Tables.Item(table_idx)
                row_count = table.Rows.Count

                start_row = cost_row + 1
                end_row = row_count - 2

                if start_row > end_row:
                    continue

                for row_idx in range(start_row, end_row + 1):
                    try:
                        row = table.Rows.Item(row_idx)

                        if row.Cells.Count < 4:
                            continue

                        target_col = min(4, row.Cells.Count)
                        cell = row.Cells.Item(target_col)
                        cell_text = cell.Range.Text.replace("\r", "").replace("\x07", "").strip()

                        if not cell_text:
                            continue

                        if is_plain_number(cell_text):
                            new_value = format_as_currency(cell_text)
                            if new_value:
                                cell.Range.Text = new_value
                                print("  Table %d, Row %d: %s -> %s" % (table_idx, row_idx, cell_text, new_value))
                                formatted += 1
                                pass_formatted += 1

                    except Exception as e:
                        print("  Error at Table %d, Row %d: %s" % (table_idx, row_idx, str(e)))

                if table_idx % 20 == 0:
                    doc.UndoClear()
                    print("  [Cleared undo buffer at table %d]" % table_idx)

            print("Pass %d complete: %d cells formatted" % (pass_num, pass_formatted))

            if pass_formatted == 0:
                print("No more changes found - stopping early.")
                break

        print("\n=== COMPLETE ===")
        print("Total formatted: %d" % formatted)

        word_app.Options.Pagination = True
        word_app.ScreenUpdating = True
        word_app.Options.CheckSpellingAsYouType = True
        word_app.Options.CheckGrammarAsYouType = True

        doc.Save()
        doc.Close()
        doc = None

    except Exception as e:
        print("FATAL ERROR: %s" % str(e))
        import traceback
        traceback.print_exc()
        if doc is not None:
            try:
                doc.Close(False)
                doc = None
            except:
                pass
        MessageBox.Show(
            "Error: %s" % str(e),
            "Error",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        )
        return

    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except:
                pass
        word_app.Quit()

    try:
        shutil.copy2(temp_path, output_path)
        os.remove(temp_path)
        print("Saved: %s" % output_path)

        MessageBox.Show(
            "Complete!\n\nTotal formatted: %d cells\n\nSaved to:\n%s" % (formatted, output_path),
            "Complete",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )
    except Exception as e:
        print("ERROR copying to output: %s" % str(e))
        MessageBox.Show(
            "Processed OK but could not save to output path.\nWorking file is at:\n%s" % temp_path,
            "Save Error",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        )


if __name__ == "__main__":
    dialog = OpenFileDialog()
    dialog.Title = "Select Word Document to Process"
    dialog.Filter = "Word Documents (*.docx)|*.docx|All Files (*.*)|*.*"
    dialog.FilterIndex = 1

    if dialog.ShowDialog() == DialogResult.OK:
        input_doc = dialog.FileName
        directory = os.path.dirname(input_doc)
        name_without_ext = os.path.splitext(os.path.basename(input_doc))[0]
        output_doc = os.path.join(directory, name_without_ext + "_currency_formatted.docx")

        print("Input:  %s" % input_doc)
        print("Output: %s" % output_doc)
        print("\nProcessing...\n")

        format_currency_cells(doc_path=input_doc, output_path=output_doc)
    else:
        print("No file selected. Exiting.")