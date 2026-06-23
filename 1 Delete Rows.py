import clr
import sys
import os
import time

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import OpenFileDialog, MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult
clr.AddReference("Microsoft.Office.Interop.Word")
import Microsoft.Office.Interop.Word as Word

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

def delete_rows_with_delete_line(doc_path, output_path):
    word_app = Word.ApplicationClass()
    word_app.Visible = False

    try:
        # Open original read-only, copy to output, then work on the copy only
        doc = word_app.Documents.Open(doc_path, ReadOnly=True)
        doc.SaveAs2(output_path)
        doc.Close()

        doc = word_app.Documents.Open(output_path, ReadOnly=False)

        word_app.ScreenUpdating = False
        word_app.Options.CheckSpellingAsYouType = False
        word_app.Options.CheckGrammarAsYouType = False

        print("Document opened")

        table_count = com_retry(lambda: doc.Tables.Count)
        print("Total tables: %d" % table_count)

        rows_deleted = 0

        for table_idx in range(1, table_count + 1):
            table = com_retry(lambda: doc.Tables.Item(table_idx))
            print("\nProcessing Table %d" % table_idx)

            # Skip table entirely if "delete line" doesn't appear anywhere in it
            table_text = com_retry(lambda: table.Range.Text).lower()
            if "delete line" not in table_text:
                print("  No 'delete line' found, skipping.")
                continue

            rows_to_delete = []
            row_count = com_retry(lambda: table.Rows.Count)

            for row_idx in range(1, row_count + 1):
                # Row-level skip
                row = com_retry(lambda: table.Rows.Item(row_idx))
                row_text = com_retry(lambda: row.Range.Text).lower()
                if "delete line" not in row_text:
                    continue

                cell_count = com_retry(lambda: row.Cells.Count)
                for cell_idx in range(1, cell_count + 1):
                    try:
                        cell = com_retry(lambda: row.Cells.Item(cell_idx))
                        cell_text = com_retry(lambda: cell.Range.Text)
                        cell_text = cell_text.replace("\r", "").replace("\x07", "").strip()

                        if cell_text.lower() == "delete line":
                            print("  Found 'delete line' in Row %d, Cell %d - marking for deletion" % (row_idx, cell_idx))
                            rows_to_delete.append(row_idx)
                            break

                    except Exception as e:
                        print("  Error checking Row %d, Cell %d: %s" % (row_idx, cell_idx, str(e)))

            # Delete in reverse order so row indices stay valid
            for row_idx in reversed(rows_to_delete):
                try:
                    row = com_retry(lambda: table.Rows.Item(row_idx))
                    com_retry(lambda: row.Delete())
                    print("  Deleted row %d" % row_idx)
                    rows_deleted += 1
                except Exception as e:
                    print("  ERROR deleting row %d: %s" % (row_idx, str(e)))

        print("\n\n=== COMPLETE ===")
        print("Deleted %d rows" % rows_deleted)

        word_app.ScreenUpdating = True
        word_app.Options.CheckSpellingAsYouType = True
        word_app.Options.CheckGrammarAsYouType = True

        print("Saving...")
        doc.Save()
        doc.Close()
        print("Saved: %s" % output_path)

        MessageBox.Show(
            "Successfully deleted %d rows containing 'delete line'!\n\nSaved to:\n%s" % (rows_deleted, output_path),
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
        word_app.ScreenUpdating = True
        word_app.Options.CheckSpellingAsYouType = True
        word_app.Options.CheckGrammarAsYouType = True
        word_app.Quit()

def main():
    dialog = OpenFileDialog()
    dialog.Title = "Select Word Document to Process"
    dialog.Filter = "Word Documents (*.docx)|*.docx|All Files (*.*)|*.*"
    dialog.FilterIndex = 1

    if dialog.ShowDialog() == DialogResult.OK:
        input_doc = dialog.FileName

        directory = os.path.dirname(input_doc)
        filename = os.path.basename(input_doc)
        name_without_ext = os.path.splitext(filename)[0]
        output_doc = os.path.join(directory, name_without_ext + "_rows_deleted.docx")

        print("Input: %s" % input_doc)
        print("Output: %s" % output_doc)
        print("\nProcessing...")

        delete_rows_with_delete_line(input_doc, output_doc)
    else:
        print("No file selected. Exiting.")

if __name__ == "__main__":
    main()