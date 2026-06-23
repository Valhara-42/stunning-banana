import clr
import os

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import OpenFileDialog, MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult

clr.AddReference("Microsoft.Office.Interop.Word")
import Microsoft.Office.Interop.Word as Word

TARGET_TEXT = "Note: Replacement of the switchboard will remediate all risk items and the associated cost in the risk register relating to the switchboard."

def remove_target_text(doc_path, output_path):
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

        print("Total tables: %d" % doc.Tables.Count)
        removed = 0

        for table_idx in range(1, doc.Tables.Count + 1):
            table = doc.Tables.Item(table_idx)

            # Skip table entirely if target text isn't present anywhere in it
            table_text = table.Range.Text
            if TARGET_TEXT not in table_text:
                continue

            print("Table %d: target text found, scanning column 3..." % table_idx)

            for row_idx in range(1, table.Rows.Count + 1):
                row = table.Rows.Item(row_idx)

                # Skip rows where column 3 doesn't exist
                if row.Cells.Count < 3:
                    continue

                try:
                    cell = row.Cells.Item(3)
                    cell_text = cell.Range.Text.replace("\r", "").replace("\x07", "").strip()

                    if TARGET_TEXT in cell_text:
                        # Replace just the target string, preserving any other content in the cell
                        new_text = cell_text.replace(TARGET_TEXT, "").strip()
                        cell.Range.Text = new_text
                        print("  Removed from Row %d" % row_idx)
                        removed += 1

                except Exception as e:
                    print("  Error at Row %d: %s" % (row_idx, str(e)))

            if table_idx % 20 == 0:
                doc.UndoClear()
                print("  [Cleared undo buffer at table %d]" % table_idx)

        print("\n=== COMPLETE ===")
        print("Removed text from %d cells" % removed)

        word_app.Options.Pagination = True
        word_app.ScreenUpdating = True
        word_app.Options.CheckSpellingAsYouType = True
        word_app.Options.CheckGrammarAsYouType = True

        doc.Save()
        doc.Close()
        print("Saved: %s" % output_path)

        MessageBox.Show(
            "Complete!\n\nRemoved target text from %d cells\n\nSaved to:\n%s" % (removed, output_path),
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
        word_app.Quit()


def main():
    dialog = OpenFileDialog()
    dialog.Title = "Select Word Document to Process"
    dialog.Filter = "Word Documents (*.docx)|*.docx|All Files (*.*)|*.*"
    dialog.FilterIndex = 1

    if dialog.ShowDialog() == DialogResult.OK:
        input_doc = dialog.FileName
        directory = os.path.dirname(input_doc)
        name_without_ext = os.path.splitext(os.path.basename(input_doc))[0]
        output_doc = os.path.join(directory, name_without_ext + "_text_removed.docx")

        print("Input:  %s" % input_doc)
        print("Output: %s" % output_doc)
        print("\nProcessing...")

        remove_target_text(input_doc, output_doc)
    else:
        print("No file selected. Exiting.")

if __name__ == "__main__":
    main()