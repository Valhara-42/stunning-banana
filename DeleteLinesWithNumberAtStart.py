import clr
import os
import re

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import OpenFileDialog, MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult

clr.AddReference("Microsoft.Office.Interop.Word")
import Microsoft.Office.Interop.Word as Word

def clean_numbered_lines(doc_path, output_path):
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
        cleaned = 0

        for table_idx in range(1, doc.Tables.Count + 1):
            table = doc.Tables.Item(table_idx)

            if "Note: " not in table.Range.Text:
                continue

            print("Table %d: found 'Note: ', scanning column 2..." % table_idx)

            for row_idx in range(1, table.Rows.Count + 1):
                row = table.Rows.Item(row_idx)

                if row.Cells.Count < 2:
                    continue

                try:
                    cell = row.Cells.Item(2)

                    # Keep \r intact for regex matching, only strip the cell marker
                    cell_text = cell.Range.Text.replace("\x07", "")

                    if "Note: " not in cell_text:
                        continue

                    # Match a newline followed by number+period and remove
                    # only that line, stopping at the next newline
                    new_text = re.sub(r'[\r\n]+\d+\.[^\r\n]*', '', cell_text).strip()

                    if new_text != cell_text.strip():
                        cell.Range.Text = new_text
                        print("  Cleaned Row %d" % row_idx)
                        print("    Before: %s" % cell_text[:120].replace("\r", " / "))
                        print("    After:  %s" % new_text[:120])
                        cleaned += 1
                    else:
                        print("  Row %d: no match found (raw: %s)" % (row_idx, cell_text[:120].replace("\r", " / ")))

                except Exception as e:
                    print("  Error at Row %d: %s" % (row_idx, str(e)))

            if table_idx % 20 == 0:
                doc.UndoClear()
                print("  [Cleared undo buffer at table %d]" % table_idx)

        print("\n=== COMPLETE ===")
        print("Cleaned %d cells" % cleaned)

        word_app.Options.Pagination = True
        word_app.ScreenUpdating = True
        word_app.Options.CheckSpellingAsYouType = True
        word_app.Options.CheckGrammarAsYouType = True

        doc.Save()
        doc.Close()
        print("Saved: %s" % output_path)

        MessageBox.Show(
            "Complete!\n\nCleaned %d cells\n\nSaved to:\n%s" % (cleaned, output_path),
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
        output_doc = os.path.join(directory, name_without_ext + "_notes_cleaned.docx")

        print("Input:  %s" % input_doc)
        print("Output: %s" % output_doc)
        print("\nProcessing...")

        clean_numbered_lines(input_doc, output_doc)
    else:
        print("No file selected. Exiting.")

if __name__ == "__main__":
    main()