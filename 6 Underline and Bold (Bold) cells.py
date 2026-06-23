import clr
import os

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import OpenFileDialog, MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult

clr.AddReference("Microsoft.Office.Interop.Word")
import Microsoft.Office.Interop.Word as Word

BOLD_TAG = "(bold)"

def apply_bold_formatting(doc_path, output_path):
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
        processed = 0

        for table_idx in range(1, doc.Tables.Count + 1):
            table = doc.Tables.Item(table_idx)

            # Skip table if "(bold)" doesn't appear anywhere in it
            if BOLD_TAG not in table.Range.Text.lower():
                continue

            print("Table %d: found '(bold)', scanning column 4..." % table_idx)

            for row_idx in range(1, table.Rows.Count + 1):
                row = table.Rows.Item(row_idx)

                # Skip rows that don't have a column 4
                if row.Cells.Count < 4:
                    continue

                try:
                    cell = row.Cells.Item(4)
                    cell_text = cell.Range.Text.replace("\r", "").replace("\x07", "").strip()

                    if BOLD_TAG not in cell_text.lower():
                        continue

                    # Remove the (bold) tag and clean up any extra whitespace
                    new_text = cell_text.replace("(bold)", "").replace("(Bold)", "").replace("(BOLD)", "").strip()

                    # Set the cleaned text
                    cell.Range.Text = new_text

                    # Apply bold and underline to the entire cell range
                    cell.Range.Bold = True
                    cell.Range.Underline = Word.WdUnderline.wdUnderlineSingle

                    print("  Formatted Row %d: '%s'" % (row_idx, new_text))
                    processed += 1

                except Exception as e:
                    print("  Error at Row %d: %s" % (row_idx, str(e)))

            if table_idx % 20 == 0:
                doc.UndoClear()
                print("  [Cleared undo buffer at table %d]" % table_idx)

        print("\n=== COMPLETE ===")
        print("Formatted %d cells" % processed)

        word_app.Options.Pagination = True
        word_app.ScreenUpdating = True
        word_app.Options.CheckSpellingAsYouType = True
        word_app.Options.CheckGrammarAsYouType = True

        doc.Save()
        doc.Close()
        print("Saved: %s" % output_path)

        MessageBox.Show(
            "Complete!\n\nFormatted %d cells\n\nSaved to:\n%s" % (processed, output_path),
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
        output_doc = os.path.join(directory, name_without_ext + "_bold_formatted.docx")

        print("Input:  %s" % input_doc)
        print("Output: %s" % output_doc)
        print("\nProcessing...")

        apply_bold_formatting(input_doc, output_doc)
    else:
        print("No file selected. Exiting.")

if __name__ == "__main__":
    main()