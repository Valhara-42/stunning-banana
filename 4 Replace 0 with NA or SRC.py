import clr
import os
import shutil
import msvcrt

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import OpenFileDialog, MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult

clr.AddReference("Microsoft.Office.Interop.Word")
import Microsoft.Office.Interop.Word as Word

def get_choice():
    """Read a single keypress from the console."""
    while True:
        if msvcrt.kbhit():
            key = msvcrt.getwche()
            print("")
            return key.strip()

def find_and_replace_zeros(doc_path, output_path):
    word_app = Word.ApplicationClass()
    word_app.Visible = True
    doc = None
    temp_path = doc_path + ".working.docx"
    applied = 0
    skipped = 0

    try:
        shutil.copy2(doc_path, temp_path)
        print("Working copy created.")

        doc = word_app.Documents.Open(temp_path, ReadOnly=False)

        word_app.Options.CheckSpellingAsYouType = False
        word_app.Options.CheckGrammarAsYouType = False
        word_app.Options.Pagination = False

        print("Total tables: %d" % doc.Tables.Count)
        print("\nFor each '0' cell found, Word will scroll to it.")
        print("Press 1 = N/A, 2 = Switchboard Replacement Cost, any other key = Skip.\n")

        for table_idx in range(1, doc.Tables.Count + 1):
            table = doc.Tables.Item(table_idx)

            for row_idx in range(10, table.Rows.Count + 1):
                row = table.Rows.Item(row_idx)

                if row.Cells.Count < 4:
                    continue

                try:
                    cell = row.Cells.Item(4)
                    cell_text = cell.Range.Text.replace("\r", "").replace("\x07", "").strip()

                    if cell_text != "0":
                        continue

                    context = ""
                    if row.Cells.Count >= 2:
                        context = row.Cells.Item(2).Range.Text.replace("\x07", "").strip()

                    # Scroll Word to this cell
                    cell.Range.Select()

                    print("----------------------------------------")
                    print("Table %d, Row %d" % (table_idx, row_idx))
                    print("Column 2: %s" % context[:200].replace("\r", " "))
                    print("  1 = N/A")
                    print("  2 = Switchboard Replacement Cost")
                    print("  Any other key = Skip")
                    print("Your choice: ", end='')

                    choice = get_choice()

                    if choice == "1":
                        cell.Range.Text = "N/A"
                        print("  -> Set to 'N/A'")
                        applied += 1
                    elif choice == "2":
                        cell.Range.Text = "Switchboard Replacement Cost"
                        print("  -> Set to 'Switchboard Replacement Cost'")
                        applied += 1
                    else:
                        print("  -> Skipped")
                        skipped += 1

                except Exception as e:
                    print("  Error at Table %d, Row %d: %s" % (table_idx, row_idx, str(e)))

        print("\n=== COMPLETE ===")
        print("Applied: %d, Skipped: %d" % (applied, skipped))

        word_app.Options.Pagination = True
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

    # Word fully quit - now safe to copy to output
    try:
        shutil.copy2(temp_path, output_path)
        os.remove(temp_path)
        print("Saved: %s" % output_path)

        MessageBox.Show(
            "Complete!\n\nApplied: %d replacements\nSkipped: %d\n\nSaved to:\n%s" % (applied, skipped, output_path),
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
        output_doc = os.path.join(directory, name_without_ext + "_zeros_reviewed.docx")

        print("Input:  %s" % input_doc)
        print("Output: %s" % output_doc)
        print("\nScanning...\n")

        find_and_replace_zeros(doc_path=input_doc, output_path=output_doc)
    else:
        print("No file selected. Exiting.")