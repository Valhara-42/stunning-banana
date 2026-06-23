import clr
import os

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import OpenFileDialog, MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult

clr.AddReference("Microsoft.Office.Interop.Word")
import Microsoft.Office.Interop.Word as Word

COLOUR_MAP = {
    # Priority Ratings
    "p5": (99,  190, 123),
    "p4": (177, 212, 127),
    "p3": (255, 235, 132),
    "p2": (252, 170, 120),
    "p1": (248, 105, 107),
    "p5 (5-10y)": (99,  190, 123),
    "p4 (2-5y)": (177, 212, 127),
    "p3 (12-24m)": (255, 235, 132),
    "p2 (0-12m)": (252, 170, 120),
    "p1 (Immediate)": (248, 105, 107),
    # Yes / No
    "no":  (255, 155, 155),
    "yes": (168, 208, 141),
    # Switchboard Condition
    "c1 - excellent":  (99,  190, 123),
    "c2 - good":       (177, 212, 127),
    "c3 - average":    (255, 235, 132),
    "c4 - poor":       (252, 170, 120),
    "c5 - end of life":(248, 105, 107),
    "c5 - very poor":  (248, 105, 107),
}

def rgb_to_wdcolor(r, g, b):
    return r + (g * 256) + (b * 65536)

def highlight_table_cells(doc_path, output_path):
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
        highlighted = 0

        for table_idx in range(1, doc.Tables.Count + 1):
            table = doc.Tables.Item(table_idx)

            table_text = table.Range.Text.lower()
            if not any(key in table_text for key in COLOUR_MAP):
                print("Table %d: no matching keywords, skipping." % table_idx)
                continue

            print("Table %d: scanning cells..." % table_idx)

            for row_idx in range(1, table.Rows.Count + 1):
                row = table.Rows.Item(row_idx)

                row_text = row.Range.Text.lower()
                if not any(key in row_text for key in COLOUR_MAP):
                    continue

                for cell_idx in range(1, row.Cells.Count + 1):
                    try:
                        cell = row.Cells.Item(cell_idx)
                        cell_text = cell.Range.Text.replace("\r", "").replace("\x07", "").strip().lower()

                        if cell_text in COLOUR_MAP:
                            r, g, b = COLOUR_MAP[cell_text]
                            cell.Shading.BackgroundPatternColor = rgb_to_wdcolor(r, g, b)
                            print("  Row %d Cell %d: '%s' -> (%d,%d,%d)" % (row_idx, cell_idx, cell_text, r, g, b))
                            highlighted += 1

                    except Exception as e:
                        print("  Error at Row %d Cell %d: %s" % (row_idx, cell_idx, str(e)))

            if table_idx % 20 == 0:
                doc.UndoClear()
                print("  [Cleared undo buffer at table %d]" % table_idx)

        print("\n=== COMPLETE ===")
        print("Highlighted %d cells" % highlighted)

        word_app.Options.Pagination = True
        word_app.ScreenUpdating = True
        word_app.Options.CheckSpellingAsYouType = True
        word_app.Options.CheckGrammarAsYouType = True

        doc.Save()
        doc.Close()
        print("Saved: %s" % output_path)

        MessageBox.Show(
            "Cell highlighting complete!\n\nHighlighted %d cells\n\nSaved to:\n%s" % (highlighted, output_path),
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
    dialog.Title = "Select Word Document to Highlight"
    dialog.Filter = "Word Documents (*.docx)|*.docx|All Files (*.*)|*.*"
    dialog.FilterIndex = 1

    if dialog.ShowDialog() == DialogResult.OK:
        input_doc = dialog.FileName
        directory = os.path.dirname(input_doc)
        name_without_ext = os.path.splitext(os.path.basename(input_doc))[0]
        output_doc = os.path.join(directory, name_without_ext + "_highlighted.docx")

        print("Input:  %s" % input_doc)
        print("Output: %s" % output_doc)
        print("\nProcessing...")

        highlight_table_cells(input_doc, output_doc)
    else:
        print("No file selected. Exiting.")

if __name__ == "__main__":
    main()