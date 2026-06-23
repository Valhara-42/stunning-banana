import clr
import os
import json

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import OpenFileDialog, FolderBrowserDialog, MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult
clr.AddReference("Microsoft.Office.Interop.Word")
import Microsoft.Office.Interop.Word as Word

def safe_print(text):
    print(str(text).encode('cp850', errors='replace').decode('cp850'))

def get_preceding_title(doc, table):
    """Get the text of the paragraph immediately before the table."""
    try:
        title_range = doc.Range(0, table.Range.Start)
        if title_range.Paragraphs.Count > 0:
            for i in range(title_range.Paragraphs.Count, 0, -1):
                para_text = title_range.Paragraphs.Item(i).Range.Text.strip()
                if para_text:
                    return para_text.replace("\r", "").replace("\x07", "").strip()
    except:
        pass
    return None

def reinsert_images(doc_path, output_path, path_map, replacement_dir=None):
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
        errors = 0
        not_found = []

        for table_idx in range(1, doc.Tables.Count + 1):
            table = doc.Tables.Item(table_idx)

            title = get_preceding_title(doc, table)
            if not title:
                title = "table_%d" % table_idx

            if title not in path_map:
                continue

            safe_print("\nFound match: %s" % title[:60])

            for entry in path_map[title]:
                row_idx = entry["row"]
                cell_idx = entry["cell"]
                img_path = entry["path"]

                # Use replacement directory if provided
                if replacement_dir:
                    filename = os.path.basename(img_path)
                    img_path = os.path.join(replacement_dir, filename)

                safe_print("  Row %d, Cell %d: %s" % (row_idx, cell_idx, img_path))

                if not os.path.exists(img_path):
                    safe_print("  ERROR: File not found: %s" % img_path)
                    not_found.append(img_path)
                    errors += 1
                    continue

                try:
                    cell = table.Rows.Item(row_idx).Cells.Item(cell_idx)
                    cell.Range.Delete()
                    cell.Range.InlineShapes.AddPicture(
                        FileName=img_path,
                        LinkToFile=False,
                        SaveWithDocument=True
                    )
                    print("  SUCCESS!")
                    processed += 1
                except Exception as e:
                    safe_print("  ERROR: %s" % str(e))
                    errors += 1

            if table_idx % 20 == 0:
                doc.UndoClear()
                print("  [Cleared undo buffer at table %d]" % table_idx)

        print("\n=== COMPLETE ===")
        print("Processed: %d, Errors: %d" % (processed, errors))

        if not_found:
            print("\nFiles not found:")
            for p in not_found:
                safe_print("  %s" % p)

        word_app.Options.Pagination = True
        word_app.ScreenUpdating = True
        word_app.Options.CheckSpellingAsYouType = True
        word_app.Options.CheckGrammarAsYouType = True

        doc.Save()
        doc.Close()
        print("Saved: %s" % output_path)

        MessageBox.Show(
            "Re-insertion complete!\n\nProcessed: %d\nErrors: %d\n\nSaved to:\n%s" % (processed, errors, output_path),
            "Complete",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )

    except Exception as e:
        safe_print("FATAL ERROR: %s" % str(e))
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
    # Step 1: Select the Word document
    doc_dialog = OpenFileDialog()
    doc_dialog.Title = "Select Word Document to Re-insert Images Into"
    doc_dialog.Filter = "Word Documents (*.docx)|*.docx|All Files (*.*)|*.*"

    if doc_dialog.ShowDialog() != DialogResult.OK:
        print("No document selected. Exiting.")
        return

    input_doc = doc_dialog.FileName

    # Step 2: Select the JSON path map
    json_dialog = OpenFileDialog()
    json_dialog.Title = "Select Image Path Map JSON"
    json_dialog.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*"
    json_dialog.InitialDirectory = os.path.dirname(input_doc)

    if json_dialog.ShowDialog() != DialogResult.OK:
        print("No JSON selected. Exiting.")
        return

    with open(json_dialog.FileName, "r") as f:
        path_map = json.load(f)

    print("Loaded %d title entries from JSON." % len(path_map))

    # Step 3: Optionally select a replacement directory
    result = MessageBox.Show(
        "Do you want to use a different folder for the images?\n\n"
        "Yes = browse for folder (e.g. HD images)\n"
        "No = use original paths from JSON",
        "Image Source",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Question
    )

    replacement_dir = None
    if result == DialogResult.Yes:
        folder_dialog = FolderBrowserDialog()
        folder_dialog.Description = "Select folder containing replacement images"
        if folder_dialog.ShowDialog() == DialogResult.OK:
            replacement_dir = folder_dialog.SelectedPath
            print("Replacement directory: %s" % replacement_dir)
        else:
            print("No folder selected, using original paths.")

    # Step 4: Output path
    directory = os.path.dirname(input_doc)
    name_without_ext = os.path.splitext(os.path.basename(input_doc))[0]
    output_doc = os.path.join(directory, name_without_ext + "_reimaged.docx")

    print("Input:  %s" % input_doc)
    print("Output: %s" % output_doc)
    print("\nProcessing...")

    reinsert_images(input_doc, output_doc, path_map, replacement_dir)

if __name__ == "__main__":
    main()