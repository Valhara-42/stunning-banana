import clr
import os
import json
import shutil

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import OpenFileDialog, MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult
clr.AddReference("Microsoft.Office.Interop.Word")
import Microsoft.Office.Interop.Word as Word

def safe_print(text):
    print(str(text).encode('cp850', errors='replace').decode('cp850'))

def get_preceding_title(doc, table):
    """Get the text of the paragraph immediately before the table."""
    try:
        r = doc.Range(table.Range.Start, table.Range.Start)
        r.MoveStart(Word.WdUnits.wdParagraph, -1)
        text = r.Text.strip().replace("\r", "").replace("\x07", "").strip()
        if text:
            return text
        # If empty try one more paragraph back (in case of blank line)
        r.MoveStart(Word.WdUnits.wdParagraph, -1)
        text = r.Text.strip().replace("\r", "").replace("\x07", "").strip()
        return text if text else None
    except:
        pass
    return None

def replace_paths_with_images(doc_path, output_path, json_path):
    word_app = Word.ApplicationClass()
    word_app.Visible = False

    try:
        
        shutil.copy2(doc_path, output_path)
        doc = word_app.Documents.Open(output_path, ReadOnly=False)
        

        word_app.ScreenUpdating = False
        word_app.Options.CheckSpellingAsYouType = False
        word_app.Options.CheckGrammarAsYouType = False
        word_app.Options.Pagination = False

        print("Total tables: %d" % doc.Tables.Count)
        processed = 0

        # path_map: { "title text": [ { "row": X, "cell": Y, "path": "..." }, ... ] }
        # Using a list per title to support multiple image cells per table
        path_map = {}

        for table_idx in range(1, doc.Tables.Count + 1):
            table = doc.Tables.Item(table_idx)
            table_text = table.Range.Text

            if "\LCE" not in table_text and "\\\\" not in table_text:
                continue

            # Get the unique title preceding this table
            title = get_preceding_title(doc, table)
            if not title:
                title = "table_%d" % table_idx  # fallback if no title found
                print("  WARNING: No title found before table %d, using fallback key." % table_idx)
            else:
                safe_print("\nProcessing Table %d (title: %s)" % (table_idx, title[:60]))

            if title not in path_map:
                path_map[title] = []

            for row_idx in range(1, table.Rows.Count + 1):
                row = table.Rows.Item(row_idx)
                row_text = row.Range.Text

                if "\LCE" not in row_text and "\\\\" not in row_text:
                    continue

                for cell_idx in range(1, row.Cells.Count + 1):
                    try:
                        cell = row.Cells.Item(cell_idx)
                        cell_text = cell.Range.Text.strip()

                        if "\LCE" not in cell_text and not cell_text.startswith("\\\\"):
                            continue

                        img_path = cell_text.replace("\r", "").replace("\x07", "").strip()
                        safe_print("  Row %d, Cell %d: %s" % (row_idx, cell_idx, img_path))

                        if os.path.exists(img_path):
                            # Save to path map before replacing
                            path_map[title].append({
                                "row": row_idx,
                                "cell": cell_idx,
                                "path": img_path
                            })

                            cell.Range.Text = ""
                            try:
                                cell.Range.InlineShapes.AddPicture(
                                    FileName=img_path,
                                    LinkToFile=False,
                                    SaveWithDocument=True
                                )
                                print("  SUCCESS!")
                                processed += 1
                            except Exception as e:
                                safe_print("  ERROR inserting: %s" % str(e))
                        else:
                            safe_print("  ERROR: File not found: %s" % img_path)

                    except Exception as e:
                        safe_print("  Error in Row %d, Cell %d: %s" % (row_idx, cell_idx, str(e)))

            if table_idx % 20 == 0:
                doc.UndoClear()
                print("  [Cleared undo buffer at table %d]" % table_idx)

        print("\n\n=== COMPLETE ===")
        print("Processed %d images" % processed)

        word_app.Options.Pagination = True
        word_app.ScreenUpdating = True
        word_app.Options.CheckSpellingAsYouType = True
        word_app.Options.CheckGrammarAsYouType = True

        doc.Save()
        doc.Close()
        print("Saved: %s" % output_path)

        with open(json_path, "w") as f:
            json.dump(path_map, f, indent=2)
        print("Path map saved: %s" % json_path)

        MessageBox.Show(
            "Successfully processed %d images!\n\nSaved to:\n%s\n\nPath map saved to:\n%s" % (processed, output_path, json_path),
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
    dialog = OpenFileDialog()
    dialog.Title = "Select Word Document to Process"
    dialog.Filter = "Word Documents (*.docx)|*.docx|All Files (*.*)|*.*"
    dialog.FilterIndex = 1

    if dialog.ShowDialog() == DialogResult.OK:
        input_doc = dialog.FileName
        directory = os.path.dirname(input_doc)
        name_without_ext = os.path.splitext(os.path.basename(input_doc))[0]
        output_doc = os.path.join(directory, name_without_ext + "_images_added.docx")
        json_doc = os.path.join(directory, name_without_ext + "_image_paths.json")

        print("Input:  %s" % input_doc)
        print("Output: %s" % output_doc)
        print("Paths:  %s" % json_doc)
        print("\nProcessing...")

        replace_paths_with_images(input_doc, output_doc, json_doc)
    else:
        print("No file selected. Exiting.")

if __name__ == "__main__":
    main()