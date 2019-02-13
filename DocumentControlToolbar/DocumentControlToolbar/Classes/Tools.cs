using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace DocumentControlToolbar {
    class Tools {

        public static void FormatTable() {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Table table = app.Selection.Range.Tables[1];

            for (int row = 1; row <= table.Rows.Count; row++) {
                for (int col = 1; col <= table.Columns.Count; col++) {
                    try {
                        table.Cell(row, col).Range.Select();
                        app.Selection.ClearFormatting();

                        if (row == 1) {
                            table.Cell(row, col).Range.set_Style(app.ActiveDocument.Styles["2016_TableHeader | 10pt bold"]);
                        } else {
                            table.Cell(row, col).Range.set_Style(app.ActiveDocument.Styles["2016_Table | 9pt"]);
                        }
                    } catch (Exception f) {
                        Debug.Print(f.Message);
                    }
                }
            }

            table.set_Style(app.ActiveDocument.Styles["MasterTable"]);
            table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);
            table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);

            try {
                table.Cell(1, 1).Row.HeadingFormat = (int)Word.WdConstants.wdToggle;
                table.ApplyStyleHeadingRows = true;
            } catch (Exception) { };
        }

        public static String LocateFile(String title) {
            OpenFileDialog file = new OpenFileDialog();
            file.Title = title;

            if (file.ShowDialog() == DialogResult.OK) {
                return file.FileName;
            }

            throw new Exception("The user did not select a file.");
        }

        public static void LoadNormalTemplate(String url) {
            if (File.Exists(url)) {
                Globals.ThisAddIn.Application.ActiveDocument.CopyStylesFromTemplate(url);
            } else {
                MessageBox.Show(
                    "Failed to load the normal template. Please try again.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void StartWait() {
            Globals.ThisAddIn.Application.System.Cursor = Word.WdCursorType.wdCursorWait;
            Globals.ThisAddIn.Application.Application.ScreenUpdating = false;
        }

        public static void EndWait() {
            Globals.ThisAddIn.Application.Application.ScreenUpdating = true;
            Globals.ThisAddIn.Application.System.Cursor = Word.WdCursorType.wdCursorNormal;
        }

        public static void SetStyle(String style) {
            try {
                Word.Application app = Globals.ThisAddIn.Application;
                app.Selection.ParagraphFormat.set_Style(app.ActiveDocument.Styles[style]);
            } catch(Exception) {
                MessageBox.Show(
                    "The style '" + style + "' does not exist. " +
                    "Please import the normal template (Doc Control >> Import Styles) and try again.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
