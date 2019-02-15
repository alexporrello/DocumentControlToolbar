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


        public static void LoadNormalTemplate() {
            String appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            String docCont = Path.Combine(appData, "DocumentControl");
            String wordLoc = Path.Combine(docCont, "normal-template.dotm");

            if (File.Exists(wordLoc)) {
                Globals.ThisAddIn.Application.ActiveDocument.UpdateStylesOnOpen = true;
                Globals.ThisAddIn.Application.ActiveDocument.set_AttachedTemplate(wordLoc);
            } else {
                String errorText = "The Document Control Toolbar could not import the normal template " +
                    "into this word document.\nThis could have happened for a number of reasons, " +
                    "but the following are the most likely:\n\t1. The normal template has become corrupt. " + 
                    "Try running CopyWordlist again. This will download a fresh normal template to your computer.\n"+
                    "\t2. If you have not yet installed the CopyWordlist app onto your computer, then you haven't "+
                    "copied the normal template to your computer. Contact your Document Control Lead about installing "+
                    "and running this program on your computer.";

                Clipboard.SetData(DataFormats.Text, (Object)errorText);

                MessageBox.Show(    
                    "Failed to load the normal template. We have copied troubleshooting steps to the system clipboard.", "Error",
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
