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
        public static void SetStyle(String style) {
            try {
                Word.Application app = Globals.ThisAddIn.Application;
                app.Selection.ParagraphFormat.set_Style(app.ActiveDocument.Styles[style]);
            } catch(Exception) {
                MessageBox.Show(
                    "The style '" + style + "' does not exist. " +
                    "Please import the latest company styles (Doc Control >> Import Styles) and try again.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void DownloadTemplateTo(String downloadTo) {
            var fileName = Guid.NewGuid().ToString() + ".bat";
            var batchPath = Path.Combine(Environment.GetEnvironmentVariable("temp"), fileName);

            var batchCode = "bitsadmin.exe /transfer \"Install Macros\" " +
                "http://github.com/alexporrello/TWBoilerplateMacros/raw/master/binaries/Normal.dotm " +
                downloadTo;

            File.WriteAllText(batchPath, batchCode);

            Process.Start(batchPath).WaitForExit();

            File.Delete(batchPath);
        }

        public static void UpdateAllFields() {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            foreach (Word.Section section in doc.Sections) {
                doc.Fields.Update();

                Word.HeadersFooters headers = section.Headers;  //Get all headers
                foreach (Word.HeaderFooter header in headers) {
                    Word.Fields fields = header.Range.Fields;
                    foreach (Word.Field field in fields) {
                        field.Update();  // update all fields in headers
                    }
                }

                Word.HeadersFooters footers = section.Footers;  //Get all footers
                foreach (Word.HeaderFooter footer in footers) {
                    Word.Fields fields = footer.Range.Fields;
                    foreach (Word.Field field in fields) {
                        field.Update();  //update all fields in footers
                    }
                }
            }
        }
    }
}
