using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace DocumentControlToolbar {
    class Tools {
        public static void SetStyle(String style) {
            //TODO confirm that style exists
            Word.Application app = Globals.ThisAddIn.Application;
            app.Selection.ParagraphFormat.set_Style(app.ActiveDocument.Styles[style]);
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
