using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace DocumentControlToolbar {
    class CrossReferenceWork {
        private void BookmarkInsertCrossReference() {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Application app = Globals.ThisAddIn.Application;

            int i = 0;

            ArrayList refsToSave = new ArrayList();

            foreach (Word.Paragraph paragraph in doc.Paragraphs) {
                Word.Style style = paragraph.get_Style() as Word.Style;
                string styleName = style.NameLocal;

                if (styleName == "Heading 1,2016_Überschrift 1,Headline 1") {
                    i++;
                } else if (styleName == "Heading 2,2016_Überschrift 2,Headline 2") {
                    i++;
                } else if (styleName == "Heading 3,2016_Überschrift 3,Headline 3") {
                    i++;
                    refsToSave.Add(i);
                }
            }

            Word.Table table = FindTable("Report Name");

            foreach (int j in refsToSave) {
                table.Rows.Add();

                table.Cell(table.Rows.Count, 1).Range.Select();
                insertCrossReference(app, Word.WdReferenceKind.wdContentText, j);

                //table.Cell(table.Rows.Count, 4).Range.Select();
                //table.Cell(table.Rows.Count, 4).Range.Delete();
                //insertCrossReference(app, Word.WdReferenceKind.wdPageNumber, j);
            }
        }

        private void insertCrossReference(Word.Application app, Word.WdReferenceKind kind, int j) {
            object ReferenceType = "Heading";
            object ReferenceItem = j;
            object InsertAsHyperlink = true;
            object IncludePosition = false;
            object SeparateNumbers = false;
            object SeparatorString = " ";

            app.Selection.InsertCrossReference(ReferenceType, kind, ReferenceItem,
                InsertAsHyperlink, IncludePosition, SeparateNumbers, SeparatorString);
        }

        /** Locates the acronym table in the document and returns it. **/
        private Word.Table FindTable(String tableName) {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            foreach (Word.Table table in doc.Tables) {
                String topLeft = table.Cell(1, 1).Range.Text;

                if (topLeft.Remove(topLeft.Length - 2).Equals(tableName)) {
                    return table;
                }
            }

            throw new AcronymTableNotFoundException("The " + tableName + " table could not be found.");
        }
    }
}
