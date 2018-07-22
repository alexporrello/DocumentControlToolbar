using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace DocumentControlToolbar {
    public partial class DocumentControlRibbon {

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e) { 

        }

        private void docPropUpdater_Click(object sender, RibbonControlEventArgs e) {
            new DocPropertiesEditor().Show();
        }

        private void runAcronymTool_Click(object sender, RibbonControlEventArgs e) {
            new AcronymTableTool();
        }

        private void applyBodyStyle_Click(object sender, RibbonControlEventArgs e) {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Document    doc = Globals.ThisAddIn.Application.ActiveDocument;

            app.Selection.ParagraphFormat.set_Style(app.ActiveDocument.Styles["2016_Bodytext | 9pt"]);
        }

        private void tableRefButton_Click(object sender, RibbonControlEventArgs e) {
            Word.Application app = Globals.ThisAddIn.Application;
            app.Selection.InsertCaption("Table", "", "InsertCaption2", Word.WdCaptionPosition.wdCaptionPositionAbove, 0);
            app.Selection.ParagraphFormat.set_Style(app.ActiveDocument.Styles["2016_Marking"]);
        }

        private void figureRefButton_Click(object sender, RibbonControlEventArgs e) {
            Word.Application app = Globals.ThisAddIn.Application;
            app.Selection.InsertCaption("Figure", "", "InsertCaption2", Word.WdCaptionPosition.wdCaptionPositionAbove, 0);
            app.Selection.ParagraphFormat.set_Style(app.ActiveDocument.Styles["2016_Marking"]);
        }

        private void keepWithNext_Click(object sender, RibbonControlEventArgs e) {
            Word.Application app = Globals.ThisAddIn.Application;
            app.Selection.ParagraphFormat.KeepWithNext = -1;
        }
    }
}
