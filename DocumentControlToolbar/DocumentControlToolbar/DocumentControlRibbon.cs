using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.Collections;

namespace DocumentControlToolbar {
    public partial class DocumentControlRibbon {

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e) { }

        /** ======================= General Group ======================= **/

        private void docPropUpdater_Click(object sender, RibbonControlEventArgs e) {
            new DocPropertiesEditor().Show();
        }

        //TODO Boilerplate Formatter

        /** ======================= Style Tools Group ======================= **/

        private void applyBodyStyle_Click(object sender, RibbonControlEventArgs e) {
            Tools.SetStyle("2016_Bodytext | 9pt");
        }

        private void keepWithNext_Click(object sender, RibbonControlEventArgs e) {
            Word.Application app = Globals.ThisAddIn.Application;
            app.Selection.ParagraphFormat.KeepWithNext = -1;
        }

        private void pageBreakBefore_Click(object sender, RibbonControlEventArgs e) {
            Word.Application app = Globals.ThisAddIn.Application;
            app.Selection.ParagraphFormat.PageBreakBefore = -1;
        }

        private void formatTable_Click(object sender, RibbonControlEventArgs e) {
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
        }

        private void insertSectionBreak_Click(object sender, RibbonControlEventArgs e) {
            Word.Application app = Globals.ThisAddIn.Application;
            app.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
        }

        /** ======================= List Tools Group ======================= **/

        private void defaultUL_Click(object sender, RibbonControlEventArgs e) {
            Tools.SetStyle("Body Text enumeration Arrow 2016 black");
        }

        private void defaultOL_Click(object sender, RibbonControlEventArgs e) {
            Tools.SetStyle("Body Text enumeration1 1. 2. 3.");
        }

        /** ======================= Acronym Table Group ======================= **/

        private void runAcronymTool_Click(object sender, RibbonControlEventArgs e) {
            new AcronymTableTool();
        }

        /** ======================= Cross-references Group ======================= **/

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

        private void updateAllFields_Click(object sender, RibbonControlEventArgs e) {
            Tools.UpdateAllFields();
        }

        /** ======================= Headings Dropdown ======================= **/

        private void headingOne_Click(object sender, RibbonControlEventArgs e) {
            Tools.SetStyle("Heading 1,2016_Überschrift 1,Headline 1");
        }

        private void headingTwo_Click(object sender, RibbonControlEventArgs e) {
            Tools.SetStyle("Heading 2,2016_Überschrift 2,Headline 2");
        }

        private void headingThree_Click(object sender, RibbonControlEventArgs e) {
            Tools.SetStyle("Heading 3,2016_Überschrift 3,Headline 3");
        }

        private void headingFour_Click(object sender, RibbonControlEventArgs e) {
            Tools.SetStyle("Heading 4,2016_Überschrift 4,Headline 4");
        }

        private void headingFive_Click(object sender, RibbonControlEventArgs e) {
            Tools.SetStyle("Heading 5,2016_Überschrift 5,Headline 5");
        }

        /** ======================= Ordered List Dropdown ======================= **/

        private void levelOneO_Click(object sender, RibbonControlEventArgs e) {
            Tools.SetStyle("Body Text enumeration1 1. 2. 3.");
        }

        private void levelTwoO_Click(object sender, RibbonControlEventArgs e) {
            Tools.SetStyle("Body Text enumeration2 a)b)c)");
        }

        /** ======================= Unordered List Dropdown ======================= **/

        private void levelOneU_Click(object sender, RibbonControlEventArgs e) {
            Tools.SetStyle("Body Text enumeration Arrow 2016 black");
        }

        private void levelTwoU_Click(object sender, RibbonControlEventArgs e) {
            Tools.SetStyle("Body Text enumeration Line1");
        }

        private void levelThreeU_Click(object sender, RibbonControlEventArgs e) {
            Tools.SetStyle("Body Text enumeration Line3");
        }

        private void levelFourU_Click(object sender, RibbonControlEventArgs e) {
            Tools.SetStyle("Body Text enumeration Point3");
        }

        private void updateWordlist_Click(object sender, RibbonControlEventArgs e) {
            WordList.DownloadAll();
        }

        private void updateDudsList_Click(object sender, RibbonControlEventArgs e) {
            WordList.DownloadDudsList();
        }

        private void excelCheckBox_Click(object sender, RibbonControlEventArgs e) {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e) {

        }
    }
}
