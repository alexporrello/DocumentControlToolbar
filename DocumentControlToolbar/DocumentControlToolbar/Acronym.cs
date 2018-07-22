using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Collections;

namespace DocumentControlToolbar {

    class AcronymTableTool {

        private Word.Application app = Globals.ThisAddIn.Application;
        private Word.Document    doc = Globals.ThisAddIn.Application.ActiveDocument;
        
        private HashSet<String> AllAcronyms   = new HashSet<String>();
        private HashSet<String> foundAcronyms = new HashSet<String>();

        private ArrayList acronymsInTable = new ArrayList();
        

        public AcronymTableTool() {
            try {
                Word.Table acronymTable = FindAcronymTable();

                CheckAcronymsInTable(acronymTable);
                GetAllAcronymsInDocument();

                foreach(String word in foundAcronyms) {
                    if (!acronymsInTable.Contains(word)) {
                        acronymTable.Rows.Add();

                        Word.Cell thisCell = acronymTable.Cell(acronymTable.Rows.Count, 1);

                        thisCell.Range.Text = word;
                        thisCell.Shading.ForegroundPatternColorIndex = Word.WdColorIndex.wdYellow;
                    }
                }

            } catch (AcronymTableNotFoundException) {
                MessageBox.Show("The acronym table could not be found.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /** Locates the acronym table in the document and returns it. **/
        private Word.Table FindAcronymTable() {
            Word.Range word = doc.Words[1];

            foreach (Word.Table table in doc.Tables) {
                String[] possibilities = { "Abbreviations", "Abbreviation", "Acronym", "Acronyms" };

                String topLeft = table.Cell(1, 1).Range.Text;

                foreach (String s in possibilities) {
                    if (topLeft.Remove(topLeft.Length-2).Equals(s)) {
                        return table;
                    }
                }
            }

            throw new AcronymTableNotFoundException("The acronym table could not be found.");
        }

        /** Checks if the acronyms in the table appear in the rest of the document. **/
        private void CheckAcronymsInTable(Word.Table acronymTable) {
            ArrayList allWords = new ArrayList();
            foreach(Word.Range word in doc.Words) {
                if(word.Text.Trim().Length > 1) {
                    allWords.Add(word.Text.Trim());
                }
            }

            for (int i = 2; i <= acronymTable.Rows.Count; i++) {
                Word.Cell thisCell = acronymTable.Cell(i, 1);

                String text = thisCell.Range.Text;
                text = text.Remove(text.Length - 2);

                acronymsInTable.Add(text);

                if(NumberOfInstances(allWords, text) == 1) {
                    thisCell.Select();
                    thisCell.Shading.ForegroundPatternColorIndex = Word.WdColorIndex.wdRed;
                }
            }
        }

        /** Counts how many times a collection contains an object **/
        private int NumberOfInstances(ArrayList collection, String query) {
            int toReturn = 0;

            foreach(String s in collection) {
                if(s.Contains(query)) {
                    toReturn += 1;
                }
            }

            return toReturn;
        }


        /** Searches through the document for words it thinks might be an acronym. **/
        private void GetAllAcronymsInDocument() {
            foreach (Word.Range word in doc.Words) {
                if (IsValidWordFirstCheck(word.Text)) {
                    foundAcronyms.Add(word.Text);
                }
            }
        }

        /** The first check to determine if a given string is a valid acronym. **/
        private Boolean IsValidWordFirstCheck(String s) {
            return s.Trim().Length > 1 && s.Equals(s.ToUpper()) && Regex.IsMatch(s, @"^[a-zA-Z]+$");
        }
    }
}
