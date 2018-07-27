using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Collections;
using System.Net;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Drawing;


namespace DocumentControlToolbar {

    class AcronymTableTool {
        private Word.Application app = Globals.ThisAddIn.Application;
        private Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

        private HashSet<String> AllAcronyms = new HashSet<String>();
        private HashSet<String> foundAcronyms = new HashSet<String>();

        private ArrayList acronymsInTable = new ArrayList();

        public AcronymTableTool() { }

        public void start() {
            try {
                Word.Table acronymTable = FindAcronymTable();

                CheckAcronymsInTable(acronymTable);
                GetAllAcronymsInDocument();

                AddFoundAcronymsToTable(acronymTable);

                Debug.Print("Done");
            } catch (CustomExceptions) {
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
                    if (topLeft.Remove(topLeft.Length - 2).Trim().Equals(s)) {
                        return table;
                    }
                }
            }

            throw new CustomExceptions("The acronym table could not be found.");
        }

        /** Checks if the acronyms in the table appear in the rest of the document. **/
        private void CheckAcronymsInTable(Word.Table acronymTable) {
            for (int i = 2; i <= acronymTable.Rows.Count; i++) {
                Word.Cell leftCell = acronymTable.Cell(i, 1);
                Word.Cell rightCell = acronymTable.Cell(i, 2);

                String acronym = leftCell.Range.Text;
                acronym = acronym.Remove(acronym.Length - 2);
                SearchForEntry(leftCell, acronym);

                String definition = rightCell.Range.Text;
                definition = definition.Remove(definition.Length - 2);
                SearchForEntry(rightCell, definition);

                acronymsInTable.Add(acronym);
            }
        }

        /** Searches if this cell's acronym appears in the document **/
        private Boolean SearchForEntry(Word.Cell thisCell, String text) {
            Boolean found = true;

            thisCell.Range.Text = "";

            if (!Find(text, false)) {
                thisCell.Shading.ForegroundPatternColorIndex = Word.WdColorIndex.wdRed;
                found = false;
            }

            thisCell.Range.Text = text;

            return found;
        }

        /** Searches through the active document for a string. Returns true if found; else, false. **/
        private Boolean Find(String text, Boolean matchCase) {
            Word.Find thisFind = doc.Content.Find;
            thisFind.Text = text;
            thisFind.Format = false;
            thisFind.Wrap = Word.WdFindWrap.wdFindContinue;
            thisFind.MatchCase = matchCase;
            thisFind.MatchWholeWord = true;

            return thisFind.Execute();
        }

        /** Counts how many times a collection contains an object **/
        private int NumberOfInstances(ArrayList collection, String query) {
            int toReturn = 0;

            foreach (String s in collection) {
                if (s.Contains(query)) {
                    toReturn += 1;
                }
            }

            return toReturn;
        }


        /** Searches through the document for words it thinks might be an acronym. **/
        private void GetAllAcronymsInDocument() {
            foreach (Word.Range word in doc.Words) {
                if (IsValidWordFirstCheck(word.Text)) {
                    if (!app.CheckSpelling(word.Text.ToLower())) {
                        foundAcronyms.Add(word.Text);
                    }
                }
            }
        }

        /** The first check to determine if a given string is a valid acronym. **/
        private Boolean IsValidWordFirstCheck(String s) {
            //TODO this will not work for things like '3G'
            return s != null && s.Trim().Length > 1 && s.Equals(s.ToUpper()) && Regex.IsMatch(s, @"^[a-zA-Z]+$");
        }

        /** Adds all found acronyms (that are not already in the table) to the table; then, sort. **/
        private void AddFoundAcronymsToTable(Word.Table acronymTable) {
            String dudsList = GetDudsList();

            foreach (String word in foundAcronyms) {
                String definition = "";

                if (!acronymsInTable.Contains(word) && !dudsList.Contains(word)) {
                    acronymTable.Rows.Add();

                    String wordList = DownloadWordlist(word.ToLower().Substring(0, 1));

                    if (wordList.Contains(word)) {
                        foreach (string s in wordList.Split('\n')) {
                            String[] split = s.Split(',');

                            if (split[0].Equals(word)) {
                                definition = split[1];
                            }
                        }
                    }

                    Word.Cell defCell = acronymTable.Cell(acronymTable.Rows.Count, 2);
                    defCell.Shading.ForegroundPatternColorIndex = Word.WdColorIndex.wdYellow;
                    defCell.Range.Text = definition;

                    Word.Cell acronymCell = acronymTable.Cell(acronymTable.Rows.Count, 1);
                    acronymCell.Shading.ForegroundPatternColorIndex = Word.WdColorIndex.wdYellow;
                    acronymCell.Range.Text = word;
                }
            }

            acronymTable.SortAscending();
        }

        /** Downloads a wordlist from the online database **/
        private String DownloadWordlist(String beginningLetter) {
            String wordList = Path.Combine(WordList.Folder, beginningLetter + ".csv");

            if (!File.Exists(wordList)) {
                WordList.DownloadAll();
            }

            return System.IO.File.ReadAllText(wordList);
        }

        /** Downloads a wordlist from the online database **/
        private String GetDudsList() {
            String acronymDuds = Path.Combine(WordList.Folder, "acronym-duds.txt");

            if (!File.Exists(acronymDuds)) {
                WordList.DownloadDudsList();
            }

            return System.IO.File.ReadAllText(acronymDuds);
        }
    }
}
