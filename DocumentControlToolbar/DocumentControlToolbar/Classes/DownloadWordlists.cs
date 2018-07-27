using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace DocumentControlToolbar {
    class WordList {

        public static String DudsListURL = 
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/acronym-duds.txt";

        public static String[] AllURLs = {
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/a.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/b.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/c.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/d.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/e.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/f.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/g.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/h.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/i.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/j.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/k.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/l.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/m.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/n.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/o.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/p.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/q.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/r.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/s.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/t.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/u.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/v.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/w.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/x.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/y.csv",
            "https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/z.csv"
        };

        public static String Folder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "DocumentControl");

        public static void DownloadAll() {
            Directory.CreateDirectory(WordList.Folder);

            foreach(String s in WordList.AllURLs) {
                using (var client = new WebClient()) {
                    client.DownloadFile(s, Path.Combine(WordList.Folder, s.Substring(s.Length - 5)));

                    if(!File.Exists(Path.Combine(WordList.Folder, s.Substring(s.Length - 5)))) {
                        throw new CouldNotDownloadFileException("The wordlists failed to download.");
                    }
                }
            }
        }

        public static void DownloadDudsList() {
            using (var client = new WebClient()) {
                String url = WordList.DudsListURL;
                client.DownloadFile(url, Path.Combine(WordList.Folder, "acronym-duds.txt"));

                if (!File.Exists(Path.Combine(WordList.Folder, "acronym-duds.txt"))) {
                    throw new CouldNotDownloadFileException("The duds list failed to download.");
                }
            }
        }
    }
}
