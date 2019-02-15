using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CopyWordlist {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
        }

        private void download_button_Click(object sender, EventArgs e) {
            DownloadWordlists();
        }

        private void DownloadWordlists() {
            String appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

            Boolean success = true;

            for (int i = 97; i <= 122; i++) {
                char character = (char)i;
                String text = character.ToString();
                if (CopyFileToAppData(appData, text) == 0) {
                    System.Windows.Forms.MessageBox.Show("The error log has been copied to your clipboard.", "Wordlists Download Failed");
                    success = false;
                    break;
                }
            }

            if (success && CopyFileToAppData(appData, "acronym-duds") == 0) {
                System.Windows.Forms.MessageBox.Show("The error log has been copied to your clipboard.", "Duds List Download Failed");
                success = false;
            } else {
                if (success && CopyFileToAppData(appData, "normal-template", ".docx") == 0) {
                    System.Windows.Forms.MessageBox.Show("The error log has been copied to your clipboard.", "Normal Template Download Failed");
                } else {
                    CopyFileToAppData(appData, "normal-template", ".dotm");
                    System.Windows.Forms.MessageBox.Show("The copy operation has completed successfully.", "Success!");
                }
            }
        }

        private int CopyFileToAppData(String appdata, String text) {
            return CopyFileToAppData(appdata, text, ".csv");
        }

        private int CopyFileToAppData(String appData, String text, String extension) {
            if (!Directory.Exists(Path.Combine(appData, "DocumentControl"))) {
                Directory.CreateDirectory(Path.Combine(appData, "DocumentControl"));
            }

            if (File.Exists(Path.Combine(appData, Path.Combine("DocumentControl", text + extension)))) {
                File.Delete(Path.Combine(appData, Path.Combine("DocumentControl", text + extension)));
            }

            try {
                File.Copy(@"Resources/" + text + extension, Path.Combine(appData, Path.Combine("DocumentControl", text + extension)));
            } catch(Exception e) {
                Clipboard.SetData(DataFormats.Text, e.StackTrace);
                return 0;
            }

            return 1;
        }
    }
}
