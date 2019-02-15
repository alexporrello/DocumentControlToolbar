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
                    System.Windows.Forms.MessageBox.Show("The error log has been copied to your clipboard.", "Download Failed");
                    success = false;
                    break;
                }
            }

            if (success && CopyFileToAppData(appData, "acronym-duds") == 0) {
                System.Windows.Forms.MessageBox.Show("The error log has been copied to your clipboard.", "Download Failed");
            } else {
                System.Windows.Forms.MessageBox.Show("The copy operation has completed successfully.", "Success!");
            }
        }

        private int CopyFileToAppData(String appData, String text) {
            if (!Directory.Exists(Path.Combine(appData, "DocumentControl"))) {
                Directory.CreateDirectory(Path.Combine(appData, "DocumentControl"));
            }

            if (File.Exists(Path.Combine(appData, Path.Combine("DocumentControl", text + ".csv")))) {
                File.Delete(Path.Combine(appData, Path.Combine("DocumentControl", text + ".csv")));
            }

            try {
                File.Copy(@"Resources/" + text + ".csv", Path.Combine(appData, Path.Combine("DocumentControl", text + ".csv")));
            } catch(Exception e) {
                Clipboard.SetData(DataFormats.Text, e.StackTrace);
                return 0;
            }

            return 1;
        }
    }
}
