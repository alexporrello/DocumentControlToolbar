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

            for (int i = 97; i <= 122; i++) {
                char character = (char) i;
                String text = character.ToString();
                CopyFileToAppData(appData, text);
            }
        }

        private void CopyFileToAppData(String appData, String text) {
            File.Copy(@"Resources/" + text + ".csv", Path.Combine(appData, Path.Combine("DocumentControl", text + ".csv")));
        }
    }
}
