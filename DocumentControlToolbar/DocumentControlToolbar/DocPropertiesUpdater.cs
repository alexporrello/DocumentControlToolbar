using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace DocumentControlToolbar {
    public partial class DocPropertiesEditor : Form {
        public DocPropertiesEditor() {
            InitializeComponent();
        }

        private void DocPropertiesUpdater_Load(object sender, EventArgs e) {

        }

        private void label1_Click(object sender, EventArgs e) {

        }

        private void label2_Click(object sender, EventArgs e) {

        }

        private void panel1_Paint(object sender, PaintEventArgs e) {

        }

        private void button1_Click(object sender, EventArgs e) {
            UpdateDocumentProperty("DocTitle", this.title);
            UpdateDocumentProperty("DocAcronym", this.acronym);
            UpdateDocumentProperty("DocNumber", this.sharePointID);
            UpdateDocumentProperty("DocReleaseDate", this.releaseDate);
            UpdateDocumentProperty("DocVersion", this.version);
            UpdateDocumentProperty("DocStatus", this.status);
            UpdateDocumentProperty("Author", this.author);
            UpdateDocumentProperty("ProjectManager", this.pm);

            UpdateDocumentProperty("RoadName", this.roadName);
            UpdateDocumentProperty("SolutionType", this.solutionType);
            UpdateDocumentProperty("SolutionAcronym", this.solutionAcronym);
            UpdateDocumentProperty("ClientAcronym", this.clientAcronym);
            UpdateDocumentProperty("Client", this.client);

            Tools.UpdateAllFields();
        }

        private void demoButton_Click(object sender, EventArgs e) {
            this.title.Text = "The Big Document";
            this.acronym.Text = "TBD";
            this.sharePointID.Text = "54004";
            this.releaseDate.Text = "01/01/01";
            this.version.Text = "v1.957";
            this.status.Text = "Draft";
            this.author.Text = "Bobby Flay";
            this.pm.Text = "Guy Fieri";
            this.roadName.Text = "The Big Road";
            this.solutionType.Text = "The Big Solution";
            this.solutionAcronym.Text = "TBS";
            this.client.Text = "The Big Client";
            this.clientAcronym.Text = "TBC";
        }

        private void UpdateDocumentProperty(String property, TextBox value) {
            if (!value.Text.Equals("")) {
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

                Microsoft.Office.Core.DocumentProperties properties;
                properties = (Office.DocumentProperties)doc.CustomDocumentProperties;

                if (ReadDocumentProperty(property) != null) {
                    properties[property].Delete();
                }

                properties.Add(property, false,
                    Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString,
                    value.Text);
            }
        }

        private string ReadDocumentProperty(string propertyName) {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            Office.DocumentProperties properties;
            properties = (Office.DocumentProperties)doc.CustomDocumentProperties;

            foreach (Office.DocumentProperty prop in properties) {
                if (prop.Name == propertyName) {
                    return prop.Value.ToString();
                }
            }
            return null;
        }
    }
}
