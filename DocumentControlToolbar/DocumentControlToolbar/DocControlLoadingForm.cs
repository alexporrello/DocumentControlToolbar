using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace DocumentControlToolbar {
    public partial class DocControlLoadingForm : Form {
        public Action Worker { get; set; }

        public DocControlLoadingForm(Action worker, String mainText) {
            InitializeComponent();

            if (worker == null) {
                throw new ArgumentNullException();
            }

            SetMainText(mainText);

            Worker = worker;
        }

        protected override void OnLoad(EventArgs e) {
            base.OnLoad(e);
            Task.Factory.StartNew(Worker).ContinueWith(t => {
                this.Close();
            }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private void AcronymTableLoadingForm_Load(object sender, EventArgs e) {

        }

        delegate void SetTextCallback(string text);

        public void SetMainText(string text) {
            if (this.OperationNameField.InvokeRequired) {
                SetTextCallback d = new SetTextCallback(SetMainText);
                this.Invoke(d, new object[] { text });
            } else {
                this.OperationNameField.Text = text;
            }
        }

        public void SetNumberingUpdate(string text) {
            if (this.StatusField.InvokeRequired) {
                SetTextCallback d = new SetTextCallback(SetNumberingUpdate);
                this.Invoke(d, new object[] { text });
            } else {
                this.StatusField.Text = text;
            }
        }

        private void OperationNameField_Click(object sender, EventArgs e) {

        }

        private void DocControlLoadingForm_Load(object sender, EventArgs e) {

        }
    }
}
