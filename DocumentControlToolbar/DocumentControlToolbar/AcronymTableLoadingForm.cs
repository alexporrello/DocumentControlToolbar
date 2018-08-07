using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocumentControlToolbar {
    public partial class AcronymTableLoadingForm : Form {

        public Action Worker { get; set; }

        public AcronymTableLoadingForm(Action worker) {
            InitializeComponent();

            if(worker == null) {
                throw new ArgumentNullException();
            }

            Worker = worker;
        }

        protected override void OnLoad(EventArgs e) {
            base.OnLoad(e);
            Task.Factory.StartNew(Worker).ContinueWith(t=> {
                this.Close();
            }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private void AcronymTableLoadingForm_Load(object sender, EventArgs e) {

        }

        delegate void SetTextCallback(string text);

        public void SetMainText(string text) {
            if (this.acronymStatus.InvokeRequired) {
                SetTextCallback d = new SetTextCallback(SetMainText);
                this.Invoke(d, new object[] { text });
            } else {
                this.acronymStatus.Text = text;
            }
        }

        public void SetNumberingUpdate(string text) {
            if (this.numberUpdate.InvokeRequired) {
                SetTextCallback d = new SetTextCallback(SetNumberingUpdate);
                this.Invoke(d, new object[] { text });
            } else {
                this.numberUpdate.Text = text;
            }
        }
    }
}
