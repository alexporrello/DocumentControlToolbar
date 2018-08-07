namespace DocumentControlToolbar {
    partial class AcronymTableLoadingForm {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.acronymStatus = new System.Windows.Forms.Label();
            this.numberUpdate = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(48, 96);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(251, 23);
            this.progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.progressBar1.TabIndex = 0;
            this.progressBar1.UseWaitCursor = true;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(21, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(302, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Sit back and relax. Do something for yourself for a change.";
            this.label1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.label1.UseWaitCursor = true;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(21, 31);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(302, 21);
            this.label2.TabIndex = 2;
            this.label2.Text = "The Acronym Table Tool is working so you don\'t have to.";
            this.label2.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.label2.UseWaitCursor = true;
            // 
            // acronymStatus
            // 
            this.acronymStatus.Location = new System.Drawing.Point(12, 52);
            this.acronymStatus.Name = "acronymStatus";
            this.acronymStatus.Size = new System.Drawing.Size(321, 14);
            this.acronymStatus.TabIndex = 3;
            this.acronymStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // numberUpdate
            // 
            this.numberUpdate.Location = new System.Drawing.Point(12, 69);
            this.numberUpdate.Name = "numberUpdate";
            this.numberUpdate.Size = new System.Drawing.Size(321, 14);
            this.numberUpdate.TabIndex = 4;
            this.numberUpdate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // AcronymTableLoadingForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(345, 131);
            this.ControlBox = false;
            this.Controls.Add(this.numberUpdate);
            this.Controls.Add(this.acronymStatus);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.progressBar1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AcronymTableLoadingForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Running...";
            this.TopMost = true;
            this.UseWaitCursor = true;
            this.Load += new System.EventHandler(this.AcronymTableLoadingForm_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        public System.Windows.Forms.Label acronymStatus;
        public System.Windows.Forms.Label numberUpdate;
    }
}