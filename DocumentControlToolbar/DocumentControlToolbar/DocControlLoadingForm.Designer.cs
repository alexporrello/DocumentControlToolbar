namespace DocumentControlToolbar {
    partial class DocControlLoadingForm {
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
            this.OperationNameField = new System.Windows.Forms.Label();
            this.StatusField = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(40, 67);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(251, 23);
            this.progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.progressBar1.TabIndex = 1;
            this.progressBar1.UseWaitCursor = true;
            // 
            // OperationNameField
            // 
            this.OperationNameField.Location = new System.Drawing.Point(12, 9);
            this.OperationNameField.Name = "OperationNameField";
            this.OperationNameField.Size = new System.Drawing.Size(321, 23);
            this.OperationNameField.TabIndex = 2;
            this.OperationNameField.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.OperationNameField.UseWaitCursor = true;
            this.OperationNameField.Click += new System.EventHandler(this.OperationNameField_Click);
            // 
            // StatusField
            // 
            this.StatusField.Location = new System.Drawing.Point(12, 32);
            this.StatusField.Name = "StatusField";
            this.StatusField.Size = new System.Drawing.Size(321, 23);
            this.StatusField.TabIndex = 3;
            this.StatusField.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.StatusField.UseWaitCursor = true;
            // 
            // DocControlLoadingForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(345, 111);
            this.ControlBox = false;
            this.Controls.Add(this.StatusField);
            this.Controls.Add(this.OperationNameField);
            this.Controls.Add(this.progressBar1);
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "DocControlLoadingForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Loading...";
            this.UseWaitCursor = true;
            this.Load += new System.EventHandler(this.DocControlLoadingForm_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ProgressBar progressBar1;
        public System.Windows.Forms.Label OperationNameField;
        public System.Windows.Forms.Label StatusField;
    }
}