namespace DocumentControlToolbar
{
    partial class DocumentControlRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public DocumentControlRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.DocControl = this.Factory.CreateRibbonTab();
            this.acronymTableGroup = this.Factory.CreateRibbonGroup();
            this.DocControl.SuspendLayout();
            this.SuspendLayout();
            // 
            // DocControl
            // 
            this.DocControl.Groups.Add(this.acronymTableGroup);
            this.DocControl.Label = "Doc Control";
            this.DocControl.Name = "DocControl";
            // 
            // acronymTableGroup
            // 
            this.acronymTableGroup.Label = "Acronym Table";
            this.acronymTableGroup.Name = "acronymTableGroup";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.DocControl);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.DocControl.ResumeLayout(false);
            this.DocControl.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup acronymTableGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab DocControl;
    }

    partial class ThisRibbonCollection
    {
        internal DocumentControlRibbon Ribbon1
        {
            get { return this.GetRibbon<DocumentControlRibbon>(); }
        }
    }
}
