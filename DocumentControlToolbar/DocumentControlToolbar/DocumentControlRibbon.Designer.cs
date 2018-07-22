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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DocumentControlRibbon));
            this.DocControl = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.textToolsGroup = this.Factory.CreateRibbonGroup();
            this.acronymTableGroup = this.Factory.CreateRibbonGroup();
            this.crossRefsGroup = this.Factory.CreateRibbonGroup();
            this.docPropUpdater = this.Factory.CreateRibbonButton();
            this.boilerplateFormat = this.Factory.CreateRibbonButton();
            this.headingsDropdown = this.Factory.CreateRibbonGallery();
            this.h1 = this.Factory.CreateRibbonButton();
            this.h2 = this.Factory.CreateRibbonButton();
            this.h3 = this.Factory.CreateRibbonButton();
            this.h4 = this.Factory.CreateRibbonButton();
            this.h5 = this.Factory.CreateRibbonButton();
            this.applyBodyStyle = this.Factory.CreateRibbonButton();
            this.keepWithNext = this.Factory.CreateRibbonButton();
            this.runAcronymTool = this.Factory.CreateRibbonButton();
            this.figureRefButton = this.Factory.CreateRibbonButton();
            this.tableRefButton = this.Factory.CreateRibbonButton();
            this.updateAllFields = this.Factory.CreateRibbonButton();
            this.updateWordlist = this.Factory.CreateRibbonButton();
            this.updateDudsList = this.Factory.CreateRibbonButton();
            this.DocControl.SuspendLayout();
            this.group1.SuspendLayout();
            this.textToolsGroup.SuspendLayout();
            this.acronymTableGroup.SuspendLayout();
            this.crossRefsGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // DocControl
            // 
            this.DocControl.Groups.Add(this.group1);
            this.DocControl.Groups.Add(this.textToolsGroup);
            this.DocControl.Groups.Add(this.acronymTableGroup);
            this.DocControl.Groups.Add(this.crossRefsGroup);
            this.DocControl.Label = "DOC CONTROL";
            this.DocControl.Name = "DocControl";
            // 
            // group1
            // 
            this.group1.Items.Add(this.docPropUpdater);
            this.group1.Items.Add(this.boilerplateFormat);
            this.group1.Label = "General";
            this.group1.Name = "group1";
            // 
            // textToolsGroup
            // 
            this.textToolsGroup.Items.Add(this.headingsDropdown);
            this.textToolsGroup.Items.Add(this.applyBodyStyle);
            this.textToolsGroup.Items.Add(this.keepWithNext);
            this.textToolsGroup.Label = "Text Tools";
            this.textToolsGroup.Name = "textToolsGroup";
            // 
            // acronymTableGroup
            // 
            this.acronymTableGroup.Items.Add(this.runAcronymTool);
            this.acronymTableGroup.Items.Add(this.updateWordlist);
            this.acronymTableGroup.Items.Add(this.updateDudsList);
            this.acronymTableGroup.Label = "Acronym Table";
            this.acronymTableGroup.Name = "acronymTableGroup";
            // 
            // crossRefsGroup
            // 
            this.crossRefsGroup.Items.Add(this.figureRefButton);
            this.crossRefsGroup.Items.Add(this.tableRefButton);
            this.crossRefsGroup.Items.Add(this.updateAllFields);
            this.crossRefsGroup.Label = "Cross-references";
            this.crossRefsGroup.Name = "crossRefsGroup";
            // 
            // docPropUpdater
            // 
            this.docPropUpdater.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.docPropUpdater.Image = global::DocumentControlToolbar.Properties.Resources.properties_icon_raw;
            this.docPropUpdater.Label = "Document Properties Editor";
            this.docPropUpdater.Name = "docPropUpdater";
            this.docPropUpdater.ScreenTip = "Opens a dialog by which users can easily update a document\'s metadata.";
            this.docPropUpdater.ShowImage = true;
            this.docPropUpdater.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.docPropUpdater_Click);
            // 
            // boilerplateFormat
            // 
            this.boilerplateFormat.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.boilerplateFormat.Image = global::DocumentControlToolbar.Properties.Resources.format_icon_raw;
            this.boilerplateFormat.Label = "Boilerplate Formatter";
            this.boilerplateFormat.Name = "boilerplateFormat";
            this.boilerplateFormat.ScreenTip = "Auto-formats boilerplate documents downloaded from our internal wiki.";
            this.boilerplateFormat.ShowImage = true;
            // 
            // headingsDropdown
            // 
            this.headingsDropdown.Buttons.Add(this.h1);
            this.headingsDropdown.Buttons.Add(this.h2);
            this.headingsDropdown.Buttons.Add(this.h3);
            this.headingsDropdown.Buttons.Add(this.h4);
            this.headingsDropdown.Buttons.Add(this.h5);
            this.headingsDropdown.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.headingsDropdown.Image = global::DocumentControlToolbar.Properties.Resources.headings_icon;
            this.headingsDropdown.Label = "Apply Headings";
            this.headingsDropdown.Name = "headingsDropdown";
            this.headingsDropdown.ShowImage = true;
            // 
            // h1
            // 
            this.h1.Label = "Level 1";
            this.h1.Name = "h1";
            // 
            // h2
            // 
            this.h2.Label = "Level 2";
            this.h2.Name = "h2";
            // 
            // h3
            // 
            this.h3.Label = "Level 3";
            this.h3.Name = "h3";
            // 
            // h4
            // 
            this.h4.Label = "Level 4";
            this.h4.Name = "h4";
            // 
            // h5
            // 
            this.h5.Label = "Level 5";
            this.h5.Name = "h5";
            // 
            // applyBodyStyle
            // 
            this.applyBodyStyle.Image = global::DocumentControlToolbar.Properties.Resources.apply_style_small_icon;
            this.applyBodyStyle.Label = " Apply Body Style ";
            this.applyBodyStyle.Name = "applyBodyStyle";
            this.applyBodyStyle.ShowImage = true;
            this.applyBodyStyle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.applyBodyStyle_Click);
            // 
            // keepWithNext
            // 
            this.keepWithNext.Image = global::DocumentControlToolbar.Properties.Resources.apply_style_small_icon;
            this.keepWithNext.Label = " Keep With Next ";
            this.keepWithNext.Name = "keepWithNext";
            this.keepWithNext.ShowImage = true;
            this.keepWithNext.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.keepWithNext_Click);
            // 
            // runAcronymTool
            // 
            this.runAcronymTool.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.runAcronymTool.Image = ((System.Drawing.Image)(resources.GetObject("runAcronymTool.Image")));
            this.runAcronymTool.Label = "Run Updater";
            this.runAcronymTool.Name = "runAcronymTool";
            this.runAcronymTool.ShowImage = true;
            this.runAcronymTool.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.runAcronymTool_Click);
            // 
            // figureRefButton
            // 
            this.figureRefButton.Image = global::DocumentControlToolbar.Properties.Resources.insert_ref_small_icon;
            this.figureRefButton.Label = " Insert Figure Ref ";
            this.figureRefButton.Name = "figureRefButton";
            this.figureRefButton.ShowImage = true;
            this.figureRefButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.figureRefButton_Click);
            // 
            // tableRefButton
            // 
            this.tableRefButton.Image = global::DocumentControlToolbar.Properties.Resources.insert_ref_small_icon;
            this.tableRefButton.Label = " Insert Table Ref ";
            this.tableRefButton.Name = "tableRefButton";
            this.tableRefButton.ShowImage = true;
            this.tableRefButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.tableRefButton_Click);
            // 
            // updateAllFields
            // 
            this.updateAllFields.Image = global::DocumentControlToolbar.Properties.Resources.update_small_icon;
            this.updateAllFields.Label = "Update All Fields";
            this.updateAllFields.Name = "updateAllFields";
            this.updateAllFields.ShowImage = true;
            // 
            // updateWordlist
            // 
            this.updateWordlist.Image = global::DocumentControlToolbar.Properties.Resources.update_small_icon;
            this.updateWordlist.Label = " Update Wordlist ";
            this.updateWordlist.Name = "updateWordlist";
            this.updateWordlist.ShowImage = true;
            // 
            // updateDudsList
            // 
            this.updateDudsList.Image = global::DocumentControlToolbar.Properties.Resources.update_small_icon;
            this.updateDudsList.Label = " Update Duds List ";
            this.updateDudsList.Name = "updateDudsList";
            this.updateDudsList.ShowImage = true;
            // 
            // DocumentControlRibbon
            // 
            this.Name = "DocumentControlRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.DocControl);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.DocControl.ResumeLayout(false);
            this.DocControl.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.textToolsGroup.ResumeLayout(false);
            this.textToolsGroup.PerformLayout();
            this.acronymTableGroup.ResumeLayout(false);
            this.acronymTableGroup.PerformLayout();
            this.crossRefsGroup.ResumeLayout(false);
            this.crossRefsGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup acronymTableGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab DocControl;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup textToolsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton applyBodyStyle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton keepWithNext;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup crossRefsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton figureRefButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton tableRefButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton docPropUpdater;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton boilerplateFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton runAcronymTool;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery headingsDropdown;
        private Microsoft.Office.Tools.Ribbon.RibbonButton h1;
        private Microsoft.Office.Tools.Ribbon.RibbonButton h2;
        private Microsoft.Office.Tools.Ribbon.RibbonButton h3;
        private Microsoft.Office.Tools.Ribbon.RibbonButton h4;
        private Microsoft.Office.Tools.Ribbon.RibbonButton h5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton updateAllFields;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton updateWordlist;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton updateDudsList;
    }

    partial class ThisRibbonCollection
    {
        internal DocumentControlRibbon Ribbon1
        {
            get { return this.GetRibbon<DocumentControlRibbon>(); }
        }
    }
}
