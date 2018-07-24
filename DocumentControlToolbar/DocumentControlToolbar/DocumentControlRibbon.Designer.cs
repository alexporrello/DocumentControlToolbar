﻿namespace DocumentControlToolbar
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
            this.group2 = this.Factory.CreateRibbonGroup();
            this.acronymTableGroup = this.Factory.CreateRibbonGroup();
            this.crossRefsGroup = this.Factory.CreateRibbonGroup();
            this.docPropUpdater = this.Factory.CreateRibbonButton();
            this.boilerplateFormat = this.Factory.CreateRibbonButton();
            this.headingsDropdown = this.Factory.CreateRibbonGallery();
            this.headingOne = this.Factory.CreateRibbonButton();
            this.headingTwo = this.Factory.CreateRibbonButton();
            this.headingThree = this.Factory.CreateRibbonButton();
            this.headingFour = this.Factory.CreateRibbonButton();
            this.headingFive = this.Factory.CreateRibbonButton();
            this.applyBodyStyle = this.Factory.CreateRibbonButton();
            this.keepWithNext = this.Factory.CreateRibbonButton();
            this.gallery2 = this.Factory.CreateRibbonGallery();
            this.levelOneU = this.Factory.CreateRibbonButton();
            this.levelTwoU = this.Factory.CreateRibbonButton();
            this.levelThreeU = this.Factory.CreateRibbonButton();
            this.levelFourU = this.Factory.CreateRibbonButton();
            this.gallery1 = this.Factory.CreateRibbonGallery();
            this.levelOneO = this.Factory.CreateRibbonButton();
            this.levelTwoO = this.Factory.CreateRibbonButton();
            this.defaultUL = this.Factory.CreateRibbonButton();
            this.defaultOL = this.Factory.CreateRibbonButton();
            this.runAcronymTool = this.Factory.CreateRibbonButton();
            this.updateWordlist = this.Factory.CreateRibbonButton();
            this.updateDudsList = this.Factory.CreateRibbonButton();
            this.figureRefButton = this.Factory.CreateRibbonButton();
            this.tableRefButton = this.Factory.CreateRibbonButton();
            this.updateAllFields = this.Factory.CreateRibbonButton();
            this.DocControl.SuspendLayout();
            this.group1.SuspendLayout();
            this.textToolsGroup.SuspendLayout();
            this.group2.SuspendLayout();
            this.acronymTableGroup.SuspendLayout();
            this.crossRefsGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // DocControl
            // 
            this.DocControl.Groups.Add(this.group1);
            this.DocControl.Groups.Add(this.textToolsGroup);
            this.DocControl.Groups.Add(this.group2);
            this.DocControl.Groups.Add(this.acronymTableGroup);
            this.DocControl.Groups.Add(this.crossRefsGroup);
            this.DocControl.Label = "Doc Control";
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
            this.textToolsGroup.Label = "Style Tools";
            this.textToolsGroup.Name = "textToolsGroup";
            // 
            // group2
            // 
            this.group2.Items.Add(this.gallery2);
            this.group2.Items.Add(this.gallery1);
            this.group2.Items.Add(this.defaultUL);
            this.group2.Items.Add(this.defaultOL);
            this.group2.Label = "List Tools";
            this.group2.Name = "group2";
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
            this.headingsDropdown.Buttons.Add(this.headingOne);
            this.headingsDropdown.Buttons.Add(this.headingTwo);
            this.headingsDropdown.Buttons.Add(this.headingThree);
            this.headingsDropdown.Buttons.Add(this.headingFour);
            this.headingsDropdown.Buttons.Add(this.headingFive);
            this.headingsDropdown.ColumnCount = 1;
            this.headingsDropdown.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.headingsDropdown.Image = global::DocumentControlToolbar.Properties.Resources.headings_icon;
            this.headingsDropdown.Label = "Apply Headings";
            this.headingsDropdown.Name = "headingsDropdown";
            this.headingsDropdown.ShowImage = true;
            // 
            // headingOne
            // 
            this.headingOne.Label = "Level 1";
            this.headingOne.Name = "headingOne";
            this.headingOne.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.headingOne_Click);
            // 
            // headingTwo
            // 
            this.headingTwo.Label = "Level 2";
            this.headingTwo.Name = "headingTwo";
            this.headingTwo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.headingTwo_Click);
            // 
            // headingThree
            // 
            this.headingThree.Label = "Level 3";
            this.headingThree.Name = "headingThree";
            this.headingThree.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.headingThree_Click);
            // 
            // headingFour
            // 
            this.headingFour.Label = "Level 4";
            this.headingFour.Name = "headingFour";
            this.headingFour.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.headingFour_Click);
            // 
            // headingFive
            // 
            this.headingFive.Label = "Level 5";
            this.headingFive.Name = "headingFive";
            this.headingFive.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.headingFive_Click);
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
            // gallery2
            // 
            this.gallery2.Buttons.Add(this.levelOneU);
            this.gallery2.Buttons.Add(this.levelTwoU);
            this.gallery2.Buttons.Add(this.levelThreeU);
            this.gallery2.Buttons.Add(this.levelFourU);
            this.gallery2.ColumnCount = 1;
            this.gallery2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.gallery2.Image = global::DocumentControlToolbar.Properties.Resources.apply_list_large_icon;
            this.gallery2.Label = "Apply Unordered List";
            this.gallery2.Name = "gallery2";
            this.gallery2.RowCount = 1;
            this.gallery2.ShowImage = true;
            // 
            // levelOneU
            // 
            this.levelOneU.Label = "Level 1";
            this.levelOneU.Name = "levelOneU";
            this.levelOneU.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.levelOneU_Click);
            // 
            // levelTwoU
            // 
            this.levelTwoU.Label = "Level 2";
            this.levelTwoU.Name = "levelTwoU";
            this.levelTwoU.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.levelTwoU_Click);
            // 
            // levelThreeU
            // 
            this.levelThreeU.Label = "Level 3";
            this.levelThreeU.Name = "levelThreeU";
            this.levelThreeU.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.levelThreeU_Click);
            // 
            // levelFourU
            // 
            this.levelFourU.Label = "Level 4";
            this.levelFourU.Name = "levelFourU";
            this.levelFourU.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.levelFourU_Click);
            // 
            // gallery1
            // 
            this.gallery1.Buttons.Add(this.levelOneO);
            this.gallery1.Buttons.Add(this.levelTwoO);
            this.gallery1.ColumnCount = 1;
            this.gallery1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.gallery1.Image = global::DocumentControlToolbar.Properties.Resources.apply_o_list_large_icon;
            this.gallery1.Label = "Apply Orderd List";
            this.gallery1.Name = "gallery1";
            this.gallery1.ShowImage = true;
            // 
            // levelOneO
            // 
            this.levelOneO.Label = "Level 1";
            this.levelOneO.Name = "levelOneO";
            this.levelOneO.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.levelOneO_Click);
            // 
            // levelTwoO
            // 
            this.levelTwoO.Label = "Level 2";
            this.levelTwoO.Name = "levelTwoO";
            this.levelTwoO.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.levelTwoO_Click);
            // 
            // defaultUL
            // 
            this.defaultUL.Image = global::DocumentControlToolbar.Properties.Resources.apply_list_small_icon;
            this.defaultUL.Label = "Apply Default UL";
            this.defaultUL.Name = "defaultUL";
            this.defaultUL.ShowImage = true;
            this.defaultUL.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.defaultUL_Click);
            // 
            // defaultOL
            // 
            this.defaultOL.Image = global::DocumentControlToolbar.Properties.Resources.apply_o_list_small_icon;
            this.defaultOL.Label = "Apply Default OL";
            this.defaultOL.Name = "defaultOL";
            this.defaultOL.ShowImage = true;
            this.defaultOL.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.defaultOL_Click);
            // 
            // runAcronymTool
            // 
            this.runAcronymTool.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.runAcronymTool.Image = ((System.Drawing.Image)(resources.GetObject("runAcronymTool.Image")));
            this.runAcronymTool.Label = "Run Acronym Tool";
            this.runAcronymTool.Name = "runAcronymTool";
            this.runAcronymTool.ShowImage = true;
            this.runAcronymTool.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.runAcronymTool_Click);
            // 
            // updateWordlist
            // 
            this.updateWordlist.Image = global::DocumentControlToolbar.Properties.Resources.update_small_icon;
            this.updateWordlist.Label = " Update Wordlist ";
            this.updateWordlist.Name = "updateWordlist";
            this.updateWordlist.ShowImage = true;
            this.updateWordlist.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.updateWordlist_Click);
            // 
            // updateDudsList
            // 
            this.updateDudsList.Image = global::DocumentControlToolbar.Properties.Resources.update_small_icon;
            this.updateDudsList.Label = " Update Duds List ";
            this.updateDudsList.Name = "updateDudsList";
            this.updateDudsList.ShowImage = true;
            this.updateDudsList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.updateDudsList_Click);
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
            this.updateAllFields.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.updateAllFields_Click);
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
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.acronymTableGroup.ResumeLayout(false);
            this.acronymTableGroup.PerformLayout();
            this.crossRefsGroup.ResumeLayout(false);
            this.crossRefsGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton updateAllFields;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton defaultUL;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton defaultOL;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup acronymTableGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton updateWordlist;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton updateDudsList;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery headingsDropdown;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton headingOne;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton headingTwo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton headingThree;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton headingFour;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton headingFive;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery gallery1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton levelOneO;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton levelTwoO;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery gallery2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton levelOneU;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton levelTwoU;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton levelThreeU;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton levelFourU;
    }

    partial class ThisRibbonCollection
    {
        internal DocumentControlRibbon Ribbon1
        {
            get { return this.GetRibbon<DocumentControlRibbon>(); }
        }
    }
}
