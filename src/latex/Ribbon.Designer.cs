namespace latex
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            this.latexTab = this.Factory.CreateRibbonTab();
            this.tableGroup = this.Factory.CreateRibbonGroup();
            this.tableButton = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.clipboardBox = this.Factory.CreateRibbonCheckBox();
            this.fileBox = this.Factory.CreateRibbonCheckBox();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.latexTab.SuspendLayout();
            this.tableGroup.SuspendLayout();
            // 
            // latexTab
            // 
            this.latexTab.Groups.Add(this.tableGroup);
            this.latexTab.Label = "Latex";
            this.latexTab.Name = "latexTab";
            // 
            // tableGroup
            // 
            this.tableGroup.DialogLauncher = ribbonDialogLauncherImpl1;
            this.tableGroup.Items.Add(this.tableButton);
            this.tableGroup.Items.Add(this.separator1);
            this.tableGroup.Items.Add(this.label1);
            this.tableGroup.Items.Add(this.clipboardBox);
            this.tableGroup.Items.Add(this.fileBox);
            this.tableGroup.Label = "Table";
            this.tableGroup.Name = "tableGroup";
            this.tableGroup.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.tableGroup_DialogLauncherClick);
            // 
            // tableButton
            // 
            this.tableButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.tableButton.Enabled = false;
            this.tableButton.Image = global::latex.Properties.Resources.button;
            this.tableButton.Label = "ToTable";
            this.tableButton.Name = "tableButton";
            this.tableButton.ShowImage = true;
            this.tableButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.tableButton_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // clipboardBox
            // 
            this.clipboardBox.Checked = true;
            this.clipboardBox.Enabled = false;
            this.clipboardBox.Label = "Clipboard";
            this.clipboardBox.Name = "clipboardBox";
            // 
            // fileBox
            // 
            this.fileBox.Enabled = false;
            this.fileBox.Label = "File";
            this.fileBox.Name = "fileBox";
            // 
            // label1
            // 
            this.label1.Label = "Save to";
            this.label1.Name = "label1";
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.latexTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.latexTab.ResumeLayout(false);
            this.latexTab.PerformLayout();
            this.tableGroup.ResumeLayout(false);
            this.tableGroup.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab latexTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup tableGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton tableButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox clipboardBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox fileBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
