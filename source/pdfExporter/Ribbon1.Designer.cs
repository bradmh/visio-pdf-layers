namespace pdfExporter
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.ExportCurrentPageButton = this.Factory.CreateRibbonButton();
            this.ExportAllPagesButton = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.InfoButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "PDF Exporter";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.ExportCurrentPageButton);
            this.group1.Items.Add(this.ExportAllPagesButton);
            this.group1.Label = "Export";
            this.group1.Name = "group1";
            // 
            // ExportCurrentPageButton
            // 
            this.ExportCurrentPageButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ExportCurrentPageButton.Image = global::pdfExporter.Properties.Resources.single;
            this.ExportCurrentPageButton.Label = "Current Page";
            this.ExportCurrentPageButton.Name = "ExportCurrentPageButton";
            this.ExportCurrentPageButton.ShowImage = true;
            this.ExportCurrentPageButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExportCurrentPageButton_Click);
            // 
            // ExportAllPagesButton
            // 
            this.ExportAllPagesButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ExportAllPagesButton.Image = global::pdfExporter.Properties.Resources.mulit;
            this.ExportAllPagesButton.Label = "All Pages";
            this.ExportAllPagesButton.Name = "ExportAllPagesButton";
            this.ExportAllPagesButton.ShowImage = true;
            this.ExportAllPagesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExportAllPagesButton_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.InfoButton);
            this.group2.Label = "About";
            this.group2.Name = "group2";
            // 
            // InfoButton
            // 
            this.InfoButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.InfoButton.Image = global::pdfExporter.Properties.Resources.about;
            this.InfoButton.Label = "About";
            this.InfoButton.Name = "InfoButton";
            this.InfoButton.ShowImage = true;
            this.InfoButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.InfoButton_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Visio.Drawing";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExportCurrentPageButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExportAllPagesButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton InfoButton;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
