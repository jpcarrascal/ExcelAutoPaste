namespace ExcelAutoPaste
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.toggleReceive = this.Factory.CreateRibbonToggleButton();
            this.pasteDirection = this.Factory.CreateRibbonDropDown();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.toggleReceive);
            this.group1.Items.Add(this.pasteDirection);
            this.group1.Label = "Auto-Paste";
            this.group1.Name = "group1";
            // 
            // toggleReceive
            // 
            this.toggleReceive.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toggleReceive.Image = ((System.Drawing.Image)(resources.GetObject("toggleReceive.Image")));
            this.toggleReceive.Label = "Watch clipboard.";
            this.toggleReceive.Name = "toggleReceive";
            this.toggleReceive.ShowImage = true;
            this.toggleReceive.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ToggleReceive_Click);
            // 
            // pasteDirection
            // 
            ribbonDropDownItemImpl1.Label = "Column";
            ribbonDropDownItemImpl2.Label = "Row";
            this.pasteDirection.Items.Add(ribbonDropDownItemImpl1);
            this.pasteDirection.Items.Add(ribbonDropDownItemImpl2);
            this.pasteDirection.Label = "Paste direction";
            this.pasteDirection.Name = "pasteDirection";
            this.pasteDirection.ShowItemImage = false;
            this.pasteDirection.SizeString = "column";
            this.pasteDirection.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.pasteDirection_SelectionChanged);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleReceive;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown pasteDirection;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
