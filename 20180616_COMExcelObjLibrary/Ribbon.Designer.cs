namespace _20180616_COMExcelObjLibrary
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
            this.rt1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btn1 = this.Factory.CreateRibbonButton();
            this.rt1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // rt1
            // 
            this.rt1.Groups.Add(this.group1);
            this.rt1.Label = "Ribbon1";
            this.rt1.Name = "rt1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btn1);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // btn1
            // 
            this.btn1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn1.Label = "Test";
            this.btn1.Name = "btn1";
            this.btn1.ShowImage = true;
            this.btn1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn1_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.rt1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.rt1.ResumeLayout(false);
            this.rt1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        public Microsoft.Office.Tools.Ribbon.RibbonTab rt1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
