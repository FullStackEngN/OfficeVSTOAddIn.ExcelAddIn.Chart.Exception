namespace ExcelAddIn.Chart.Exception
{
    partial class RibbonChartException : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonChartException()
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
            this.groupException = this.Factory.CreateRibbonGroup();
            this.buttonAddChart = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupException.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupException);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // groupException
            // 
            this.groupException.Items.Add(this.buttonAddChart);
            this.groupException.Label = "Chart Exception Sample";
            this.groupException.Name = "groupException";
            // 
            // buttonAddChart
            // 
            this.buttonAddChart.Label = "Add Chart";
            this.buttonAddChart.Name = "buttonAddChart";
            this.buttonAddChart.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAddChart_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonChartException_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupException.ResumeLayout(false);
            this.groupException.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupException;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddChart;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonChartException Ribbon1
        {
            get { return this.GetRibbon<RibbonChartException>(); }
        }
    }
}
